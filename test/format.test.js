import { test } from "node:test";
import assert from "node:assert/strict";
import { toTelegramHtml, splitHtml, htmlToPlain, latexToUnicode, visibleLength } from "../src/format.js";

const ALLOWED_TAGS = new Set(["b", "i", "s", "u", "code", "pre", "a", "blockquote"]);

function assertValidHtml(html) {
  const stack = [];
  for (const m of html.matchAll(/<(\/?)([a-z]+)[^>]*>/g)) {
    const [, close, name] = m;
    assert.ok(ALLOWED_TAGS.has(name), `tag <${name}> not allowed in: ${html}`);
    if (close) assert.equal(stack.pop(), name, `bad nesting in: ${html}`);
    else stack.push(name);
  }
  assert.equal(stack.length, 0, `unclosed tags ${stack} in: ${html}`);
  // no raw < or & outside tags/entities
  const text = html.replace(/<\/?[a-z]+[^>]*>/g, "");
  assert.ok(!/<|>/.test(text), "unescaped angle bracket");
  assert.ok(!/&(?!amp;|lt;|gt;|quot;)/.test(text), "unescaped ampersand");
}

test("example from the task renders without Markdown or LaTeX residue", () => {
  const src = String.raw`### Задача 4. Изменение капитала

**1. Формула:**
\dot{k}=s\sqrt{k}-(n+g+\delta)k

**2. Подставляем:**
$\dot{k} = 0{,}5 \cdot \sqrt{36} - 0{,}1 \cdot 36$

$$\dot{k} = 3 - 3{,}6 = \boxed{-0{,}6}$$`;
  const html = toTelegramHtml(src);
  assertValidHtml(html);
  assert.match(html, /^<b>Задача 4\. Изменение капитала<\/b>/);
  assert.match(html, /<b>1\. Формула:<\/b>/);
  assert.match(html, /k̇ = s√k−\(n\+g\+δ\)k/);
  assert.match(html, /k̇ = 0,5 · √36 − 0,1 · 36/);
  assert.match(html, /<b>−0,6<\/b>/);
  for (const bad of ["###", "**", "\\", "$", "{", "}"]) assert.ok(!html.includes(bad), `residue ${bad}: ${html}`);
});

test("LaTeX commands to Unicode", () => {
  const cases = [
    [String.raw`\frac{1}{2}`, "½"],
    [String.raw`\frac{x+1}{x-1}`, "(x+1)/(x−1)"],
    [String.raw`\frac{a}{b}`, "a/b"],
    [String.raw`x^2 + y^{10}`, "x² + y¹⁰"],
    [String.raw`x_{t+1}`, "xₜ₊₁"],
    [String.raw`\sqrt[3]{8}`, "∛8"],
    [String.raw`\sqrt{a+b}`, "√(a+b)"],
    [String.raw`30^\circ`, "30°"],
    [String.raw`a \leq b \neq c`, "a ≤ b ≠ c"],
    [String.raw`\text{если } x \in \mathbb{R}`, "если  x ∈ ℝ"],
    [String.raw`\lim_{n \to \infty}`, "lim_(n → ∞)"],
    [String.raw`\left( \frac{1}{4} \right)`, "( ¼ )"],
  ];
  for (const [src, want] of cases) assert.equal(latexToUnicode(src, true).replace(/\s+/g, " ").trim(), want.replace(/\s+/g, " "));
});

test("cases and matrices become readable lines", () => {
  const html = toTelegramHtml(String.raw`$$\begin{cases} x+y=1 \\ x-y=3 \end{cases}$$
\begin{pmatrix} 1 & 2 \\ 3 & 4 \end{pmatrix}`);
  assert.match(html, /⎧ x\+y = 1\n⎩ x−y = 3/);
  assert.match(html, /\( 1  2 \)\n\( 3  4 \)/);
});

test("HTML special characters are escaped, model cannot inject tags", () => {
  const html = toTelegramHtml(`a < b & c > d <script>alert(1)</script> <b>x</b>`);
  assertValidHtml(html);
  assert.ok(html.includes("&lt;script&gt;"));
  assert.ok(html.includes("&lt;b&gt;x&lt;/b&gt;"));
});

test("code is preserved verbatim and escaped", () => {
  const html = toTelegramHtml("Пример:\n```python\nif a < b and x**2:\n    print('**')\n```\nи `a_b*c*`");
  assertValidHtml(html);
  assert.match(html, /<pre><code class="language-python">if a &lt; b and x\*\*2:\n    print\('\*\*'\)<\/code><\/pre>/);
  assert.match(html, /<code>a_b\*c\*<\/code>/);
});

test("lists, italics, links, quotes, tables; snake_case and 2*3*4 untouched", () => {
  const html = toTelegramHtml(`- первый
* второй
  - вложенный
1. шаг
*курсив* и _тоже_, snake_case_name, 2*3*4
[сайт](https://example.com/?a=1&b=2)
> цитата
| a | b |
|---|---|
| 1 | 2 |`);
  assertValidHtml(html);
  assert.match(html, /^• первый\n• второй\n   ◦ вложенный\n1\. шаг/);
  assert.match(html, /<i>курсив<\/i> и <i>тоже<\/i>, snake_case_name, 2\*3\*4/);
  assert.match(html, /<a href="https:\/\/example\.com\/\?a=1&amp;b=2">сайт<\/a>/);
  assert.match(html, /<blockquote>цитата<\/blockquote>/);
  assert.match(html, /<pre>a │ b\n1 │ 2<\/pre>/);
});

test("unbalanced markdown never produces unbalanced tags", () => {
  for (const s of ["**жирный без конца", "*a", "__x", "~~y", "### **заголовок", "**a *b** c*", "$x^2", "\\frac{1}{"]) {
    assertValidHtml(toTelegramHtml(s));
  }
});

test("dollar amounts are not mangled", () => {
  const html = toTelegramHtml("Цена $5 и $10, итого 15$");
  assert.equal(html, "Цена $5 и $10, итого 15$");
});

test("splitHtml keeps every chunk under the limit and loses nothing", () => {
  const md = Array.from({ length: 400 }, (_, i) => `**Шаг ${i}.** Вычисляем $x^2 + ${i}$ и *проверяем* ответ.`).join("\n\n");
  const html = toTelegramHtml(md);
  const parts = splitHtml(html, 1000);
  assert.ok(parts.length > 5);
  for (const p of parts) { assert.ok(visibleLength(p) <= 1000, `chunk too long: ${visibleLength(p)}`); assertValidHtml(p); }
  const joined = parts.map(htmlToPlain).join("\n\n");
  assert.equal(joined.replace(/\s+/g, " "), htmlToPlain(html).replace(/\s+/g, " "));
});

test("splitHtml reopens tags cut in the middle (long bold, long code)", () => {
  const html = toTelegramHtml("**" + "слово ".repeat(500).trim() + "**\n```\n" + "line\n".repeat(400) + "```");
  const parts = splitHtml(html, 700);
  for (const p of parts) { assertValidHtml(p); assert.ok(visibleLength(p) <= 700); }
  assert.ok(parts[1].startsWith("<b>"));
  assert.ok(parts.some(p => p.startsWith("<pre><code>")));
  assert.equal(parts.map(htmlToPlain).join("").replace(/\s+/g, ""), htmlToPlain(html).replace(/\s+/g, ""));
});

test("splitHtml handles text without spaces and emoji (surrogate pairs)", () => {
  const html = "😀".repeat(3000);
  const parts = splitHtml(html, 1000);
  for (const p of parts) assert.ok(p.length <= 1000);
  assert.equal(parts.join(""), html);
  assert.ok(!parts.some(p => /^[\uDC00-\uDFFF]/.test(p)), "split inside a surrogate pair");
});

test("entities count as one visible character", () => {
  assert.equal(visibleLength("&lt;&amp;&gt;"), 3);
  assert.equal(htmlToPlain("<b>a &lt; b</b>"), "a < b");
});
