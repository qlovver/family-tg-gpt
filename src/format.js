// Converts model output (Markdown + LaTeX) into safe Telegram HTML
// and splits long HTML into chunks without breaking tags.
//
// Pipeline: extract code -> convert LaTeX to Unicode -> escape HTML ->
// convert Markdown to markers/tags -> restore code -> markers to tags.
// Only tags produced here reach Telegram, so model output can never
// inject arbitrary HTML.

// Private-use characters as temporary markers (never appear in normal text).
const B_OPEN = "", B_CLOSE = "";
const I_OPEN = "", I_CLOSE = "";
const S_OPEN = "", S_CLOSE = "";
const CODE_SLOT = ""; // <n> placeholder for code
const CODE_END = "";

export function escapeHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

export function unescapeHtml(s) {
  return String(s)
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"')
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
    .replace(/&amp;/g, "&");
}

/* ------------------------------------------------------------------ */
/* LaTeX -> Unicode                                                    */
/* ------------------------------------------------------------------ */

const SYMBOLS = {
  alpha: "α", beta: "β", gamma: "γ", delta: "δ", epsilon: "ε", varepsilon: "ε",
  zeta: "ζ", eta: "η", theta: "θ", vartheta: "ϑ", iota: "ι", kappa: "κ",
  lambda: "λ", mu: "μ", nu: "ν", xi: "ξ", pi: "π", varpi: "ϖ", rho: "ρ",
  varrho: "ϱ", sigma: "σ", varsigma: "ς", tau: "τ", upsilon: "υ", phi: "φ",
  varphi: "φ", chi: "χ", psi: "ψ", omega: "ω",
  Gamma: "Γ", Delta: "Δ", Theta: "Θ", Lambda: "Λ", Xi: "Ξ", Pi: "Π",
  Sigma: "Σ", Upsilon: "Υ", Phi: "Φ", Psi: "Ψ", Omega: "Ω",
  cdot: "·", times: "×", div: "÷", pm: "±", mp: "∓", ast: "∗", star: "⋆",
  le: "≤", leq: "≤", leqslant: "≤", ge: "≥", geq: "≥", geqslant: "≥",
  ne: "≠", neq: "≠", approx: "≈", sim: "∼", simeq: "≃", cong: "≅",
  equiv: "≡", propto: "∝", ll: "≪", gg: "≫", lt: "<", gt: ">",
  infty: "∞", to: "→", rightarrow: "→", leftarrow: "←", longrightarrow: "⟶",
  Rightarrow: "⇒", Leftarrow: "⇐", Leftrightarrow: "⇔", leftrightarrow: "↔",
  implies: "⇒", iff: "⇔", mapsto: "↦", uparrow: "↑", downarrow: "↓",
  in: "∈", notin: "∉", ni: "∋", subset: "⊂", subseteq: "⊆", supset: "⊃",
  supseteq: "⊇", cup: "∪", cap: "∩", setminus: "∖", emptyset: "∅",
  varnothing: "∅", forall: "∀", exists: "∃", neg: "¬", lnot: "¬",
  land: "∧", wedge: "∧", lor: "∨", vee: "∨", oplus: "⊕", otimes: "⊗",
  partial: "∂", nabla: "∇", int: "∫", iint: "∬", iiint: "∭", oint: "∮",
  sum: "∑", prod: "∏", coprod: "∐", circ: "∘", degree: "°", angle: "∠",
  perp: "⊥", parallel: "∥", mid: "∣", ldots: "…", dots: "…", cdots: "⋯",
  vdots: "⋮", ddots: "⋱", prime: "′", hbar: "ℏ", ell: "ℓ", Re: "ℜ", Im: "ℑ",
  aleph: "ℵ", therefore: "∴", because: "∵", triangle: "△", square: "□",
  langle: "⟨", rangle: "⟩", lfloor: "⌊", rfloor: "⌋", lceil: "⌈", rceil: "⌉",
  vert: "|", Vert: "‖", lvert: "|", rvert: "|", lVert: "‖", rVert: "‖",
  backslash: "\\", percent: "%",
  quad: "  ", qquad: "    ", enspace: " ", thinspace: " ",
  displaystyle: "", textstyle: "", limits: "", nolimits: "", left: "", right: "",
  big: "", Big: "", bigg: "", Bigg: "", bigl: "", bigr: "", Bigl: "", Bigr: "",
  nonumber: "", notag: "", hline: "", centering: "", noindent: "",
};

// Function-like names rendered as plain words.
const FUNCS = new Set([
  "sin", "cos", "tan", "tg", "cot", "ctg", "sec", "csc", "arcsin", "arccos",
  "arctan", "arctg", "arcctg", "sinh", "cosh", "tanh", "sh", "ch", "th", "log",
  "lg", "ln", "exp", "lim", "max", "min", "sup", "inf", "det", "dim", "ker",
  "deg", "gcd", "arg", "mod", "bmod", "liminf", "limsup", "Pr",
]);

const BLACKBOARD = { R: "ℝ", N: "ℕ", Z: "ℤ", Q: "ℚ", C: "ℂ", P: "ℙ", E: "𝔼" };

const ACCENTS = {
  dot: "̇", ddot: "̈", hat: "̂", widehat: "̂",
  tilde: "̃", widetilde: "̃", bar: "̄", overline: "̅",
  vec: "⃗", overrightarrow: "⃗", acute: "́", grave: "̀",
  check: "̌", breve: "̆", underline: "̲",
};

const TEXT_CMDS = new Set([
  "text", "textrm", "textit", "textbf", "mathrm", "mathit", "mathbf", "mathsf",
  "mathtt", "mathcal", "mathscr", "mathfrak", "boldsymbol", "bm", "operatorname",
  "textnormal", "emph", "mbox", "hbox", "texttt", "textsf",
]);

const SUP = {
  "0": "⁰", "1": "¹", "2": "²", "3": "³", "4": "⁴", "5": "⁵", "6": "⁶", "7": "⁷",
  "8": "⁸", "9": "⁹", "+": "⁺", "-": "⁻", "−": "⁻", "=": "⁼", "(": "⁽", ")": "⁾",
  a: "ᵃ", b: "ᵇ", c: "ᶜ", d: "ᵈ", e: "ᵉ", f: "ᶠ", g: "ᵍ", h: "ʰ", i: "ⁱ",
  j: "ʲ", k: "ᵏ", l: "ˡ", m: "ᵐ", n: "ⁿ", o: "ᵒ", p: "ᵖ", r: "ʳ", s: "ˢ",
  t: "ᵗ", u: "ᵘ", v: "ᵛ", w: "ʷ", x: "ˣ", y: "ʸ", z: "ᶻ", T: "ᵀ", "*": "*",
  "′": "′", "°": "°", "∘": "°", " ": "",
};

const SUB = {
  "0": "₀", "1": "₁", "2": "₂", "3": "₃", "4": "₄", "5": "₅", "6": "₆", "7": "₇",
  "8": "₈", "9": "₉", "+": "₊", "-": "₋", "−": "₋", "=": "₌", "(": "₍", ")": "₎",
  a: "ₐ", e: "ₑ", h: "ₕ", i: "ᵢ", j: "ⱼ", k: "ₖ", l: "ₗ", m: "ₘ", n: "ₙ",
  o: "ₒ", p: "ₚ", r: "ᵣ", s: "ₛ", t: "ₜ", u: "ᵤ", v: "ᵥ", x: "ₓ", " ": "",
};

const VULGAR = { "1/2": "½", "1/3": "⅓", "2/3": "⅔", "1/4": "¼", "3/4": "¾", "1/5": "⅕", "1/8": "⅛" };

function mapScript(s, table) {
  let out = "";
  for (const ch of s) {
    if (!(ch in table)) return null;
    out += table[ch];
  }
  return out;
}

// "Simple" means it can be written without parentheses: x, 12, k̇, x², 0,5
function isSimple(s) {
  return /^[\p{L}\p{N}.,\p{M}²³¹⁰-₟′]+$/u.test(s) || /^\(.*\)$/.test(s) && balancedParens(s);
}

function balancedParens(s) {
  // true if the outer parens of s wrap the whole expression
  let depth = 0;
  for (let i = 0; i < s.length; i++) {
    if (s[i] === "(") depth++;
    else if (s[i] === ")") { depth--; if (depth === 0 && i < s.length - 1) return false; }
  }
  return depth === 0;
}

function wrap(s) {
  s = s.trim();
  return isSimple(s) ? s : `(${s})`;
}

function readGroup(s, i) {
  // s[i] === "{" ; returns [inner, nextIndex]
  let depth = 0;
  for (let j = i; j < s.length; j++) {
    if (s[j] === "\\") { j++; continue; }
    if (s[j] === "{") depth++;
    else if (s[j] === "}") { depth--; if (depth === 0) return [s.slice(i + 1, j), j + 1]; }
  }
  return [s.slice(i + 1), s.length];
}

function skipSpaces(s, i) {
  while (i < s.length && (s[i] === " " || s[i] === "\t")) i++;
  return i;
}

// Reads one argument: {group}, \command, or a single (code point) character.
function readArg(s, i) {
  i = skipSpaces(s, i);
  if (i >= s.length) return ["", i];
  if (s[i] === "{") return readGroup(s, i);
  if (s[i] === "\\") {
    const m = /^\\([A-Za-z]+|.)/.exec(s.slice(i));
    if (m) return [m[0], i + m[0].length];
  }
  const cp = String.fromCodePoint(s.codePointAt(i));
  return [cp, i + cp.length];
}

function readOptional(s, i) {
  i = skipSpaces(s, i);
  if (s[i] !== "[") return [null, i];
  const end = s.indexOf("]", i);
  if (end < 0) return [null, i];
  return [s.slice(i + 1, end), end + 1];
}

function applyAccent(arg, mark) {
  const chars = [...arg];
  if (!chars.length) return "";
  if (mark === "̅" || mark === "̲") return chars.map(c => c + mark).join("");
  // accent on last letter (k̇, x̂); for longer args put it on the last char
  return chars.join("") + mark;
}

function renderEnv(name, body, math) {
  const rows = body.split(/\\\\/).map(r => r.replace(/\\hline/g, "").trim()).filter(Boolean);
  const conv = rows.map(r => r.split("&").map(c => latexToUnicode(c, math).trim()));
  const base = name.replace(/\*$/, "");
  if (base === "cases" || base === "dcases") {
    const lines = conv.map(cells => cells.filter(Boolean).join(",  "));
    if (lines.length === 1) return `{ ${lines[0]}`;
    return lines.map((l, k) => (k === 0 ? "⎧ " : k === lines.length - 1 ? "⎩ " : "⎨ ") + l).join("\n");
  }
  const brackets = {
    pmatrix: ["(", ")"], bmatrix: ["[", "]"], Bmatrix: ["{", "}"],
    vmatrix: ["|", "|"], Vmatrix: ["‖", "‖"], matrix: ["", ""], smallmatrix: ["", ""],
  };
  if (base in brackets) {
    const [l, r] = brackets[base];
    const widths = [];
    conv.forEach(cells => cells.forEach((c, k) => { widths[k] = Math.max(widths[k] || 0, [...c].length); }));
    return conv.map(cells => {
      const inner = cells.map((c, k) => c + " ".repeat(widths[k] - [...c].length)).join("  ");
      return `${l} ${inner} ${r}`.trim();
    }).join("\n");
  }
  // align, aligned, gather, array, eqnarray, split ...
  return conv.map(cells => cells.join(" ").replace(/\s+/g, " ").trim()).join("\n");
}

/**
 * Converts LaTeX to readable Unicode.
 * math=true : full math mode (^, _, -, *, braces are interpreted).
 * math=false: only backslash commands are converted (safe for plain text).
 */
export function latexToUnicode(src, math = true) {
  const s = String(src);
  let out = "";
  let i = 0;
  while (i < s.length) {
    const ch = s[i];

    if (ch === "\\") {
      const rest = s.slice(i + 1);
      const word = /^[A-Za-z]+/.exec(rest);
      if (!word) {
        const next = rest[0];
        if (next === undefined) { i++; continue; }
        i += 2;
        if (next === "\\") { out += math ? "\n" : "\\"; continue; }
        if (",;:> ".includes(next)) { out += " "; continue; }
        if (next === "!") continue;
        if ("{}%$&#_".includes(next)) { out += next; continue; }
        if (next === "|") { out += "‖"; continue; }
        out += next;
        continue;
      }
      const name = word[0];
      let j = i + 1 + name.length;

      if (name === "frac" || name === "dfrac" || name === "tfrac" || name === "cfrac") {
        const [a, j1] = readArg(s, j);
        const [b, j2] = readArg(s, j1);
        const num = latexToUnicode(a, true).trim();
        const den = latexToUnicode(b, true).trim();
        out += VULGAR[`${num}/${den}`] || `${wrap(num)}/${wrap(den)}`;
        i = j2; continue;
      }
      if (name === "binom" || name === "dbinom" || name === "tbinom") {
        const [a, j1] = readArg(s, j);
        const [b, j2] = readArg(s, j1);
        out += `C(${latexToUnicode(a, true).trim()}, ${latexToUnicode(b, true).trim()})`;
        i = j2; continue;
      }
      if (name === "sqrt") {
        const [n, j1] = readOptional(s, j);
        const [a, j2] = readArg(s, j1);
        const inner = latexToUnicode(a, true).trim();
        const root = n === "3" ? "∛" : n === "4" ? "∜" : n ? `${mapScript(n, SUP) ?? `(${n})`}√` : "√";
        out += root + wrap(inner);
        i = j2; continue;
      }
      if (name in ACCENTS) {
        const [a, j1] = readArg(s, j);
        out += applyAccent(latexToUnicode(a, math).trim(), ACCENTS[name]);
        i = j1; continue;
      }
      if (name === "boxed" || name === "fbox" || name === "framebox") {
        const [a, j1] = readArg(s, j);
        out += B_OPEN + latexToUnicode(a, name === "boxed" ? true : math).trim() + B_CLOSE;
        i = j1; continue;
      }
      if (name === "mathbb") {
        const [a, j1] = readArg(s, j);
        out += [...a].map(c => BLACKBOARD[c] || c).join("");
        i = j1; continue;
      }
      if (TEXT_CMDS.has(name)) {
        const [a, j1] = readArg(s, j);
        // \text{...} contains normal text: only commands inside are converted
        const isText = name.startsWith("text") || name === "mbox" || name === "hbox" || name === "emph";
        out += latexToUnicode(a, isText ? false : math);
        i = j1; continue;
      }
      if (name === "color" || name === "textcolor") {
        const [, j1] = readArg(s, j);
        if (name === "textcolor") {
          const [a, j2] = readArg(s, j1);
          out += latexToUnicode(a, math); i = j2;
        } else i = j1;
        continue;
      }
      if (name === "underbrace" || name === "overbrace" || name === "phantom" ||
          name === "hphantom" || name === "vphantom" || name === "label" || name === "tag") {
        const [a, j1] = readArg(s, j);
        if (name === "underbrace" || name === "overbrace") out += latexToUnicode(a, math);
        if (name === "tag") out += `  (${a})`;
        i = j1; continue;
      }
      if (name === "begin") {
        const [env, j1] = readArg(s, j);
        const endTag = `\\end{${env}}`;
        let k = s.indexOf(endTag, j1);
        if (k < 0) k = s.length;
        let body = s.slice(j1, k);
        if (env === "array" && body.trimStart().startsWith("{")) {
          body = body.slice(readGroup(body, body.indexOf("{"))[1]); // drop column spec {cc|c}
        }
        out += renderEnv(env, body, true);
        i = k < s.length ? k + endTag.length : s.length;
        continue;
      }
      if (name === "end") { const [, j1] = readArg(s, j); i = j1; continue; }
      if (name === "left" || name === "right") {
        // \left( -> (, \left. -> nothing
        j = skipSpaces(s, j);
        if (s[j] === ".") { i = j + 1; continue; }
        i = j; continue;
      }
      if (FUNCS.has(name)) {
        out += name === "bmod" ? " mod " : name;
        // add a space before a following letter argument: \sin x -> sin x
        i = j; continue;
      }
      if (name in SYMBOLS) {
        out += SYMBOLS[name];
        i = j;
        // swallow a single space after word commands like "\cdot x"
        continue;
      }
      // Unknown command: keep the name without the backslash.
      out += name;
      i = j;
      continue;
    }

    if (math === "hybrid" && (ch === "^" || ch === "_")) {
      // Text that contains LaTeX: convert x_{t+1}, \int_0^1, a_n, 30^\circ,
      // but leave snake_case_names alone.
      const next = s[i + 1];
      const prevWord = /[A-Za-z0-9]*$/.exec(out)[0];
      const prevChar = out.slice(-1);
      const afterArg = readArg(s, i + 1)[1];
      const follow = s[afterArg] || " ";
      const braced = next === "{";
      const afterSymbol = prevChar && /[^\sA-Za-z0-9\p{L}]/u.test(prevChar) && /[^\s]/.test(prevChar);
      const singleLetter = prevWord.length === 1 && /[^A-Za-z0-9_]/.test(follow) &&
        !/[A-Za-z0-9_]/.test(out.slice(-2, -1) || " ");
      if (next === "\\" || braced || afterSymbol || singleLetter) {
        // fall through to math handling below
      } else { out += ch; i++; continue; }
    } else if (math === "hybrid" || !math) { out += ch; i++; continue; }

    if (ch === "^" || ch === "_") {
      const [a, j1] = readArg(s, i + 1);
      let arg = latexToUnicode(a, true).trim();
      if (ch === "^" && (arg === "\\circ" || arg === "∘" || arg === "°")) { out += "°"; i = j1; continue; }
      const mapped = mapScript(arg, ch === "^" ? SUP : SUB);
      if (mapped !== null && arg.length) out += mapped;
      else if ([...arg].length === 1) out += (ch === "^" ? "^" : "_") + arg;
      else out += `${ch}(${arg})`;
      i = j1; continue;
    }
    if (ch === "{") {
      const [inner, j1] = readGroup(s, i);
      out += latexToUnicode(inner, true);
      i = j1; continue;
    }
    if (ch === "}") { i++; continue; }
    if (ch === "~") { out += " "; i++; continue; }
    if (ch === "&") { out += " "; i++; continue; }
    if (ch === "-") { out += "−"; i++; continue; }
    if (ch === "*") { out += "·"; i++; continue; }
    if (ch === "'") { out += "′"; i++; continue; }
    out += ch;
    i++;
  }
  return out;
}

// Prettify spacing around operators produced from math.
function tidyMath(s) {
  return s
    .replace(/[ \t]+/g, " ")
    .replace(/ *([=≈≠≤≥<>⇒⇔→]) */g, " $1 ")
    .replace(/ +\n/g, "\n")
    .replace(/\n +/g, "\n")
    .trim();
}

function convertMathSegment(body) {
  return tidyMath(latexToUnicode(body, true));
}

// Lines like "\dot{k}=s\sqrt{k}-(n+g+\delta)k" written without $...$
function looksLikeBareFormula(line) {
  const t = line.trim();
  if (!t || /[А-Яа-яЁё]/.test(t)) return false;
  if (!/\\[A-Za-z]+/.test(t)) return false;
  // mostly symbols, few long English words
  const words = t.replace(/\\[A-Za-z]+/g, " ").match(/[A-Za-z]{4,}/g) || [];
  return words.length <= 2;
}

/**
 * Replaces all LaTeX in text (outside code) with Unicode.
 */
export function convertLatex(text) {
  let s = text;
  // display math
  s = s.replace(/\$\$([\s\S]+?)\$\$/g, (_, m) => "\n" + convertMathSegment(m) + "\n");
  s = s.replace(/\\\[([\s\S]+?)\\\]/g, (_, m) => "\n" + convertMathSegment(m) + "\n");
  // inline math
  s = s.replace(/\\\(([\s\S]+?)\\\)/g, (_, m) => convertMathSegment(m));
  s = s.replace(/(^|[^\\$\p{L}\p{N}])\$(?!\s)([^$\n]+?)(?<!\s)\$(?![\p{N}])/gu, (all, pre, m) => {
    // "$5 и $10" is money, not math: math rarely contains Russian words.
    if (/[А-Яа-яЁё]{2,}/.test(m.replace(/\\text\{[^}]*\}/g, ""))) return all;
    return pre + convertMathSegment(m);
  });
  // bare environments
  s = s.replace(/\\begin\{([a-zA-Z*]+)\}[\s\S]*?\\end\{\1\}/g, m => latexToUnicode(m, true));
  // bare formula lines and remaining commands
  s = s.split("\n").map(line => {
    if (looksLikeBareFormula(line)) {
      const indent = /^\s*/.exec(line)[0];
      return indent + convertMathSegment(line);
    }
    return /\\[A-Za-z]|[_^]\{/.test(line) ? latexToUnicode(line, "hybrid") : line;
  }).join("\n");
  // simple powers in plain text: x^2, 10^{-3}
  s = s.replace(/([\p{L}\p{N})\]])\^(\{([^{}\n]{1,12})\}|[\p{N}]{1,3}|[+\-−]?[\p{N}]{1,3})/gu, (m, base, raw, braced) => {
    const arg = braced !== undefined ? braced : raw;
    const mapped = mapScript(arg.replace(/-/g, "−"), SUP);
    return mapped ? base + mapped : m;
  });
  // collapse blank-line runs produced by display math
  return s.replace(/\n{3,}/g, "\n\n");
}

/* ------------------------------------------------------------------ */
/* Markdown -> Telegram HTML                                           */
/* ------------------------------------------------------------------ */

function inlineMarkdown(line) {
  let s = line;
  // links [text](url)
  s = s.replace(/\[([^\]\n]+)\]\((https?:\/\/[^\s)]+)\)/g,
    (_, t, u) => `<a href="${u.replace(/"/g, "&quot;")}">${t}</a>`);
  // bold **x** and __x__
  s = s.replace(/\*\*(?=\S)([^\n]*?\S)\*\*/g, `${B_OPEN}$1${B_CLOSE}`);
  s = s.replace(/(^|[^\p{L}\p{N}_])__(?=\S)([^\n]*?\S)__(?![\p{L}\p{N}_])/gu, `$1${B_OPEN}$2${B_CLOSE}`);
  // strikethrough
  s = s.replace(/~~(?=\S)([^\n]*?\S)~~/g, `${S_OPEN}$1${S_CLOSE}`);
  // italic *x* and _x_ (word-bounded to avoid a*b*c and snake_case)
  s = s.replace(/(^|[^\p{L}\p{N}*\\])\*(?=[^\s*])([^*\n]*?[^\s*])\*(?![\p{L}\p{N}*])/gu, `$1${I_OPEN}$2${I_CLOSE}`);
  s = s.replace(/(^|[^\p{L}\p{N}_\\])_(?=[^\s_])([^_\n]*?[^\s_])_(?![\p{L}\p{N}_])/gu, `$1${I_OPEN}$2${I_CLOSE}`);
  // leftover unmatched ** markers are noise
  s = s.replace(/\*\*/g, "");
  return s;
}

function renderTable(rows) {
  const cells = rows
    .filter(r => !/^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$/.test(r))
    .map(r => r.trim().replace(/^\|/, "").replace(/\|$/, "").split("|").map(c =>
      c.trim().replace(/\*\*/g, "").replace(/(^|\s)[*_](\S[^*_]*?)[*_](?=\s|$)/g, "$1$2")));
  const widths = [];
  for (const row of cells) row.forEach((c, k) => { widths[k] = Math.max(widths[k] || 0, [...c].length); });
  const lines = cells.map(row => row.map((c, k) => c + " ".repeat(widths[k] - [...c].length)).join(" │ ").trimEnd());
  return `<pre>${lines.join("\n")}</pre>`;
}

function blockMarkdown(text) {
  const lines = text.split("\n");
  const out = [];
  for (let k = 0; k < lines.length; k++) {
    let line = lines[k];

    // tables
    if (/^\s*\|.*\|\s*$/.test(line)) {
      const rows = [];
      while (k < lines.length && /^\s*\|.*\|\s*$/.test(lines[k])) rows.push(lines[k++]);
      k--;
      out.push(renderTable(rows));
      continue;
    }
    // blockquotes (already escaped: ">" is "&gt;")
    if (/^\s*&gt; ?/.test(line)) {
      const q = [];
      while (k < lines.length && /^\s*&gt; ?/.test(lines[k])) q.push(lines[k++].replace(/^\s*&gt; ?/, ""));
      k--;
      out.push(`<blockquote>${q.map(inlineMarkdown).join("\n")}</blockquote>`);
      continue;
    }
    // headings
    const h = /^\s{0,3}#{1,6}\s+(.+?)\s*#*\s*$/.exec(line);
    if (h) {
      const title = h[1].replace(/\*\*/g, "").replace(/__/g, "");
      out.push(`${B_OPEN}${inlineMarkdown(title)}${B_CLOSE}`);
      continue;
    }
    // horizontal rules
    if (/^\s*([-*_])(\s*\1){2,}\s*$/.test(line)) { out.push("──────────"); continue; }
    // bullet lists: "- x", "* x", "+ x" (keep indentation for nesting)
    const li = /^(\s*)[-*+•]\s+(.*)$/.exec(line);
    if (li) {
      const depth = Math.min(Math.floor(li[1].replace(/\t/g, "  ").length / 2), 3);
      out.push(`${"   ".repeat(depth)}${depth ? "◦" : "•"} ${inlineMarkdown(li[2])}`);
      continue;
    }
    // numbered lists "1) x" / "1. x": bold the number for readability
    const ol = /^(\s*)(\d{1,3})[.)]\s+(.*)$/.exec(line);
    if (ol) {
      out.push(`${ol[1]}${ol[2]}. ${inlineMarkdown(ol[3])}`);
      continue;
    }
    out.push(inlineMarkdown(line));
  }
  return out.join("\n");
}

// Converts markers to real tags in one pass. Guarantees proper nesting:
// duplicates are merged and crossing ranges are closed/re-opened.
function markersToTags(s) {
  const OPEN = { [B_OPEN]: "b", [I_OPEN]: "i", [S_OPEN]: "s" };
  const CLOSE = { [B_CLOSE]: "b", [I_CLOSE]: "i", [S_CLOSE]: "s" };
  const stack = [];
  let res = "";
  for (const ch of s) {
    if (ch in OPEN) {
      const tag = OPEN[ch];
      if (!stack.includes(tag)) { stack.push(tag); res += `<${tag}>`; }
    } else if (ch in CLOSE) {
      const tag = CLOSE[ch];
      const idx = stack.lastIndexOf(tag);
      if (idx < 0) continue;
      const inner = stack.splice(idx);
      res += inner.slice().reverse().map(t => `</${t}>`).join("");
      for (const t of inner.slice(1)) { stack.push(t); res += `<${t}>`; }
    } else res += ch;
  }
  res += stack.reverse().map(t => `</${t}>`).join("");
  return res.replace(/<(b|i|s)><\/\1>/g, "");
}

/**
 * Main entry: model text -> Telegram HTML (parse_mode: "HTML").
 */
export function toTelegramHtml(input) {
  let text = String(input ?? "").replace(/\r\n?/g, "\n").replace(/[-]/g, "");
  const codes = [];
  const slot = html => { codes.push(html); return `${CODE_SLOT}${codes.length - 1}${CODE_END}`; };

  // fenced code blocks
  text = text.replace(/```([\w+-]*)[^\n]*\n([\s\S]*?)(?:```|$)/g, (_, lang, code) => {
    const body = escapeHtml(code.replace(/\n$/, ""));
    const cls = lang ? ` class="language-${lang.replace(/[^\w+-]/g, "")}"` : "";
    return slot(`<pre><code${cls}>${body}</code></pre>`);
  });
  // inline code
  text = text.replace(/`([^`\n]+)`/g, (_, code) => {
    // `\frac{1}{2}` style "code" from the model is usually math, not code
    if (/\\[A-Za-z]+/.test(code) && !/[;(){}]\s*$/.test(code) && !/[=]>|function|return/.test(code)) {
      return convertMathSegment(code);
    }
    return slot(`<code>${escapeHtml(code)}</code>`);
  });

  text = convertLatex(text);
  text = escapeHtml(text);
  text = blockMarkdown(text);
  text = markersToTags(text);
  text = text.replace(new RegExp(`${CODE_SLOT}(\\d+)${CODE_END}`, "g"), (_, n) => codes[Number(n)]);
  // tidy whitespace: no trailing spaces, max one blank line
  text = text.replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
  return text;
}

/**
 * Plain-text fallback from our HTML (used if Telegram rejects the markup).
 */
export function htmlToPlain(html) {
  return unescapeHtml(String(html).replace(/<a href="([^"]*)">([\s\S]*?)<\/a>/g, "$2 ($1)").replace(/<[^>]+>/g, ""));
}

/* ------------------------------------------------------------------ */
/* Splitting                                                           */
/* ------------------------------------------------------------------ */

// Visible length as Telegram counts it (UTF-16 units of text after parsing).
export function visibleLength(html) {
  return htmlToPlain(html).length;
}

function tokenize(html) {
  const tokens = [];
  const re = /<\/?[a-zA-Z][^>]*>|&(?:[a-z]+|#\d+);|[\s\S]/gu;
  let m;
  while ((m = re.exec(html))) {
    const t = m[0];
    if (t[0] === "<" && t.length > 1) {
      const close = t[1] === "/";
      const name = /^<\/?([a-zA-Z-]+)/.exec(t)[1].toLowerCase();
      tokens.push({ kind: close ? "close" : "open", name, raw: t, len: 0 });
    } else {
      tokens.push({ kind: "text", raw: t, len: t[0] === "&" && t.length > 1 ? 1 : t.length });
    }
  }
  return tokens;
}

/**
 * Splits Telegram HTML into chunks whose visible length is <= limit.
 * Prefers paragraph, then line, then word boundaries. Open tags are closed
 * at the end of a chunk and re-opened at the start of the next one.
 */
export function splitHtml(html, limit = 3900) {
  if (visibleLength(html) <= limit) return html ? [html] : [];
  const tokens = tokenize(html);
  const chunks = [];
  let start = 0;
  let carryOpen = []; // tags open at the start of the current chunk

  while (start < tokens.length) {
    const stack = [...carryOpen];
    let len = 0;
    let best = { para: -1, line: -1, word: -1 };
    let stacks = {};
    let i = start;
    for (; i < tokens.length; i++) {
      const t = tokens[i];
      if (t.kind === "open") { stack.push(t); continue; }
      if (t.kind === "close") {
        const idx = stack.map(x => x.name).lastIndexOf(t.name);
        if (idx >= 0) stack.splice(idx, 1);
        continue;
      }
      if (len + t.len > limit) break;
      len += t.len;
      if (t.raw === "\n") {
        const prevNl = i > start && tokens[i - 1].raw === "\n";
        if (prevNl) { best.para = i; stacks.para = [...stack]; }
        best.line = i; stacks.line = [...stack];
      } else if (t.raw === " ") {
        best.word = i; stacks.word = [...stack];
      }
    }
    if (i >= tokens.length) {
      chunks.push({ open: carryOpen, tokens: tokens.slice(start), close: stack });
      break;
    }
    // choose the best break that keeps the chunk reasonably full
    const minFill = Math.floor((i - start) * 0.5);
    let cut, cutStack;
    for (const kind of ["para", "line", "word"]) {
      if (best[kind] >= 0 && best[kind] - start >= minFill) { cut = best[kind]; cutStack = stacks[kind]; break; }
    }
    if (cut === undefined) {
      for (const kind of ["para", "line", "word"]) {
        if (best[kind] > start) { cut = best[kind]; cutStack = stacks[kind]; break; }
      }
    }
    if (cut === undefined) { cut = i - 1; cutStack = [...stack]; }
    // when cutting hard (no whitespace) recompute the stack at the cut
    chunks.push({ open: carryOpen, tokens: tokens.slice(start, cut + 1), close: cutStack });
    carryOpen = cutStack;
    start = cut + 1;
  }

  return chunks.map(({ open, tokens: toks, close }) => {
    const head = open.map(t => t.raw).join("");
    const body = toks.map(t => t.raw).join("");
    const tail = [...close].reverse().map(t => `</${t.name}>`).join("");
    return (head + body + tail).replace(/^\s+|\s+$/g, "")
      // drop empty tag pairs created by the split
      .replace(/<(b|i|s|u|code|pre|blockquote)>\s*<\/\1>/g, "");
  }).filter(c => htmlToPlain(c).trim());
}
