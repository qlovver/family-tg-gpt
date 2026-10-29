// Colour themes shared by Word, Excel and PowerPoint, plus custom overrides
// from the request: primary/accent/bg/text/heading colours and fonts,
// as hex ("#C62828") or colour names ("красный", "navy").


export const THEMES = {
  ocean:    { dark: "0B2545", primary: "13315C", accent: "2EC4B6", accent2: "8DA9C4", bg: "F6F9FC", card: "FFFFFF", text: "1B2A41", heading: "13315C", muted: "5C6B7A", onDark: "FFFFFF", onDarkMuted: "B9CCE0", line: "D9E2EC", font: "Calibri", headFont: "Calibri" },
  sunset:   { dark: "3A1F33", primary: "B33F62", accent: "F9A03F", accent2: "F4D35E", bg: "FFF8F1", card: "FFFFFF", text: "2D1E2F", heading: "8E2C4E", muted: "7A6670", onDark: "FFFFFF", onDarkMuted: "F2C9B8", line: "F0DCCF", font: "Calibri", headFont: "Calibri" },
  forest:   { dark: "17332A", primary: "2F6B4F", accent: "D4A72C", accent2: "8FBC94", bg: "F5F8F2", card: "FFFFFF", text: "1E2B24", heading: "245640", muted: "5F6F66", onDark: "FFFFFF", onDarkMuted: "C2D8C9", line: "DCE6DA", font: "Calibri", headFont: "Calibri" },
  minimal:  { dark: "111111", primary: "1F1F1F", accent: "E63946", accent2: "A8A8A8", bg: "FFFFFF", card: "F5F5F5", text: "1F1F1F", heading: "111111", muted: "6B6B6B", onDark: "FFFFFF", onDarkMuted: "C9C9C9", line: "E5E5E5", font: "Calibri", headFont: "Calibri" },
  lavender: { dark: "2E1A47", primary: "5A3E85", accent: "F28482", accent2: "B8A1D9", bg: "F8F5FC", card: "FFFFFF", text: "2A2140", heading: "4A2F75", muted: "6E6483", onDark: "FFFFFF", onDarkMuted: "D3C4EA", line: "E6DDF2", font: "Calibri", headFont: "Calibri" },
  night:    { dark: "0B1120", primary: "1E293B", accent: "38BDF8", accent2: "A78BFA", bg: "0F172A", card: "1E293B", text: "E2E8F0", heading: "F8FAFC", muted: "94A3B8", onDark: "F8FAFC", onDarkMuted: "94A3B8", line: "334155", font: "Calibri", headFont: "Calibri", darkContent: true },
  classic:  { dark: "1C2541", primary: "1C2541", accent: "C9A227", accent2: "8C9AB8", bg: "FBF9F4", card: "FFFFFF", text: "1C2541", heading: "1C2541", muted: "6A6F7E", onDark: "FFFFFF", onDarkMuted: "D8CFAF", line: "E7E1D2", font: "Calibri", headFont: "Georgia" },
};
const THEME_ALIASES = {
  blue: "ocean", синяя: "ocean", синий: "ocean", океан: "ocean", business: "ocean", деловая: "ocean",
  warm: "sunset", тёплая: "sunset", теплая: "sunset", закат: "sunset", red: "sunset", orange: "sunset",
  green: "forest", зелёная: "forest", зеленая: "forest", природа: "forest", eco: "forest",
  white: "minimal", минимализм: "minimal", строгая: "minimal", black: "minimal",
  purple: "lavender", фиолетовая: "lavender", лавандовая: "lavender", pink: "lavender", нежная: "lavender",
  dark: "night", тёмная: "night", темная: "night", ночь: "night", tech: "night",
  gold: "classic", классика: "classic", классическая: "classic", history: "classic", история: "classic",
};

export function pickTheme(name) {
  const n = String(name || "").toLowerCase().trim();
  return THEMES[n] ? n : THEME_ALIASES[n] || "ocean";
}


export function tint(hex, amount) {
  // mix with white; amount = share of the colour (1 = unchanged)
  const n = parseInt(hex, 16);
  const ch = s => Math.round(255 - (255 - ((n >> s) & 255)) * amount);
  return [16, 8, 0].map(s => ch(s).toString(16).padStart(2, "0")).join("").toUpperCase();
}

export function shade(hex, amount) {
  // mix with black; amount = share of the colour
  const n = parseInt(hex, 16);
  const ch = s => Math.round(((n >> s) & 255) * amount);
  return [16, 8, 0].map(s => ch(s).toString(16).padStart(2, "0")).join("").toUpperCase();
}

export function luminance(hex) {
  const n = parseInt(hex, 16);
  const [r, g, b] = [16, 8, 0].map(s => ((n >> s) & 255) / 255);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

const COLOR_NAMES = {
  красный: "C62828", red: "C62828", алый: "E53935", бордовый: "7F1D1D", "тёмно-красный": "8E1B1B", burgundy: "7F1D1D", maroon: "7F1D1D",
  оранжевый: "EF6C00", orange: "EF6C00", персиковый: "FFB38A", peach: "FFB38A", коралловый: "F28482", coral: "F28482",
  жёлтый: "F9A825", желтый: "F9A825", yellow: "F9A825", золотой: "C9A227", gold: "C9A227", горчичный: "C8A02B",
  зелёный: "2E7D32", зеленый: "2E7D32", green: "2E7D32", салатовый: "7CB342", lime: "7CB342", оливковый: "6B7B2A", olive: "6B7B2A",
  мятный: "3FBF9F", mint: "3FBF9F", изумрудный: "047857", emerald: "047857", бирюзовый: "00897B", teal: "00897B", turquoise: "00A3A3",
  голубой: "0288D1", "light blue": "0288D1", sky: "0EA5E9", синий: "1565C0", blue: "1565C0", "тёмно-синий": "0D2A5C", "темно-синий": "0D2A5C", navy: "0D2A5C",
  фиолетовый: "6A1B9A", purple: "6A1B9A", violet: "7C3AED", сиреневый: "9575CD", лавандовый: "B39DDB", lavender: "B39DDB",
  розовый: "D81B60", pink: "D81B60", малиновый: "C2185B", magenta: "C2185B", пудровый: "E8B4B8",
  коричневый: "6D4C41", brown: "6D4C41", бежевый: "D7C4A3", beige: "D7C4A3", песочный: "E2C799",
  серый: "616161", gray: "616161", grey: "616161", графитовый: "374151", graphite: "374151", серебряный: "A8A8A8", silver: "A8A8A8",
  чёрный: "111111", черный: "111111", black: "111111", белый: "FFFFFF", white: "FFFFFF", кремовый: "FBF6EA", cream: "FBF6EA",
};

export function parseColor(v) {
  if (v == null) return null;
  const s = String(v).trim().toLowerCase().replace(/ё/g, "ё");
  let m;
  if ((m = /^#?([0-9a-f]{6})$/i.exec(s))) return m[1].toUpperCase();
  if ((m = /^#?([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(s))) return (m[1] + m[1] + m[2] + m[2] + m[3] + m[3]).toUpperCase();
  return COLOR_NAMES[s] || COLOR_NAMES[s.replace(/ая$|ое$|ые$|ую$/, "ый")] || null;
}

const FONT_RE = /^[A-Za-zА-Яа-яЁё0-9 .\-]{2,40}$/;

/**
 * attrs: { theme, primary|color, accent, bg|background, text, heading, font, headfont }
 * Returns a full theme object (same keys as THEMES entries).
 */
export function resolveTheme(attrs = {}) {
  const a = Object.fromEntries(Object.entries(attrs || {}).map(([k, v]) => [String(k).toLowerCase().replace(/[-_]/g, ""), v]));
  const t = { ...THEMES[pickTheme(a.theme)] };
  const primary = parseColor(a.primary ?? a.color ?? a.main);
  const accent = parseColor(a.accent ?? a.accentcolor ?? a.secondary);
  const bg = parseColor(a.bg ?? a.background ?? a.fon);
  const text = parseColor(a.text ?? a.textcolor);
  const heading = parseColor(a.heading ?? a.headingcolor ?? a.titlecolor ?? a.headings);
  if (primary) {
    Object.assign(t, { primary, dark: luminance(primary) < 0.18 ? primary : shade(primary, 0.5), accent2: tint(primary, 0.5), onDarkMuted: tint(primary, 0.35) });
    if (!t.darkContent) t.heading = luminance(primary) > 0.6 ? shade(primary, 0.6) : primary;
    t.line = tint(primary, 0.18);
  }
  if (accent) t.accent = accent;
  if (bg) {
    t.bg = bg;
    if (luminance(bg) < 0.35) {
      Object.assign(t, { darkContent: true, text: "E5E7EB", heading: "FFFFFF", muted: "9CA3AF", card: shadeUp(bg), line: shadeUp(shadeUp(bg)), dark: shade(bg, 0.7) });
    } else {
      Object.assign(t, { darkContent: false, card: "FFFFFF" });
      if (THEMES[pickTheme(a.theme)].darkContent) Object.assign(t, { text: "1F2937", heading: primary || "1F2937", muted: "6B7280", line: "E5E7EB" });
    }
  }
  if (text) t.text = text;
  if (heading) t.heading = heading;
  if (a.font && FONT_RE.test(a.font)) t.font = String(a.font).trim();
  const hf = a.headfont ?? a.headingfont ?? a.titlefont;
  if (hf && FONT_RE.test(hf)) t.headFont = String(hf).trim();
  else if (a.font && FONT_RE.test(a.font)) t.headFont = String(a.font).trim();
  return t;
}

// lighten a dark colour a little (for cards on dark backgrounds)
function shadeUp(hex) {
  const n = parseInt(hex, 16);
  const ch = s => Math.min(255, ((n >> s) & 255) + 22);
  return [16, 8, 0].map(s => ch(s).toString(16).padStart(2, "0")).join("").toUpperCase();
}
