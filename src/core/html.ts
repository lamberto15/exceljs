// Simple HTML → RichText (rich-text editors, CMS fields): bold / italic / underline / strike,
// colors, line breaks. Not an HTML renderer: unknown tags are dropped, their text kept.

import type { Color, Font, RichText, RichTextRun } from './types.js';

const ENT: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' };

function decode(s: string): string {
  return s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (m, e: string) => {
    if (e[0] === '#') {
      const n = e[1] === 'x' || e[1] === 'X' ? Number.parseInt(e.slice(2), 16) : Number(e.slice(1));
      return Number.isFinite(n) && n > 0 && n < 0x110000 ? String.fromCodePoint(n) : m;
    }
    return ENT[e.toLowerCase()] ?? m;
  });
}

/** CSS / HTML color → ARGB ("#f00", "#ff0000", "rgb(255,0,0)", a few names). */
function color(v: string): Color | undefined {
  const s = v.trim().toLowerCase();
  const named: Record<string, string> = {
    black: '000000',
    white: 'FFFFFF',
    red: 'FF0000',
    green: '008000',
    blue: '0000FF',
    yellow: 'FFFF00',
    orange: 'FFA500',
    gray: '808080',
    grey: '808080',
    purple: '800080',
  };
  let hex = named[s];
  if (!hex) {
    const h = /^#([0-9a-f]{3}|[0-9a-f]{6})$/.exec(s);
    if (h) hex = h[1]!.length === 3 ? [...h[1]!].map((x) => x + x).join('') : h[1]!;
    const r = /^rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/.exec(s);
    if (r)
      hex = [r[1], r[2], r[3]]
        .map((x) => Math.min(255, +x!).toString(16).padStart(2, '0'))
        .join('');
  }
  return hex ? { argb: `FF${hex.toUpperCase()}` } : undefined;
}

/** Font changes of an opening tag (undefined: no formatting). */
function tagFont(name: string, attrs: string): Font | undefined {
  const f: { -readonly [K in keyof Font]: Font[K] } = {};
  if (name === 'b' || name === 'strong') f.bold = true;
  else if (name === 'i' || name === 'em') f.italic = true;
  else if (name === 'u' || name === 'ins') f.underline = 'single';
  else if (name === 's' || name === 'strike' || name === 'del') f.strike = true;
  else if (name === 'sup') f.vertAlign = 'superscript';
  else if (name === 'sub') f.vertAlign = 'subscript';
  const fc = /\bcolor\s*=\s*["']?([^"'\s>]+)/i.exec(attrs);
  if (name === 'font' && fc) f.color = color(fc[1]!);
  const st = /\bstyle\s*=\s*"([^"]*)"|\bstyle\s*=\s*'([^']*)'/i.exec(attrs);
  if (st) {
    for (const decl of (st[1] ?? st[2] ?? '').split(';')) {
      const k = decl.indexOf(':');
      if (k < 0) continue;
      const p = decl.slice(0, k).trim().toLowerCase();
      const v = decl
        .slice(k + 1)
        .trim()
        .toLowerCase();
      if (p === 'color') f.color = color(v);
      else if (p === 'font-weight' && (v === 'bold' || +v >= 600)) f.bold = true;
      else if (p === 'font-style' && v === 'italic') f.italic = true;
      else if (p === 'text-decoration' || p === 'text-decoration-line') {
        if (v.includes('underline')) f.underline = 'single';
        if (v.includes('line-through')) f.strike = true;
      }
    }
  }
  for (const k in f) if (f[k as keyof Font] === undefined) delete f[k as keyof Font];
  return Object.keys(f).length ? f : undefined;
}

const BLOCK = new Set(['p', 'div', 'li', 'tr', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6']);

/**
 * Convert simple HTML to RichText: `<b>`/`<strong>`, `<i>`/`<em>`, `<u>`, `<s>`/`<del>`,
 * `<sup>`/`<sub>`, `<font color>`, `<span style="color / font-weight / font-style /
 * text-decoration">`, `<br>` and block tags (p, div, li, ...) as line breaks, entities decoded.
 * Other tags are dropped (their text kept). Store the result in a cell; text with line breaks
 * wraps automatically.
 */
export function richTextFromHtml(html: string): RichText {
  const runs: RichTextRun[] = [];
  const stack: { name: string; font: Font | undefined }[] = [];
  const current = (): Font | undefined => {
    let f: Font | undefined;
    for (const s of stack) if (s.font) f = { ...f, ...s.font };
    return f;
  };
  const push = (text: string): void => {
    if (!text) return;
    const font = current();
    const last = runs[runs.length - 1];
    if (last && JSON.stringify(last.font) === JSON.stringify(font)) {
      runs[runs.length - 1] = font ? { text: last.text + text, font } : { text: last.text + text };
    } else runs.push(font ? { text, font } : { text });
  };
  const re = /<(\/?)([a-z][a-z0-9]*)([^>]*)>|<!--[\s\S]*?-->/gi;
  let last = 0;
  let started = false;
  for (let m = re.exec(html); m; m = re.exec(html)) {
    push(decode(html.slice(last, m.index).replace(/\s+/g, ' ')));
    last = re.lastIndex;
    if (!m[2]) continue; // comment
    const name = m[2].toLowerCase();
    const closing = m[1] === '/';
    if (name === 'br') push('\n');
    else if (BLOCK.has(name)) {
      if (!closing && started) push('\n');
    } else if (!closing) {
      if (!/\/\s*$/.test(m[3]!)) stack.push({ name, font: tagFont(name, m[3]!) });
    } else {
      const k = stack.map((s) => s.name).lastIndexOf(name);
      if (k >= 0) stack.splice(k, 1);
    }
    if (runs.length) started = true;
  }
  push(decode(html.slice(last).replace(/\s+/g, ' ')));
  // trim the outer whitespace
  if (runs.length) {
    const a = runs[0]!;
    runs[0] = { ...a, text: a.text.replace(/^[ \n]+/, '') };
    const z = runs[runs.length - 1]!;
    runs[runs.length - 1] = { ...z, text: z.text.replace(/[ \n]+$/, '') };
  }
  return { richText: runs.filter((r) => r.text !== '') };
}
