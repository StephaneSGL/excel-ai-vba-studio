/* global document, window */

export const MAX_HTML_CLIPBOARD_ROWS = 1_048_576;
export const MAX_HTML_CLIPBOARD_COLUMNS = 16_384;
export const MAX_HTML_CLIPBOARD_CELLS = 1_000_000;
export const MAX_HTML_CLIPBOARD_DOM_NODES = 50_000;
export const MAX_HTML_CLIPBOARD_CHARACTERS = 4 * 1024 * 1024;

const DEFAULT_HTML_CLIPBOARD_LIMITS = Object.freeze({
  maxRows: MAX_HTML_CLIPBOARD_ROWS,
  maxColumns: MAX_HTML_CLIPBOARD_COLUMNS,
  maxCells: MAX_HTML_CLIPBOARD_CELLS,
  maxDomNodes: MAX_HTML_CLIPBOARD_DOM_NODES,
});

export class HtmlClipboardLimitError extends Error {
  constructor(message) {
    super(message);
    this.name = 'HtmlClipboardLimitError';
  }
}

export function isHtmlClipboardDestinationInBounds(
  startRow,
  startColumn,
  height,
  width,
) {
  return Number.isSafeInteger(startRow)
    && Number.isSafeInteger(startColumn)
    && Number.isSafeInteger(height)
    && Number.isSafeInteger(width)
    && startRow >= 0
    && startColumn >= 0
    && height > 0
    && width > 0
    && height <= MAX_HTML_CLIPBOARD_ROWS - startRow
    && width <= MAX_HTML_CLIPBOARD_COLUMNS - startColumn;
}

function rejectHtmlClipboard(message) {
  throw new HtmlClipboardLimitError(message);
}

function validateHtmlMarkupBudget(html) {
  let markupTokens = 0;
  let cursor = -1;
  while ((cursor = html.indexOf('<', cursor + 1)) >= 0) {
    markupTokens += 1;
    if (markupTokens > MAX_HTML_CLIPBOARD_DOM_NODES) {
      rejectHtmlClipboard('HTML clipboard payload exceeds the markup budget');
    }
  }
}

function normalizeColor(value) {
  if (!value) return null;
  const text = value.trim().toLowerCase();
  if (!text || text === 'transparent' || text === 'inherit' || text === 'initial') return null;
  if (/^#[0-9a-f]{6}$/i.test(text)) return text;
  if (/^#[0-9a-f]{3}$/i.test(text)) {
    return `#${text[1]}${text[1]}${text[2]}${text[2]}${text[3]}${text[3]}`.toLowerCase();
  }
  const rgb = /^rgba?\(([^)]+)\)$/.exec(text);
  if (rgb) {
    const parts = rgb[1].split(',').map(it => Number.parseFloat(it.trim()));
    if (parts.length >= 3 && parts.every((it, index) => index > 2 || Number.isFinite(it))) {
      if (parts.length > 3 && parts[3] === 0) return null;
      const [r, g, b] = parts.slice(0, 3).map(it => Math.max(0, Math.min(255, Math.round(it))));
      return `#${[r, g, b].map(it => it.toString(16).padStart(2, '0')).join('')}`;
    }
  }
  return null;
}

function parseFontSize(value) {
  if (!value) return null;
  const px = /([\d.]+)px/i.exec(value);
  if (px) return Math.round(Number.parseFloat(px[1]) * 0.75);
  const pt = /([\d.]+)pt/i.exec(value);
  if (pt) return Math.round(Number.parseFloat(pt[1]));
  return null;
}

function borderFromCss(value) {
  if (!value || /none|hidden/i.test(value)) return null;
  const color = normalizeColor(value) || '#000000';
  let style = 'thin';
  if (/double/i.test(value)) style = 'double';
  else if (/dotted/i.test(value)) style = 'dotted';
  else if (/dashed/i.test(value)) style = 'dashed';
  else {
    const width = /([\d.]+)px/i.exec(value);
    if (width && Number.parseFloat(width[1]) >= 2) style = 'medium';
  }
  return [style, color];
}

function styleFromCell(cell) {
  const computed = window.getComputedStyle(cell);
  const style = {};
  const font = {};
  const fontFamily = computed.fontFamily.split(',')[0]?.replace(/^["']|["']$/g, '').trim();
  if (fontFamily) font.name = fontFamily;
  const fontSize = parseFontSize(computed.fontSize);
  if (fontSize) font.size = fontSize;
  if (computed.fontWeight === 'bold' || Number.parseInt(computed.fontWeight, 10) >= 600) font.bold = true;
  if (computed.fontStyle === 'italic') font.italic = true;
  if (Object.keys(font).length > 0) style.font = font;

  const color = normalizeColor(computed.color);
  if (color) style.color = color;
  const bgcolor = normalizeColor(computed.backgroundColor);
  if (bgcolor && bgcolor !== '#ffffff') style.bgcolor = bgcolor;
  if (computed.textDecorationLine.includes('underline')) style.underline = true;
  if (computed.textDecorationLine.includes('line-through')) style.strike = true;
  if (['left', 'center', 'right'].includes(computed.textAlign)) style.align = computed.textAlign;
  if (['top', 'middle', 'bottom'].includes(computed.verticalAlign)) style.valign = computed.verticalAlign;
  if (['pre-wrap', 'pre-line', 'break-spaces'].includes(computed.whiteSpace)) style.textwrap = true;

  const border = {};
  const top = borderFromCss(computed.borderTop);
  const right = borderFromCss(computed.borderRight);
  const bottom = borderFromCss(computed.borderBottom);
  const left = borderFromCss(computed.borderLeft);
  if (top) border.top = top;
  if (right) border.right = right;
  if (bottom) border.bottom = bottom;
  if (left) border.left = left;
  if (Object.keys(border).length > 0) style.border = border;

  return Object.keys(style).length > 0 ? style : null;
}

function collectCellText(root) {
  let text = '';
  const pending = [root];
  while (pending.length > 0) {
    const node = pending.pop();
    if (node.nodeType === 3) {
      text += node.nodeValue || '';
      continue;
    }
    if (node.nodeType !== 1) continue;
    if (node.tagName === 'BR') {
      text += '\n';
      continue;
    }
    const children = node.childNodes;
    for (let index = children.length - 1; index >= 0; index -= 1) {
      pending.push(children[index]);
    }
  }
  return text;
}

function cellText(cell) {
  return collectCellText(cell)
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function escapeHtml(text) {
  return String(text ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function borderSideToCss(side) {
  if (!side || !side[0]) return null;
  const [styleName, color] = side;
  let width = '1px';
  let lineStyle = 'solid';
  if (styleName === 'medium') width = '2px';
  else if (styleName === 'thick') width = '3px';
  else if (styleName === 'dashed') lineStyle = 'dashed';
  else if (styleName === 'dotted') lineStyle = 'dotted';
  else if (styleName === 'double') lineStyle = 'double';
  return `${width} ${lineStyle} ${color}`;
}

function styleToCss(style, defaultStyle) {
  const parts = [];
  const font = style.font || {};
  const defaultFont = defaultStyle.font || {};
  if (font.name && font.name !== defaultFont.name) {
    parts.push(`font-family:${font.name}`);
  }
  if (font.size && font.size !== defaultFont.size) {
    parts.push(`font-size:${font.size}pt`);
  }
  if (font.bold) parts.push('font-weight:bold');
  if (font.italic) parts.push('font-style:italic');
  if (style.color && style.color !== defaultStyle.color) {
    parts.push(`color:${style.color}`);
  }
  if (style.bgcolor && style.bgcolor !== defaultStyle.bgcolor
    && style.bgcolor !== '#ffffff' && style.bgcolor !== '#fff') {
    parts.push(`background-color:${style.bgcolor}`);
  }
  const decorations = [];
  if (style.underline) decorations.push('underline');
  if (style.strike) decorations.push('line-through');
  if (decorations.length > 0) {
    parts.push(`text-decoration:${decorations.join(' ')}`);
  }
  if (style.align && style.align !== defaultStyle.align) {
    parts.push(`text-align:${style.align}`);
  }
  if (style.valign && style.valign !== defaultStyle.valign) {
    parts.push(`vertical-align:${style.valign}`);
  }
  if (style.textwrap) {
    parts.push('white-space:pre-wrap');
  }
  if (style.border) {
    const { border } = style;
    const top = borderSideToCss(border.top);
    const right = borderSideToCss(border.right);
    const bottom = borderSideToCss(border.bottom);
    const left = borderSideToCss(border.left);
    if (top) parts.push(`border-top:${top}`);
    if (right) parts.push(`border-right:${right}`);
    if (bottom) parts.push(`border-bottom:${bottom}`);
    if (left) parts.push(`border-left:${left}`);
  }
  return parts.join(';');
}

function formatCellHtml(text, style, defaultStyle, extraAttrs = '') {
  const css = styleToCss(style, defaultStyle);
  const styleAttr = css ? ` style="${css}"` : '';
  const body = style.textwrap
    ? escapeHtml(text)
    : escapeHtml(text).replace(/\n/g, '<br>');
  return `<td${extraAttrs}${styleAttr}>${body}</td>`;
}

/**
 * Serialize a selected range to HTML table for system clipboard (Excel / WPS).
 */
export function formatHtmlClipboard(data, range) {
  const {
    sri, eri, sci, eci,
  } = range;
  const defaultStyle = data.defaultStyle();
  const rows = [];

  for (let ri = sri; ri <= eri; ri += 1) {
    let rowHtml = '<tr>';
    for (let ci = sci; ci <= eci; ci += 1) {
      const merge = data.merges.getFirstIncludes(ri, ci);
      if (merge && (merge.sri !== ri || merge.sci !== ci)) {
        continue;
      }
      const style = data.getCellStyleOrDefault(ri, ci);
      const text = data.getCellTextOrDefault(ri, ci);
      let extraAttrs = '';
      if (merge) {
        const rowspan = Math.min(merge.eri, eri) - ri + 1;
        const colspan = Math.min(merge.eci, eci) - ci + 1;
        const attrs = [];
        if (rowspan > 1) attrs.push(`rowspan="${rowspan}"`);
        if (colspan > 1) attrs.push(`colspan="${colspan}"`);
        if (attrs.length > 0) extraAttrs = ` ${attrs.join(' ')}`;
      }
      rowHtml += formatCellHtml(text, style, defaultStyle, extraAttrs);
    }
    rowHtml += '</tr>';
    rows.push(rowHtml);
  }

  const table = `<table>${rows.join('')}</table>`;
  return `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>${table}</body></html>`;
}

function validateTableDomSize(table, maximumNodes) {
  let visited = 0;
  const pending = [table];
  while (pending.length > 0) {
    const node = pending.pop();
    visited += 1;
    if (visited > maximumNodes) {
      rejectHtmlClipboard('HTML clipboard table exceeds the DOM traversal budget');
    }
    const children = node.childNodes || [];
    for (let index = children.length - 1; index >= 0; index -= 1) {
      pending.push(children[index]);
    }
  }
}

function parseSpan(rawValue, maximum, label) {
  if (rawValue === null) return 1;
  const normalized = rawValue.trim();
  if (!/^\d+$/.test(normalized)) return 1;
  const significant = normalized.replace(/^0+/, '') || '0';
  const maximumText = String(maximum);
  if (
    significant.length > maximumText.length
    || (significant.length === maximumText.length && significant > maximumText)
  ) {
    rejectHtmlClipboard(`${label} exceeds the worksheet limit`);
  }
  return Math.max(1, Number(significant));
}

export function parseHtmlClipboardTable(
  table,
  limits = DEFAULT_HTML_CLIPBOARD_LIMITS,
) {
  validateTableDomSize(table, limits.maxDomNodes);
  const tableRows = Array.from(table.rows).filter(
    row => typeof row.closest !== 'function' || row.closest('table') === table,
  );
  if (tableRows.length > limits.maxRows) {
    rejectHtmlClipboard('HTML clipboard table exceeds the worksheet row limit');
  }

  const occupiedThroughRow = new Uint32Array(limits.maxColumns);
  const placements = [];
  let expandedCells = 0;
  let height = 0;
  let width = 0;

  tableRows.forEach((tr, ri) => {
    let ci = 0;
    Array.from(tr.cells).forEach((td) => {
      const rowspan = parseSpan(td.getAttribute('rowspan'), limits.maxRows, 'rowspan');
      const colspan = parseSpan(td.getAttribute('colspan'), limits.maxColumns, 'colspan');

      let placed = false;
      while (!placed) {
        while (ci < limits.maxColumns && occupiedThroughRow[ci] > ri) ci += 1;
        if (ci > limits.maxColumns - colspan) {
          rejectHtmlClipboard('HTML clipboard table exceeds the worksheet column limit');
        }
        let blockedColumn = -1;
        for (let column = ci; column < ci + colspan; column += 1) {
          if (occupiedThroughRow[column] > ri) {
            blockedColumn = column;
            break;
          }
        }
        if (blockedColumn < 0) placed = true;
        else ci = blockedColumn + 1;
      }

      if (ri > limits.maxRows - rowspan) {
        rejectHtmlClipboard('HTML clipboard table exceeds the worksheet row limit');
      }
      if (rowspan > Math.floor((limits.maxCells - expandedCells) / colspan)) {
        rejectHtmlClipboard('HTML clipboard table exceeds the expanded-cell budget');
      }
      expandedCells += rowspan * colspan;
      const nextHeight = Math.max(height, ri + rowspan);
      const nextWidth = Math.max(width, ci + colspan);
      if (nextWidth > 0 && nextHeight > Math.floor(limits.maxCells / nextWidth)) {
        rejectHtmlClipboard('HTML clipboard table exceeds the bounding-cell budget');
      }

      for (let column = ci; column < ci + colspan; column += 1) {
        occupiedThroughRow[column] = Math.max(
          occupiedThroughRow[column],
          ri + rowspan,
        );
      }
      placements.push({ td, ri, ci, rowspan, colspan });
      height = nextHeight;
      width = nextWidth;
      ci += colspan;
    });
  });

  if (placements.length === 0) return null;
  const rows = [];
  rows.length = height;
  const merges = [];
  placements.forEach((placement) => {
    const {
      td, ri, ci, rowspan, colspan,
    } = placement;
    const row = rows[ri] || [];
    rows[ri] = row;
    row[ci] = {
      text: cellText(td),
      style: styleFromCell(td),
      merge: rowspan > 1 || colspan > 1 ? [rowspan - 1, colspan - 1] : undefined,
    };
    if (rowspan > 1 || colspan > 1) {
      merges.push({
        sri: ri,
        sci: ci,
        eri: ri + rowspan - 1,
        eci: ci + colspan - 1,
      });
    }
    for (let r = ri; r < ri + rowspan; r += 1) {
      rows[r] = rows[r] || [];
      for (let c = ci; c < ci + colspan; c += 1) {
        if (r !== ri || c !== ci) rows[r][c] = { covered: true };
      }
    }
  });
  return {
    rows,
    merges,
    height,
    width,
  };
}

export function parseHtmlClipboard(html) {
  if (!html) return null;
  if (html.length > MAX_HTML_CLIPBOARD_CHARACTERS) {
    rejectHtmlClipboard('HTML clipboard payload exceeds the input limit');
  }
  validateHtmlMarkupBudget(html);
  const container = document.createElement('div');
  container.style.cssText = 'position:fixed;left:-10000px;top:-10000px;visibility:hidden;pointer-events:none;';
  container.innerHTML = html;
  document.body.appendChild(container);
  try {
    const table = container.querySelector('table');
    if (!table) return null;
    return parseHtmlClipboardTable(table);
  } finally {
    document.body.removeChild(container);
  }
}
