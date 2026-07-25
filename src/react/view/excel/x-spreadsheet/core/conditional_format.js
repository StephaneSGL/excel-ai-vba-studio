import { expr2xy } from './alphabet';

function normalizeAddress(address) {
  const withoutSheet = `${address || ''}`.split('!').pop().replace(/\$/g, '');
  return withoutSheet;
}

function parseRange(ref) {
  const [start, end = start] = normalizeAddress(ref).split(':');
  if (!start || !end) return null;
  try {
    const [sci, sri] = expr2xy(start);
    const [eci, eri] = expr2xy(end);
    return {
      sri: Math.min(sri, eri),
      sci: Math.min(sci, eci),
      eri: Math.max(sri, eri),
      eci: Math.max(sci, eci),
    };
  } catch {
    return null;
  }
}

function rangesForRef(ref) {
  return `${ref || ''}`
    .trim()
    .split(/\s+/)
    .map(parseRange)
    .filter(Boolean);
}

function cellValue(cell) {
  if (!cell) return '';
  if (cell.formulaResult !== undefined) return cell.formulaResult;
  const text = `${cell.text ?? ''}`.trim();
  if (text === '') return '';
  const number = Number(text.replace(/\s/g, '').replace(',', '.'));
  return Number.isNaN(number) ? text : number;
}

function numberValue(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  const parsed = Number(`${value ?? ''}`.replace(/\s/g, '').replace(',', '.'));
  return Number.isFinite(parsed) ? parsed : null;
}

function valuesInRanges(data, ranges) {
  const values = [];
  ranges.forEach((range) => {
    for (let ri = range.sri; ri <= range.eri; ri += 1) {
      for (let ci = range.sci; ci <= range.eci; ci += 1) {
        const value = numberValue(cellValue(data.rows.getCell(ri, ci)));
        if (value !== null) values.push(value);
      }
    }
  });
  return values;
}

function inRanges(ri, ci, ranges) {
  return ranges.some(range => (
    ri >= range.sri && ri <= range.eri && ci >= range.sci && ci <= range.eci
  ));
}

function formulaValue(formula) {
  if (typeof formula === 'number' || typeof formula === 'boolean') return formula;
  const text = `${formula ?? ''}`.trim().replace(/^=/, '');
  if (/^".*"$/.test(text)) return text.slice(1, -1);
  const number = numberValue(text);
  return number === null ? text : number;
}

function compareCell(rule, value) {
  const formulae = Array.isArray(rule.formulae) ? rule.formulae.map(formulaValue) : [];
  const first = formulae[0];
  const second = formulae[1];
  const numericValue = numberValue(value);
  const numericFirst = numberValue(first);
  const numericSecond = numberValue(second);
  const left = numericValue !== null && numericFirst !== null ? numericValue : `${value ?? ''}`;
  const right = numericValue !== null && numericFirst !== null ? numericFirst : `${first ?? ''}`;
  switch (rule.operator) {
    case 'equal': return left === right;
    case 'notEqual': return left !== right;
    case 'greaterThan': return left > right;
    case 'greaterThanOrEqual': return left >= right;
    case 'lessThan': return left < right;
    case 'lessThanOrEqual': return left <= right;
    case 'between':
      return numericValue !== null && numericFirst !== null && numericSecond !== null
        && numericValue >= numericFirst && numericValue <= numericSecond;
    case 'notBetween':
      return numericValue !== null && numericFirst !== null && numericSecond !== null
        && (numericValue < numericFirst || numericValue > numericSecond);
    default:
      return false;
  }
}

function colorHex(color, fallback = '#638ec6') {
  const argb = color && typeof color === 'object' ? color.argb : undefined;
  const normalized = `${argb || ''}`.replace(/^#/, '');
  if (/^[0-9a-f]{8}$/i.test(normalized)) return `#${normalized.slice(2).toLowerCase()}`;
  if (/^[0-9a-f]{6}$/i.test(normalized)) return `#${normalized.toLowerCase()}`;
  return fallback;
}

function mixHex(start, end, ratio) {
  const clamp = Math.max(0, Math.min(1, ratio));
  const components = [1, 3, 5].map((offset) => {
    const from = parseInt(start.slice(offset, offset + 2), 16);
    const to = parseInt(end.slice(offset, offset + 2), 16);
    return Math.round(from + (to - from) * clamp).toString(16).padStart(2, '0');
  });
  return `#${components.join('')}`;
}

function percentile(values, percent) {
  if (!values.length) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  const position = (sorted.length - 1) * Math.max(0, Math.min(100, percent)) / 100;
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  if (lower === upper) return sorted[lower];
  return sorted[lower] + (sorted[upper] - sorted[lower]) * (position - lower);
}

function thresholdValue(cfvo, values) {
  const min = values.length ? Math.min(...values) : 0;
  const max = values.length ? Math.max(...values) : 0;
  const raw = numberValue(cfvo?.value ?? cfvo?.val) ?? 0;
  switch (cfvo?.type) {
    case 'min': return min;
    case 'max': return max;
    case 'percent': return min + (max - min) * raw / 100;
    case 'percentile': return percentile(values, raw);
    default: return raw;
  }
}

function colorScaleStyle(rule, value, values) {
  const number = numberValue(value);
  const colors = Array.isArray(rule.color) ? rule.color.map(item => colorHex(item)) : [];
  const points = Array.isArray(rule.cfvo)
    ? rule.cfvo.map(item => thresholdValue(item, values))
    : [];
  if (number === null || colors.length < 2 || points.length !== colors.length) return null;
  if (number <= points[0]) return { bgcolor: colors[0] };
  for (let index = 1; index < points.length; index += 1) {
    if (number <= points[index]) {
      const span = points[index] - points[index - 1];
      const ratio = span === 0 ? 1 : (number - points[index - 1]) / span;
      return { bgcolor: mixHex(colors[index - 1], colors[index], ratio) };
    }
  }
  return { bgcolor: colors[colors.length - 1] };
}

const icon = (glyph, color) => ({ glyph, color });
const RED = '#c62828';
const ORANGE = '#ef6c00';
const YELLOW = '#f9a825';
const GREEN = '#2e7d32';
const GRAY = '#667085';
const BLACK = '#202124';

// Canvas text cannot reliably render colour emoji in every VS Code font stack.
// Use plain glyphs with an explicit colour so workbook icon sets are always visible.
const ICON_SETS = {
  '3Arrows': [icon('▼', RED), icon('→', YELLOW), icon('▲', GREEN)],
  '3ArrowsGray': [icon('▼', GRAY), icon('→', GRAY), icon('▲', GRAY)],
  '3Flags': [icon('▼', RED), icon('◆', YELLOW), icon('▲', GREEN)],
  '3Signs': [icon('●', RED), icon('●', YELLOW), icon('●', GREEN)],
  '3Symbols': [icon('×', RED), icon('!', YELLOW), icon('✓', GREEN)],
  '3Symbols2': [icon('×', RED), icon('!', YELLOW), icon('✓', GREEN)],
  '3TrafficLights1': [icon('●', RED), icon('●', YELLOW), icon('●', GREEN)],
  '3TrafficLights2': [icon('●', RED), icon('●', YELLOW), icon('●', GREEN)],
  '4Arrows': [icon('↓', RED), icon('↘', ORANGE), icon('↗', YELLOW), icon('↑', GREEN)],
  '4ArrowsGray': [icon('↓', GRAY), icon('↘', GRAY), icon('↗', GRAY), icon('↑', GRAY)],
  '4Rating': [icon('●', RED), icon('●●', ORANGE), icon('●●●', YELLOW), icon('●●●●', GREEN)],
  '4RedToBlack': [icon('●', RED), icon('●', ORANGE), icon('●', BLACK), icon('●', BLACK)],
  '4TrafficLights': [icon('●', RED), icon('●', ORANGE), icon('●', YELLOW), icon('●', GREEN)],
  '5Arrows': [icon('↓', RED), icon('↘', ORANGE), icon('→', YELLOW), icon('↗', '#7cb342'), icon('↑', GREEN)],
  '5ArrowsGray': [icon('↓', GRAY), icon('↘', GRAY), icon('→', GRAY), icon('↗', GRAY), icon('↑', GRAY)],
  '5Quarters': [icon('○', RED), icon('◔', ORANGE), icon('◑', YELLOW), icon('◕', '#7cb342'), icon('●', GREEN)],
  '5Rating': [icon('●', RED), icon('●●', ORANGE), icon('●●●', YELLOW), icon('●●●●', '#7cb342'), icon('●●●●●', GREEN)],
};

function iconForRule(rule, value, values) {
  const number = numberValue(value);
  if (number === null) return null;
  const thresholds = Array.isArray(rule.cfvo)
    ? rule.cfvo.map(item => thresholdValue(item, values))
    : [];
  if (!thresholds.length) return null;
  let index = 0;
  thresholds.forEach((threshold, thresholdIndex) => {
    if (number >= threshold) index = thresholdIndex;
  });
  const icons = ICON_SETS[rule.iconSet] || ICON_SETS['3TrafficLights1'];
  index = Math.max(0, Math.min(icons.length - 1, index));
  if (rule.reverse) index = icons.length - 1 - index;
  return {
    ...icons[index],
    showValue: rule.showValue !== false,
  };
}

function dataBarForRule(rule, value, values) {
  const number = numberValue(value);
  if (number === null || !values.length) return null;
  const min = rule.cfvo?.[0] ? thresholdValue(rule.cfvo[0], values) : Math.min(...values);
  const max = rule.cfvo?.[1] ? thresholdValue(rule.cfvo[1], values) : Math.max(...values);
  const ratio = max === min ? 1 : (number - min) / (max - min);
  return {
    ratio: Math.max(0, Math.min(1, ratio)),
    color: colorHex(rule.color, '#638ec6'),
  };
}

function ruleMatches(rule, value, values) {
  switch (rule.type) {
    case 'cellIs':
      return compareCell(rule, value);
    case 'containsText': {
      const needle = `${rule.text ?? formulaValue(rule.formulae?.[0]) ?? ''}`.toLocaleLowerCase();
      return needle !== '' && `${value ?? ''}`.toLocaleLowerCase().includes(needle);
    }
    case 'top10': {
      const number = numberValue(value);
      if (number === null || !values.length) return false;
      const rank = Math.max(1, Number(rule.rank) || 10);
      const sorted = [...values].sort((a, b) => (rule.bottom ? a - b : b - a));
      const threshold = sorted[Math.min(sorted.length - 1, rank - 1)];
      return rule.bottom ? number <= threshold : number >= threshold;
    }
    case 'aboveAverage': {
      const number = numberValue(value);
      if (number === null || !values.length) return false;
      const average = values.reduce((sum, item) => sum + item, 0) / values.length;
      return rule.aboveAverage === false ? number < average : number > average;
    }
    case 'colorScale':
    case 'iconSet':
    case 'dataBar':
      return true;
    default:
      return false;
  }
}

function mergePreferred(current, incoming) {
  if (!incoming) return current;
  const merged = { ...incoming, ...current };
  if (incoming.font || current.font) {
    merged.font = { ...(incoming.font || {}), ...(current.font || {}) };
  }
  return merged;
}

export function evaluateConditionalFormatting(data, ri, ci) {
  const definitions = Array.isArray(data.conditionalFormattings)
    ? data.conditionalFormattings
    : [];
  if (!definitions.length) return null;
  const result = { style: {}, icon: null, dataBar: null };
  let matched = false;
  definitions.forEach((definition) => {
    const ranges = rangesForRef(definition.ref);
    if (!inRanges(ri, ci, ranges)) return;
    const values = valuesInRanges(data, ranges);
    const value = cellValue(data.rows.getCell(ri, ci));
    const rules = [...(definition.rules || [])].sort(
      (left, right) => (Number(left.priority) || Number.MAX_SAFE_INTEGER)
        - (Number(right.priority) || Number.MAX_SAFE_INTEGER),
    );
    for (const rule of rules) {
      if (!ruleMatches(rule, value, values)) continue;
      matched = true;
      if (rule.type === 'colorScale') {
        result.style = mergePreferred(result.style, colorScaleStyle(rule, value, values));
      } else if (rule.type === 'iconSet') {
        result.icon = result.icon || iconForRule(rule, value, values);
      } else if (rule.type === 'dataBar') {
        result.dataBar = result.dataBar || dataBarForRule(rule, value, values);
      } else if (rule.displayStyle) {
        result.style = mergePreferred(result.style, rule.displayStyle);
      }
      if (rule.stopIfTrue) break;
    }
  });
  return matched ? result : null;
}
