//* global window */
import { h } from './element';
import Suggest from './suggest';
import Datepicker from './datepicker';
import { cssPrefix } from '../config';
import { getFontSizePxByPt } from '../core/font';
import { resolveExcelCellBg, resolveExcelCellColor } from '../../theme';
// import { mouseMoveUp } from '../event';

const EDITOR_CELL_HORIZONTAL_PADDING = 20;
// editor-area uses border-box; 2px border on top and bottom
const EDITOR_AREA_BORDER_INSET = 4;

function escapeCssDoubleQuotedString(value) {
  return `${value}`
    .replace(/\\/g, '\\\\')
    .replace(/"/g, '\\"')
    .replace(/[\r\n\f]/g, ' ');
}

function applyTextareaLayout(editor, fontSizePt) {
  const { textEl, areaOffset } = editor;
  if (!areaOffset) return;
  editor.editorFontSizePt = fontSizePt;
  const { height } = areaOffset;
  const sizePx = getFontSizePxByPt(fontSizePt ?? 11);
  const lineH = sizePx + 2;
  const contentHeight = height - EDITOR_AREA_BORDER_INSET;
  textEl.css('box-sizing', 'border-box');
  textEl.css('line-height', `${lineH}px`);
  textEl.css('height', `${contentHeight}px`);
  const paddingTop = Math.max(0, Math.round((contentHeight - lineH) / 2));
  textEl.css('padding-top', `${paddingTop}px`);
  textEl.css('padding-bottom', '0');
}

function buildEditorFontCss(font, defaultFontName = 'Arial') {
  const sizePx = getFontSizePxByPt(font?.size ?? 11);
  const parts = [];
  if (font?.italic) parts.push('italic');
  if (font?.bold) parts.push('bold');
  parts.push(`${sizePx}px`);
  const name = escapeCssDoubleQuotedString(font?.name || defaultFontName);
  parts.push(`"${name}"`);
  return parts.join(' ');
}

function resetTextareaSize() {
  const {
    areaEl, areaOffset, suggest, suggestPosition, textEl,
  } = this;
  if (!areaOffset) return;

  window.requestAnimationFrame(() => {
    if (!this.areaOffset || !suggest) return;
    const textarea = textEl.el;
    const lineCount = textarea.value.split('\n').length;
    if (lineCount > 1) {
      textEl.css('padding-top', '0');
    } else {
      applyTextareaLayout(this, this.editorFontSizePt ?? 11);
    }
    if (textarea.scrollHeight > textarea.clientHeight + 1) {
      textEl.css('padding-top', '0');
    }
    textEl.css('overflow-y', textarea.scrollHeight > textarea.clientHeight + 1 ? 'auto' : 'hidden');
    const sOffset = { left: 0 };
    sOffset[suggestPosition || 'top'] = areaEl.offset().height;
    suggest.setOffset(sOffset);
  });
}

function insertText({ target }, itxt) {
  const { value, selectionEnd } = target;
  const ntxt = `${value.slice(0, selectionEnd)}${itxt}${value.slice(selectionEnd)}`;
  target.value = ntxt;
  target.setSelectionRange(selectionEnd + 1, selectionEnd + 1);

  this.inputText = ntxt;
  this.textlineEl.html(ntxt);
  resetTextareaSize.call(this);
}

function keydownEventHandler(evt) {
  const { keyCode, altKey } = evt;
  if (keyCode !== 13 && keyCode !== 9) evt.stopPropagation();
  if (keyCode === 27) {
    evt.preventDefault();
    this.change('cancel', this.inputText);
    return;
  }
  // macOS Option key inserts special symbols; allow Option+Enter for newline only
  if (altKey && keyCode !== 13) {
    evt.preventDefault();
    return;
  }
  if (keyCode === 13 && altKey) {
    insertText.call(this, evt, '\n');
    evt.stopPropagation();
  }
  if (keyCode === 13 && !altKey) {
    evt.preventDefault();
    this.clear();
  }
  // 不知道为什么单元格事件被吞了, Windows上正常
  if (evt.metaKey && navigator.userAgent.includes('Mac OS')) {
    const newEvent = new evt.constructor(evt.type, evt);
    document.body.dispatchEvent(newEvent);
  }
}

function inputEventHandler(evt) {
  const v = evt.target.value;
  // console.log(evt, 'v:', v);
  const { suggest, textlineEl, validator } = this;
  const { cell } = this;
  this.referenceRange = null;
  if (cell !== null) {
    if (('editable' in cell && cell.editable === true) || (cell.editable === undefined)) {
      this.inputText = v;
      if (validator) {
        if (validator.type === 'list') {
          suggest.search(v);
        } else {
          suggest.hide();
        }
      } else {
        const start = v.lastIndexOf('=');
        if (start !== -1) {
          suggest.search(v.substring(start + 1));
        } else {
          suggest.hide();
        }
      }
      textlineEl.html(v);
      resetTextareaSize.call(this);
      this.change('input', v);
    } else {
      evt.target.value = cell.text || '';
    }
  } else {
    this.inputText = v;
    if (validator) {
      if (validator.type === 'list') {
        suggest.search(v);
      } else {
        suggest.hide();
      }
    } else {
      const start = v.lastIndexOf('=');
      if (start !== -1) {
        suggest.search(v.substring(start + 1));
      } else {
        suggest.hide();
      }
    }
    textlineEl.html(v);
    resetTextareaSize.call(this);
    this.change('input', v);
  }
}

function setTextareaRange(position) {
  const { el } = this.textEl;
  setTimeout(() => {
    el.focus();
    el.setSelectionRange(position, position);
  }, 0);
}

function setText(text, position) {
  const { textEl, textlineEl } = this;
  // firefox bug
  textEl.el.blur();

  textEl.val(text);
  textlineEl.html(text);
  setTextareaRange.call(this, position);
}

function suggestItemClick(it) {
  const { inputText, validator } = this;
  let position = 0;
  if (validator && validator.type === 'list') {
    this.inputText = it;
    position = this.inputText.length;
  } else {
    const start = inputText.lastIndexOf('=');
    const sit = inputText.substring(0, start + 1);
    let eit = inputText.substring(start + 1);
    if (eit.indexOf(')') !== -1) {
      eit = eit.substring(eit.indexOf(')'));
    } else {
      eit = '';
    }
    this.inputText = `${sit + it.key}(`;
    // console.log('inputText:', this.inputText);
    position = this.inputText.length;
    this.inputText += `)${eit}`;
  }
  setText.call(this, this.inputText, position);
}

function resetSuggestItems() {
  this.suggest.setItems(this.formulas);
}

function updateEditorText(editor, text, position) {
  editor.inputText = text;
  editor.textEl.val(text);
  editor.textlineEl.html(text);
  editor.textEl.el.focus();
  editor.textEl.el.setSelectionRange(position, position);
  resetTextareaSize.call(editor);
}

function dateFormat(d) {
  let month = d.getMonth() + 1;
  let date = d.getDate();
  if (month < 10) month = `0${month}`;
  if (date < 10) date = `0${date}`;
  return `${d.getFullYear()}-${month}-${date}`;
}

export default class Editor {
  constructor(formulas, viewFn, rowHeight) {
    this.viewFn = viewFn;
    this.rowHeight = rowHeight;
    this.formulas = formulas;
    this.suggest = new Suggest(formulas, (it) => {
      suggestItemClick.call(this, it);
    });
    this.datepicker = new Datepicker();
    this.datepicker.change((d) => {
      // console.log('d:', d);
      this.setText(dateFormat(d));
      this.clear();
    });
    this.areaEl = h('div', `${cssPrefix}-editor-area`)
      .children(
        this.textEl = h('textarea', '')
          .on('input', evt => inputEventHandler.call(this, evt))
          .on('paste.stop', () => { })
          .on('keydown', evt => keydownEventHandler.call(this, evt)),
        this.textlineEl = h('div', 'textline'),
        this.suggest.el,
        this.datepicker.el,
      )
      .on('mousemove.stop', () => { })
      .on('mousedown.stop', () => { });
    this.el = h('div', `${cssPrefix}-editor`)
      .child(this.areaEl).hide();
    this.suggest.bindInputEvents(this.textEl);

    this.areaOffset = null;
    this.freeze = { w: 0, h: 0 };
    this.cell = null;
    this.inputText = '';
    this.referenceRange = null;
    this.change = () => { };
  }

  setFreezeLengths(width, height) {
    this.freeze.w = width;
    this.freeze.h = height;
  }

  clear() {
    // const { cell } = this;
    // const cellText = (cell && cell.text) || '';
    if (this.inputText !== '') {
      this.change('finished', this.inputText);
    }
    this.cell = null;
    this.areaOffset = null;
    this.inputText = '';
    this.referenceRange = null;
    this.setFormulaTargetActive(false);
    this.el.hide();
    this.textEl.val('');
    this.textlineEl.html('');
    resetSuggestItems.call(this);
    this.datepicker.hide();
  }

  cancel() {
    this.cell = null;
    this.areaOffset = null;
    this.inputText = '';
    this.referenceRange = null;
    this.setFormulaTargetActive(false);
    this.el.hide();
    this.textEl.val('');
    this.textlineEl.html('');
    resetSuggestItems.call(this);
    this.datepicker.hide();
  }

  setOffset(offset, suggestPosition = 'top') {
    const {
      textEl, areaEl, suggest, freeze, el,
    } = this;
    if (offset) {
      this.areaOffset = offset;
      this.suggestPosition = suggestPosition;
      const {
        left, top, width, height, l, t,
      } = offset;
      const elOffset = { left: 0, top: 0 };
      if (freeze.w > l && freeze.h > t) {
        //
      } else if (freeze.w < l && freeze.h < t) {
        elOffset.left = freeze.w;
        elOffset.top = freeze.h;
      } else if (freeze.w > l) {
        elOffset.top = freeze.h;
      } else if (freeze.h > t) {
        elOffset.left = freeze.w;
      }
      el.offset(elOffset);
      areaEl.offset({ left: left - elOffset.left, top: top - elOffset.top, width });
      areaEl.css('height', `${height}px`);
      areaEl.css('min-height', `${height}px`);
      areaEl.css('overflow', 'hidden');
      textEl.offset({ width: width - EDITOR_CELL_HORIZONTAL_PADDING });
      textEl.cssRemoveKeys('min-height');
      applyTextareaLayout(this, 11);
      resetTextareaSize.call(this);
      suggest.hide();
    }
  }

  setCell(cell, validator, cellStyle) {
    if (cell && cell.editable === false) return;

    // console.log('::', validator);
    const { el, datepicker, suggest } = this;
    el.show();
    this.cell = cell;
    const text = (cell && cell.text) || '';
    this.setText(text);
    if (cellStyle) {
      this.applyCellStyle(cellStyle);
    }

    this.validator = validator;
    if (validator) {
      const { type } = validator;
      if (type === 'date') {
        datepicker.show();
        if (!/^\s*$/.test(text)) {
          datepicker.setValue(text);
        }
      }
      if (type === 'list') {
        suggest.setItems(validator.values());
        suggest.search('');
      }
    }
  }

  setText(text, position = text.length) {
    this.inputText = text;
    this.referenceRange = null;
    // console.log('text>>:', text);
    setText.call(this, text, position);
    resetTextareaSize.call(this);
  }

  setFormulaReference(referenceText) {
    const textarea = this.textEl.el;
    const value = textarea.value;
    let start = textarea.selectionStart == null ? value.length : textarea.selectionStart;
    let end = textarea.selectionEnd == null ? start : textarea.selectionEnd;

    if (this.referenceRange && this.referenceRange.end <= value.length) {
      start = this.referenceRange.start;
      end = this.referenceRange.end;
    }

    const nextText = `${value.slice(0, start)}${referenceText}${value.slice(end)}`;
    const nextPosition = start + referenceText.length;
    this.referenceRange = { start, end: nextPosition };
    updateEditorText(this, nextText, nextPosition);
    this.change('input', nextText);
  }

  setFormulaReferenceRange(start, end) {
    this.referenceRange = { start, end };
  }

  setCursorPosition(position) {
    setTextareaRange.call(this, position);
  }

  setFormulaTargetActive(active) {
    this.el.active(active, 'formula-target-mode');
  }

  applyCellStyle(cellStyle) {
    const { textEl, textlineEl, areaEl } = this;
    if (!cellStyle) return;
    const font = cellStyle.font;
    if (font) {
      const fontCss = buildEditorFontCss(font);
      const lineHeight = `${getFontSizePxByPt(font.size ?? 11) + 2}px`;
      textEl.css('font', fontCss);
      textlineEl.css('font', fontCss);
      textlineEl.css('line-height', lineHeight);
      applyTextareaLayout(this, font.size ?? 11);
      resetTextareaSize.call(this);
    }
    const color = resolveExcelCellColor(cellStyle.color);
    textEl.css('color', color);
    textlineEl.css('color', color);
    areaEl.css('background-color', resolveExcelCellBg(cellStyle.bgcolor));
    textEl.css('background-color', 'transparent');
  }
}

export { escapeCssDoubleQuotedString };
