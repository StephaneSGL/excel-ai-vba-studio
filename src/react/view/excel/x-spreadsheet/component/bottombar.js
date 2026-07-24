import { h } from './element';
import { bindClickoutside, unbindClickoutside, mouseMoveUp } from './event';
import { cssPrefix } from '../config';
import FormInput from './form_input';
import { t, tf } from '../locale/locale';

const menuItems = [
  { key: 'rename',    title: tf('contextmenu.renameSheet'),    icon: 'edit' },
  { key: 'duplicate', title: tf('contextmenu.duplicateSheet'), icon: 'files' },
  { key: 'divider' },
  { key: 'delete',    title: tf('contextmenu.deleteSheet'),    icon: 'trash' },
];

const TAB_SCROLL_STEP = 180;

function buildMenuItem(item) {
  if (item.key === 'divider') {
    return h('div', `${cssPrefix}-item divider`);
  }
  return h('div', `${cssPrefix}-item`)
    .children(
      h('i', `codicon codicon-${item.icon}`),
      h('span', 'menu-title').child(item.title()),
    )
    .on('click', () => {
      this.itemClick(item.key);
      this.hide();
    });
}

function buildMenu() {
  return menuItems.map(it => buildMenuItem.call(this, it));
}

class ContextMenu {
  constructor() {
    this.el = h('div', `${cssPrefix}-contextmenu`)
      .css('width', '180px')
      .children(...buildMenu.call(this))
      .hide();
    this.itemClick = () => {};
  }

  hide() {
    const { el } = this;
    el.hide();
    unbindClickoutside(el);
  }

  setOffset(offset) {
    const { el } = this;
    el.offset(offset);
    el.show();
    bindClickoutside(el);
  }
}

function buildNavBtn(name, icon, titleKey, onClick) {
  let iconChild;
  if (name === 'first') {
    iconChild = h('span', `${cssPrefix}-sheet-nav-icon first`).children(
      h('span', 'bar'),
      h('i', 'codicon codicon-chevron-left'),
    );
  } else if (name === 'last') {
    iconChild = h('span', `${cssPrefix}-sheet-nav-icon last`).children(
      h('i', 'codicon codicon-chevron-right'),
      h('span', 'bar'),
    );
  } else {
    iconChild = h('i', `codicon codicon-${icon}`);
  }
  return h('button', `${cssPrefix}-sheet-nav-btn ${name}`)
    .attr('type', 'button')
    .attr('title', t(titleKey))
    .child(iconChild)
    .on('click', (evt) => {
      evt.preventDefault();
      onClick();
    });
}

export default class Bottombar {
  constructor(addFunc = () => {},
    swapFunc = () => {},
    menuFunc = () => {},
    updateFunc = () => {},
    moveFunc = () => {}) {
    this.swapFunc = swapFunc;
    this.updateFunc = updateFunc;
    this.moveFunc = moveFunc;
    this.addFunc = addFunc;
    this.dataNames = [];
    this.activeEl = null;
    this.contextEl = null;
    this.items = [];
    this.dragFromIndex = -1;
    this.moreOpen = false;
    this.dropMarkerEl = h('div', `${cssPrefix}-sheet-drop-marker`).hide();
    this.contextMenu = new ContextMenu();
    this.contextMenu.itemClick = (key) => menuFunc(key);

    this.navFirstEl = buildNavBtn('first', 'chevron-left', 'bottombar.firstSheet', () => {
      this.scrollTabsTo('start');
    });
    this.navPrevEl = buildNavBtn('prev', 'chevron-left', 'bottombar.prevSheet', () => {
      this.scrollTabsBy(-TAB_SCROLL_STEP);
    });
    this.navNextEl = buildNavBtn('next', 'chevron-right', 'bottombar.nextSheet', () => {
      this.scrollTabsBy(TAB_SCROLL_STEP);
    });
    this.navLastEl = buildNavBtn('last', 'chevron-right', 'bottombar.lastSheet', () => {
      this.scrollTabsTo('end');
    });
    this.navEl = h('div', `${cssPrefix}-sheet-nav`).children(
      this.navFirstEl,
      this.navPrevEl,
      this.navNextEl,
      this.navLastEl,
    );

    this.menuEl = h('ul', `${cssPrefix}-menu`);
    this.menuEl.on('scroll', () => this.updateNavState());
    this.menuEl.on('wheel', (evt) => {
      if (Math.abs(evt.deltaY) > Math.abs(evt.deltaX)) {
        evt.preventDefault();
        this.menuEl.el.scrollLeft += evt.deltaY;
        this.updateNavState();
      }
    });

    this.moreListEl = h('div', `${cssPrefix}-sheet-more-list`);
    this.moreMenuEl = h('div', `${cssPrefix}-sheet-more-menu`)
      .child(this.moreListEl)
      .hide();
    this.moreBtnEl = h('button', `${cssPrefix}-sheet-action-btn more`)
      .attr('type', 'button')
      .attr('title', t('bottombar.moreSheets'))
      .child(h('i', 'codicon codicon-ellipsis'))
      .hide()
      .on('click', (evt) => {
        evt.preventDefault();
        evt.stopPropagation();
        this.toggleMoreMenu();
      });

    this.addBtnEl = h('button', `${cssPrefix}-sheet-action-btn add`)
      .attr('type', 'button')
      .attr('title', t('bottombar.addSheet'))
      .child(h('i', 'codicon codicon-add'))
      .on('click', (evt) => {
        evt.preventDefault();
        addFunc();
      });

    this.actionsEl = h('div', `${cssPrefix}-sheet-actions`).children(
      this.moreBtnEl,
      this.moreMenuEl,
      this.addBtnEl,
    );

    this.tabsEl = h('div', `${cssPrefix}-sheet-tabs`).children(
      this.menuEl,
      this.actionsEl,
    );

    this.el = h('div', `${cssPrefix}-bottombar`).children(
      this.contextMenu.el,
      this.dropMarkerEl,
      this.navEl,
      this.tabsEl,
    );

    this.resizeObserver = typeof ResizeObserver !== 'undefined'
      ? new ResizeObserver(() => this.updateNavState())
      : null;
    if (this.resizeObserver) {
      this.resizeObserver.observe(this.menuEl.el);
      this.resizeObserver.observe(this.tabsEl.el);
    }
  }

  scrollTabsBy(delta) {
    this.menuEl.el.scrollLeft += delta;
    this.updateNavState();
  }

  scrollTabsTo(edge) {
    const el = this.menuEl.el;
    el.scrollLeft = edge === 'start' ? 0 : el.scrollWidth;
    this.updateNavState();
  }

  updateNavState() {
    const menu = this.menuEl.el;
    const maxScroll = Math.max(0, menu.scrollWidth - menu.clientWidth);
    const left = menu.scrollLeft;
    const atStart = left <= 1;
    const atEnd = left >= maxScroll - 1;
    const scrollOverflow = maxScroll > 1;
    const moreVisible = this.moreBtnEl.el.style.display !== 'none';
    const moreWidth = moreVisible ? this.moreBtnEl.el.offsetWidth + 2 : 0;
    // Hysteresis: decide visibility as if more were hidden, to avoid show/hide flicker.
    const needsMore = menu.scrollWidth > menu.clientWidth + moreWidth + 1;

    this.setNavDisabled(this.navFirstEl, !scrollOverflow || atStart);
    this.setNavDisabled(this.navPrevEl, !scrollOverflow || atStart);
    this.setNavDisabled(this.navNextEl, !scrollOverflow || atEnd);
    this.setNavDisabled(this.navLastEl, !scrollOverflow || atEnd);

    if (needsMore) {
      this.moreBtnEl.show();
    } else {
      this.hideMoreMenu();
      this.moreBtnEl.hide();
    }
  }

  setNavDisabled(btn, disabled) {
    if (disabled) {
      btn.addClass('disabled');
      btn.attr('disabled', 'disabled');
    } else {
      btn.removeClass('disabled');
      btn.el.removeAttribute('disabled');
    }
  }

  scrollActiveIntoView() {
    if (!this.activeEl) return;
    const menu = this.menuEl.el;
    const tab = this.activeEl.el;
    const tabLeft = tab.offsetLeft;
    const tabRight = tabLeft + tab.offsetWidth;
    const viewLeft = menu.scrollLeft;
    const viewRight = viewLeft + menu.clientWidth;
    if (tabLeft < viewLeft) {
      menu.scrollLeft = Math.max(0, tabLeft - 8);
    } else if (tabRight > viewRight) {
      menu.scrollLeft = tabRight - menu.clientWidth + 8;
    }
    this.updateNavState();
  }

  toggleMoreMenu() {
    if (this.moreOpen) {
      this.hideMoreMenu();
    } else {
      this.showMoreMenu();
    }
  }

  showMoreMenu() {
    this.moreListEl.html('');
    for (let i = 0; i < this.items.length; i += 1) {
      const name = this.dataNames[i];
      const item = this.items[i];
      const row = h('div', `${cssPrefix}-sheet-more-item${item === this.activeEl ? ' active' : ''}`)
        .child(name)
        .on('click', () => {
          this.hideMoreMenu();
          this.clickSwap2(item);
          this.scrollActiveIntoView();
        });
      this.moreListEl.child(row);
    }
    this.moreMenuEl.show();
    this.moreBtnEl.addClass('active');
    this.moreOpen = true;
    bindClickoutside(this.actionsEl, () => this.hideMoreMenu());
  }

  hideMoreMenu() {
    if (!this.moreOpen) return;
    this.moreMenuEl.hide();
    this.moreBtnEl.removeClass('active');
    this.moreOpen = false;
    unbindClickoutside(this.actionsEl);
  }

  bindTabItem(item, options) {
    item.on('mousedown', (evt) => {
      if (evt.button !== 0) return;
      if (options.mode === 'read') {
        this.clickSwap2(item);
        return;
      }
      const fromIndex = this.items.findIndex(it => it === item);
      if (fromIndex < 0) return;
      const startX = evt.clientX;
      let dragging = false;
      let lastClientX = startX;
      evt.preventDefault();
      mouseMoveUp(window, (e) => {
        lastClientX = e.clientX;
        if (!dragging && Math.abs(e.clientX - startX) > 5) {
          dragging = true;
          this.dragFromIndex = fromIndex;
          item.addClass('dragging');
        }
        if (!dragging) return;
        const toIndex = this.findTabDropIndex(e.clientX);
        this.showDropMarker(toIndex);
      }, () => {
        item.removeClass('dragging');
        this.hideDropMarker();
        if (dragging) {
          const toIndex = this.findTabDropIndex(lastClientX);
          if (toIndex >= 0 && toIndex !== fromIndex) {
            this.moveFunc(fromIndex, toIndex);
          }
        } else {
          this.clickSwap2(item);
        }
        this.dragFromIndex = -1;
      });
    }).on('contextmenu', (evt) => {
      if (options.mode === 'read') return;
      evt.preventDefault();
      const rect = item.box();
      const bar = this.el.box();
      this.contextEl = item;
      this.contextMenu.setOffset({
        left: rect.left - bar.left,
        bottom: bar.bottom - rect.top + 1,
      });
    }).on('dblclick', () => {
      if (options.mode === 'read') return;
      const index = this.items.findIndex(it => it === item);
      if (index >= 0) this.startRename(index);
    });
  }

  addItem(name, active, options) {
    this.dataNames.push(name);
    const item = h('li', active ? 'active' : '').child(name);
    this.bindTabItem(item, options);
    if (options.mode === 'read' && !this.addHidden) {
      this.addHidden = true;
      this.addBtnEl.hide();
    }
    if (active) {
      this.clickSwap(item);
    }
    this.items.push(item);
    this.menuEl.child(item);
    requestAnimationFrame(() => {
      if (active) this.scrollActiveIntoView();
      else this.updateNavState();
    });
  }

  insertItem(index, name, active, options) {
    const item = h('li', active ? 'active' : '').child(name);
    this.bindTabItem(item, options);
    this.dataNames.splice(index, 0, name);
    this.items.splice(index, 0, item);
    const next = this.items[index + 1];
    if (next) {
      this.menuEl.el.insertBefore(item.el, next.el);
    } else {
      this.menuEl.child(item);
    }
    if (active) {
      if (this.activeEl !== null) {
        this.activeEl.toggle();
      }
      this.activeEl = item;
      this.swapFunc(index);
    }
    requestAnimationFrame(() => {
      if (active) this.scrollActiveIntoView();
      else this.updateNavState();
    });
  }

  getContextSheetIndex() {
    if (!this.contextEl) return -1;
    return this.items.findIndex(it => it === this.contextEl);
  }

  moveItem(from, to) {
    if (from === to || from < 0 || to < 0 || from >= this.items.length || to >= this.items.length) {
      return;
    }
    const item = this.items[from];
    const name = this.dataNames[from];
    this.items.splice(from, 1);
    this.dataNames.splice(from, 1);
    this.items.splice(to, 0, item);
    this.dataNames.splice(to, 0, name);
    this.menuEl.removeChild(item.el);
    const next = this.items[to + 1];
    if (next) {
      this.menuEl.el.insertBefore(item.el, next.el);
    } else {
      this.menuEl.child(item);
    }
    if (this.activeEl === item) {
      this.swapFunc(to);
    }
    requestAnimationFrame(() => this.updateNavState());
  }

  findTabDropIndex(clientX) {
    const { items } = this;
    if (!items.length) return -1;
    for (let i = 0; i < items.length; i += 1) {
      const rect = items[i].box();
      const mid = rect.left + rect.width / 2;
      if (clientX < mid) return i;
    }
    return items.length - 1;
  }

  showDropMarker(index) {
    if (index < 0 || index >= this.items.length) {
      this.hideDropMarker();
      return;
    }
    const item = this.items[index];
    const rect = item.box();
    const bar = this.el.box();
    this.dropMarkerEl.offset({
      left: rect.left - bar.left,
      bottom: 0,
    }).show();
  }

  hideDropMarker() {
    this.dropMarkerEl.hide();
  }

  startRename(index) {
    if (index < 0 || index >= this.items.length) return;
    const item = this.items[index];
    const v = this.dataNames[index];
    const input = new FormInput('auto', '');
    input.val(v);
    input.input.on('blur', ({ target }) => {
      const { value } = target;
      this.renameItem(index, value.trim() || v);
    });
    input.input.on('keydown', (evt) => {
      if (evt.key === 'Enter') {
        evt.target.blur();
      }
      if (evt.key === 'Escape') {
        this.renameItem(index, v);
      }
    });
    item.html('').child(input.el);
    input.focus();
  }

  renameItem(index, value) {
    this.dataNames.splice(index, 1, value);
    this.items[index].html('').child(value);
    this.updateFunc(index, value);
  }

  clear() {
    this.hideMoreMenu();
    this.items.forEach((it) => {
      this.menuEl.removeChild(it.el);
    });
    this.items = [];
    this.dataNames = [];
    this.contextEl = null;
    this.activeEl = null;
    requestAnimationFrame(() => this.updateNavState());
  }

  deleteItem() {
    const { activeEl } = this;
    const deleteEl = this.contextEl;
    if (!deleteEl || this.items.length <= 1) {
      return [-1];
    }
    const index = this.items.findIndex(it => it === deleteEl);
    this.items.splice(index, 1);
    this.dataNames.splice(index, 1);
    this.menuEl.removeChild(deleteEl.el);
    this.contextEl = null;
    if (activeEl === deleteEl) {
      const [f] = this.items;
      this.activeEl = f;
      this.activeEl.toggle();
      requestAnimationFrame(() => {
        this.scrollActiveIntoView();
        this.updateNavState();
      });
      return [index, 0];
    }
    requestAnimationFrame(() => this.updateNavState());
    return [index, -1];
  }

  clickSwap2(item) {
    const index = this.items.findIndex(it => it === item);
    this.clickSwap(item);
    this.activeEl.toggle();
    this.swapFunc(index);
    this.scrollActiveIntoView();
  }

  clickSwap(item) {
    if (this.activeEl !== null) {
      this.activeEl.toggle();
    }
    this.activeEl = item;
  }
}
