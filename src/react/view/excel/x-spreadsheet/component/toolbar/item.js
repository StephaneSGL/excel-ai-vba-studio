import { cssPrefix } from '../../config';
import tooltip, { shouldShowToolbarTooltip } from '../tooltip';
import { h } from '../element';
import { t } from '../../locale/locale';

export default class Item {
  // tooltip
  // tag: the subclass type
  // shortcut: shortcut key
  constructor(tag, shortcut, value) {
    this.tip = '';
    if (tag) this.tip = t(`toolbar.${tag.replace(/-[a-z]/g, c => c[1].toUpperCase())}`);
    if (shortcut) this.tip += ` (${shortcut})`;
    this.tag = tag;
    this.shortcut = shortcut;
    this.value = value;
    this.el = this.element();
    this.change = () => {};
  }

  element() {
    const { tip } = this;
    const btn = h('div', `${cssPrefix}-toolbar-btn`)
      .on('mouseenter', (evt) => {
        if (this.tip && shouldShowToolbarTooltip(evt, btn)) {
          tooltip(this.tip, btn.el);
        }
      })
      .attr('data-tooltip', tip);
    return btn;
  }
  setState() {}
}
