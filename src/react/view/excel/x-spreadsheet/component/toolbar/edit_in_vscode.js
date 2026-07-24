import Item from './item';
import Icon from '../icon';

export default class EditInVSCode extends Item {
  constructor() {
    super('editInVscode');
  }

  element() {
    return super.element()
      .addClass('vscode')
      .child(new Icon('vscode'))
      .on('click', () => {
        this.change('edit-in-vscode');
      });
  }
}
