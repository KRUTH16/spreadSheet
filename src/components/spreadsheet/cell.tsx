import { Component, h, Prop, Event, EventEmitter, State, Element, Watch } from '@stencil/core';
import { Cell } from './types';

@Component({
  tag: 'app-cell',
  styleUrl: 'cell.css',
  shadow: true,
})
export class AppCell {
  @Prop() cell!: Cell; // data for this cell
  @Prop() row!: number; // row index
  @Prop() col!: number; // column index

  @Prop() isSelected: boolean = false;
  @Prop() isInRange: boolean = false;

  // @Prop({ reflect: true }) isInRange: boolean = false;

  private hasEmitted: boolean = false;

  @Event({ bubbles: true, composed: true })
  cellValueChanged: EventEmitter<{ row: number; col: number; value: string }>;

  @Event({ bubbles: true, composed: true })
  cellSelected: EventEmitter<{ row: number; col: number }>;

  @Element() element: HTMLElement;

  @State() isEditing: boolean = false;


  @Watch('isSelected')
  resetEdit() {
    this.isEditing = false;
    this.hasEmitted = false;
  }



  private onBlurHandler(e: UIEvent) {
    setTimeout(() => {
      if (!this.isEditing || this.hasEmitted) return;

      const editableDiv = this.element.shadowRoot?.querySelector('div');
      const newValue = editableDiv?.textContent?.trim() || '';

      this.cellValueChanged.emit({
        row: this.row,
        col: this.col,
        value: newValue,
      });

      this.hasEmitted = true;
      this.isEditing = false;
    }, 0);
  }

  private onKeyHandler(e: KeyboardEvent) {
    if (e.key === 'Enter') {
      e.preventDefault(); // Prevent newline in editable div
      const editableDiv = this.element.shadowRoot?.querySelector('div');
      const newValue = editableDiv?.textContent?.trim() || '';

      this.cellValueChanged.emit({ row: this.row, col: this.col, value: newValue });
      this.hasEmitted = true;
      this.isEditing = false;
      (e.target as HTMLElement).blur();
    }
  }


  private startEditing() {
    this.hasEmitted = false;
    this.isEditing = true;

    setTimeout(() => {
      const el = this.element.shadowRoot.querySelector('div');

      if (el) {
        el.textContent = this.cell.value ?? '';
   
        el.focus();

        // Place cursor at end
        const range = document.createRange();
        range.selectNodeContents(el);
        range.collapse(false);
        const sel = window.getSelection();
        sel?.removeAllRanges();
        sel?.addRange(range);
      }
    }, 0);
  }

  render() {
    
    const cell = this.cell ?? { value: '', formula: undefined, style: {} };
    const style = cell.style ?? {};

    return (
      <div
        class={{ cell: true, selected: this.isSelected, range: this.isInRange }}
        // selected={this.isSelected ? 'true' : 'false'}

        style={{
          fontWeight: style.bold ? 'bold' : 'normal',
          fontStyle: style.italic ? 'italic' : 'normal',
          textDecoration: style.underline ? 'underline' : 'none',
          backgroundColor: style.bgColor ?? 'white',
          textAlign: style.align ?? 'left',
          color: style.color ?? 'black',
        }}
        contentEditable={this.isEditing}
        // onClick={() => this.cellSelected.emit({ row: this.row, col: this.col })}
        onClick={() => {
          if (!this.isEditing) {
            this.cellSelected.emit({ row: this.row, col: this.col });
          }
        }}
        onDblClick={() => this.startEditing()}
        onBlur={e => this.onBlurHandler(e)}
        onKeyDown={e => this.onKeyHandler(e)}
      >
        {!this.isEditing ? cell.value ?? '' : ''}
      </div>
    );
  }
}
