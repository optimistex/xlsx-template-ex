import {Cell} from "exceljs";

export class CellRange {
  constructor(public top: number, public left: number, public bottom: number, public right: number) {
  }

  public static createFromCells(cellTL: Cell, cellBR: Cell): CellRange {
    return new CellRange(+cellTL.row, +cellTL.col, +cellBR.row, +cellBR.col);
  }

  /** Just for clone */
  public static createFromRange(range: CellRange): CellRange {
    return new CellRange(range.top, range.left, range.bottom, range.right);
  }

  public get valid() {
    return this.top > 0 && this.top <= this.bottom && this.left >= 0 && this.left <= this.right;
  }

  public get countRows(): number {
    return this.bottom - this.top + 1;
  }

  public get countColumns(): number {
    return this.right - this.left + 1;
  }

  public move(dRow: number, dCol: number): void {
    this.top += dRow;
    this.bottom += dRow;

    this.left += dCol;
    this.right += dCol;
  }

  // public grow(range: CellRange): void {
  //   this.top = Math.min(this.top, range.top);
  //   this.left = Math.min(this.left, range.left);
  //   this.bottom = Math.max(this.bottom, range.bottom);
  //   this.right = Math.max(this.right, range.right);
  // }
}
