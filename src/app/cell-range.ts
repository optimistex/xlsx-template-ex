import {Cell} from "exceljs";

export class CellRange {
  public top: number;
  public left: number;
  public bottom: number;
  public right: number;

  constructor(top: number, left: number, bottom: number, right: number) {
    this.top = top;
    this.left = left;
    this.bottom = bottom;
    this.right = right;
  }

  public static createFromCells(cellTL: Cell, cellBR: Cell): CellRange {
    return new CellRange(+cellTL.row, +cellTL.col, +cellBR.row, +cellBR.col);
  }

  /** Just for clone */
  public static createFromRange(range: CellRange): CellRange {
    return new CellRange(range.top, range.left, range.bottom, range.right);
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

  public grow(range: CellRange): void {
    this.top = Math.min(this.top, range.top);
    this.left = Math.min(this.left, range.left);
    this.bottom = Math.max(this.bottom, range.bottom);
    this.right = Math.max(this.right, range.right);
  }
}
