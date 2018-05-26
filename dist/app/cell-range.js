"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
class CellRange {
    constructor(top, left, bottom, right) {
        this.top = top;
        this.left = left;
        this.bottom = bottom;
        this.right = right;
    }
    static createFromCells(cellTL, cellBR) {
        return new CellRange(+cellTL.row, +cellTL.col, +cellBR.row, +cellBR.col);
    }
    /** Just for clone */
    static createFromRange(range) {
        return new CellRange(range.top, range.left, range.bottom, range.right);
    }
    get countRows() {
        return this.bottom - this.top + 1;
    }
    get countColumns() {
        return this.right - this.left + 1;
    }
    move(dRow, dCol) {
        this.top += dRow;
        this.bottom += dRow;
        this.left += dCol;
        this.right += dCol;
    }
}
exports.CellRange = CellRange;
