/**
 * @property {(number|string)} top
 * @property {(number|string)} left
 * @property {(number|string)} bottom
 * @property {(number|string)} right
 */
class CellRange {
    /**
     * @param {(number|string)} top
     * @param {(number|string)} left
     * @param {(number|string)} bottom
     * @param {(number|string)} right
     */
    constructor(top, left, bottom, right) {
        this.top = top;
        this.left = left;
        this.bottom = bottom;
        this.right = right;
    }

    /**
     * @param {Cell} cellTL top left
     * @param {Cell} cellBR bottom right
     * @return {CellRange}
     */
    static createFromCells(cellTL, cellBR) {
        return new CellRange(cellTL.row, cellTL.col, cellBR.row, cellBR.col);
    }

    /**
     * Just for clone
     * @param range
     * @return CellRange
     */
    static createFromRange(range) {
        return new CellRange(range.top, range.left, range.bottom, range.right);
    }

    /**
     * @return {number}
     */
    get countRows() {
        return this.bottom - this.top + 1;
    }

    /**
     * @return {number}
     */
    get countColumns() {
        return this.right - this.left + 1;
    }

    /**
     * @param {number} dRow
     * @param {number} dCol
     */
    move(dRow, dCol) {
        this.top += dRow;
        this.bottom += dRow;

        this.left += dCol;
        this.right += dCol;
    }

    /**
     * @param range
     */
    grow(range) {
        this.top = Math.min(this.top, range.top);
        this.left = Math.min(this.left, range.left);
        this.bottom = Math.max(this.bottom, range.bottom);
        this.right = Math.max(this.right, range.right);
    }
}

module.exports = CellRange;