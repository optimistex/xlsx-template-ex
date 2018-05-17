const _ = require('lodash');
const Excel = require('exceljs');
const CellRange = require('./cell-range');

/**
 * Callback for iterate cells
 * @callback iterateCells
 * @param {Cell} cell
 * @return false - whether to break iteration
 */

/**
 * @property {Worksheet} worksheet
 */
class WorkSheetHelper {
    /**
     * @param {Worksheet} worksheet
     */
    constructor(worksheet) {
        this.worksheet = worksheet;
    }

    /**
     * @return {Workbook}
     */
    get workbook() {
        return this.worksheet.workbook;
    }

    /**
     * @param {string} fileName
     * @param {Cell} cell
     */
    addImage(fileName, cell) {
        const imgId = this.workbook.addImage({filename: fileName, extension: 'jpeg'});

        const cellRange = this.getMergeRange(cell);
        if (cellRange) {
            this.worksheet.addImage(imgId, {
                tl: {col: cellRange.left - 0.99999, row: cellRange.top - 0.99999},
                br: {col: cellRange.right, row: cellRange.bottom}
            });
        } else {
            this.worksheet.addImage(imgId, {
                tl: {col: cell.col - 0.99999, row: cell.row - 0.99999},
                br: {col: cell.col, row: cell.row},
            });
        }
    }

    /**
     * @param {number|string} srcRowStart
     * @param {number|string} srcRowEnd
     * @param {number} countClones
     */
    cloneRows(srcRowStart, srcRowEnd, countClones = 1) {
        const countRows = srcRowEnd - srcRowStart + 1;
        const dxRow = countRows * countClones;
        const lastRow = this.getSheetDimension().bottom + dxRow;

        // Move rows below
        for (let rowSrcNumber = lastRow; rowSrcNumber > srcRowEnd; rowSrcNumber--) {
            const rowSrc = this.worksheet.getRow(rowSrcNumber);
            const rowDest = this.worksheet.getRow(rowSrcNumber + dxRow);
            this.moveRow(rowSrc, rowDest);
        }

        // Clone target rows
        for (let rowSrcNumber = srcRowEnd; rowSrcNumber >= srcRowStart; rowSrcNumber--) {
            const rowSrc = this.worksheet.getRow(rowSrcNumber);
            for (let cloneNumber = countClones; cloneNumber > 0; cloneNumber--) {
                const rowDest = this.worksheet.getRow(rowSrcNumber + countRows * cloneNumber);
                this.copyRow(rowSrc, rowDest);
            }
        }
    }

    /**
     * @param {Cell} cell
     * @return {CellRange}
     */
    getMergeRange(cell) {
        if (cell.isMerged && Array.isArray(this.worksheet.model['merges'])) {
            const address = cell.type === Excel.ValueType.Merge ? cell.master.address : cell.address;
            const cellRangeStr = this.worksheet.model['merges']
                .find(item => item.indexOf(address + ':') !== -1);
            if (cellRangeStr) {
                const [cellTlAdr, cellBrAdr] = cellRangeStr.split(':', 2);
                return CellRange.createFromCells(
                    this.worksheet.getCell(cellTlAdr),
                    this.worksheet.getCell(cellBrAdr)
                );
            }
        }
        return null;
    }

    /**
     * @param {Row} rowSrc
     * @param {Row} rowDest
     */
    moveRow(rowSrc, rowDest) {
        this.copyRow(rowSrc, rowDest);
        this.clearRow(rowSrc);
    }

    /**
     * @param {Row} rowSrc
     * @param {Row} rowDest
     */
    copyRow(rowSrc, rowDest) {
        /** @var {RowModel} */
        const rowModel = _.cloneDeep(rowSrc.model);
        rowModel.number = rowDest.number;
        rowModel.cells = [];
        rowDest.model = rowModel;

        const lastCol = this.getSheetDimension().right;
        for (let colNumber = lastCol; colNumber > 0; colNumber--) {
            const cell = rowSrc.getCell(colNumber);
            const newCell = rowDest.getCell(colNumber);
            this.copyCell(cell, newCell);
        }
    }

    /**
     * @param {Cell} cellSrc
     * @param {Cell} cellDest
     */
    copyCell(cellSrc, cellDest) {
        // skip submerged cells
        if (cellSrc.isMerged && (cellSrc.type === Excel.ValueType.Merge)) {
            return;
        }

        /** @var {CellModel} */
        const storeCellModel = _.cloneDeep(cellSrc.model);
        storeCellModel.address = cellDest.address;

        // Move a merge range
        const cellRange = this.getMergeRange(cellSrc);
        if (cellRange) {
            cellRange.move(cellDest.row - cellSrc.row, cellDest.col - cellSrc.col);
            this.worksheet.unMergeCells(cellRange.top, cellRange.left, cellRange.bottom, cellRange.right);
            this.worksheet.mergeCells(cellRange.top, cellRange.left, cellRange.bottom, cellRange.right);
        }

        // Move an image
        this.worksheet.getImages().forEach(image => {
            const rng = image.range;
            if (rng.tl.row <= cellSrc.row && rng.br.row >= cellSrc.row &&
                rng.tl.col <= cellSrc.col && rng.br.col >= cellSrc.col) {
                rng.tl.row += cellDest.row - cellSrc.row;
                rng.br.row += cellDest.row - cellSrc.row;
            }

        });

        cellDest.model = storeCellModel;
    }

    /**
     * @param {Row} row
     */
    clearRow(row) {
        // noinspection JSValidateTypes
        row.model = {number: row.number, cells: []};
    }

    /**
     * @param {Cell} cell
     */
    clearCell(cell) {
        // noinspection JSValidateTypes
        cell.model = {address: cell.address};
    }

    /**
     * @return {CellRange}
     */
    getSheetDimension() {
        const dm = this.worksheet.dimensions['model'];
        return new CellRange(dm.top, dm.left, dm.bottom, dm.right);
    }

    /**
     * Iterate cells from the left of the top to the right of the bottom
     * @param {CellRange} cellRange
     * @param {iterateCells} callBack
     */
    eachCell(cellRange, callBack) {
        for (let r = cellRange.top; r <= cellRange.bottom; r++) {
            const row = this.worksheet.findRow(r);
            for (let c = cellRange.left; c <= cellRange.right; c++) {
                const cell = row.findCell(c);
                if (cell && cell.type !== Excel.ValueType.Merge) {
                    if (callBack(cell) === false) {
                        return;
                    }
                }
            }
        }
    }

    /**
     * Iterate cells from the right of the bottom to the top of the left
     * @param {CellRange} cellRange
     * @param {iterateCells} callBack
     */
    eachCellReverse(cellRange, callBack) {
        for (let r = cellRange.bottom; r >= cellRange.top; r--) {
            const row = this.worksheet.findRow(r);
            for (let c = cellRange.right; c >= cellRange.left; c--) {
                const cell = row.findCell(c);
                if (cell && cell.type !== Excel.ValueType.Merge) {
                    if (callBack(cell) === false) {
                        return;
                    }
                }
            }
        }
    }
}

module.exports = WorkSheetHelper;