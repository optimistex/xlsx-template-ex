const fs = require('fs');
const _ = require('lodash');
const Excel = require('exceljs');

module.exports.xlsxBuildByTemplate2 = (data, templateFileName) => {
    if (!fs.existsSync(templateFileName)) {
        return Promise.reject(`File ${templateFileName} does not exist`);
    }
    if (typeof data !== "object") {
        return Promise.reject('The data must be an object');
    }

    const workbook = new Excel.Workbook();
    return workbook.xlsx.readFile(templateFileName).then(() => {
        const wsh = new WorkSheetHelper(workbook.worksheets[0]);

        const templateEngine = new TemplateEngine(wsh, data);
        templateEngine.execute();

        return workbook.xlsx.writeBuffer();
    });
};

/**
 * @property {string} rawExpression
 * @property {string} expression
 * @property {string} valueName
 * @property {Array<{pipeName: string, pipeParameters: string[]}>} pipes
 */
class TemplateExpression {
    /**
     * @param {string} rawExpression
     * @param {string} expression
     */
    constructor(rawExpression, expression) {
        this.rawExpression = rawExpression;
        this.expression = expression;
        const expressionParts = this.expression.split('|');
        this.valueName = expressionParts[0];
        this.pipes = [];
        const pipes = expressionParts.slice(1);
        pipes.forEach(pipe => {
            const pipeParts = pipe.split(':');
            this.pipes.push({pipeName: pipeParts[0], pipeParameters: pipeParts.slice(1)});
        });
    }
}

/**
 * @property {WorkSheetHelper} wsh
 * @property {object} data
 */
class TemplateEngine {
    /**
     * @param {WorkSheetHelper} wsh
     * @param {object} data
     */
    constructor(wsh, data) {
        this.wsh = wsh;
        this.data = data;
        // noinspection RegExpRedundantEscape
        this.regExpBlocks = /\[\[.+?\]\]/g;
        this.regExpValues = /{{.+?}}/g;
    }

    execute() {
        this.processBlocks(this.wsh.getSheetDimension(), this.data);
        this.processValues(this.wsh.getSheetDimension(), this.data);
    }

    /**
     * @param {CellRange} cellRange
     * @param {object} data
     * @return {CellRange} the new range
     */
    processBlocks(cellRange, data) {
        let restart;
        do {
            restart = false;
            this.wsh.eachCell(cellRange, (cell) => {
                let cVal = cell.value;
                if (typeof cVal !== "string") {
                    return null;
                }
                const matches = cVal.match(this.regExpBlocks);
                if (!Array.isArray(matches) || !matches.length) {
                    return null;
                }

                matches.forEach(rawExpression => {
                    const tplExp = new TemplateExpression(rawExpression, rawExpression.slice(2, -2));
                    cVal = cVal.replace(tplExp.rawExpression, '');
                    cell.value = cVal;
                    cellRange = this.processBlockPipes(cellRange, cell, tplExp.pipes, data[tplExp.valueName]);
                });

                restart = true;
                return false;
            });
        } while (restart);
        return cellRange;
    }

    /**
     * @param {CellRange} cellRange
     * @param {object} data
     */
    processValues(cellRange, data) {
        this.wsh.eachCell(cellRange, (cell) => {
            let cVal = cell.value;
            if (typeof cVal !== "string") {
                return;
            }
            const matches = cVal.match(this.regExpValues);
            if (!Array.isArray(matches) || !matches.length) {
                return;
            }

            matches.forEach(rawExpression => {
                const tplExp = new TemplateExpression(rawExpression, rawExpression.slice(2, -2));
                let resultValue = data[tplExp.valueName] || '';
                resultValue = this.processValuePipes(cell, tplExp.pipes, resultValue);
                cVal = cVal.replace(tplExp.rawExpression, resultValue);
            });
            cell.value = cVal;
        });
    }

    /**
     * @param {Cell} cell
     * @param {Array<{pipeName: string, pipeParameters: string[]}>} pipes
     * @param {string} value
     * @return {string}
     */
    processValuePipes(cell, pipes, value) {
        pipes.forEach(pipe => {
            switch (pipe.pipeName) {
                case 'date':
                    value = this.valuePipeDate.apply(this, [value].concat(pipe.pipeParameters));
                    break;
                case 'image':
                    value = this.valuePipeImage.apply(this, [cell, value].concat(pipe.pipeParameters));
                    // value = 'todo: past image'; //todo: past image
                    break;
            }
        });
        return value;
    }

    /**
     * @param {CellRange} cellRange
     * @param {Cell} cell
     * @param {Array<{pipeName: string, pipeParameters: string[]}>} pipes
     * @param {object} data
     * @return {CellRange} the new cell range
     */
    processBlockPipes(cellRange, cell, pipes, data) {
        // console.log('bp', pipes, data);
        const newRange = CellRange.createFromRange(cellRange);
        pipes.forEach(pipe => {
            switch (pipe.pipeName) {
                case 'repeat-rows':
                    const insertedRows = this.blockPipeRepeatRows.apply(this, [cell, data].concat(pipe.pipeParameters));
                    newRange.bottom += insertedRows;
                    break;
            }
        });
        return newRange;
    }

    /**
     * @param {number|string} date
     * @return {string}
     */
    valuePipeDate(date) {
        return date ? (new Date(date)).toLocaleDateString() : '';
    }

    /**
     * @param {Cell} cell
     * @param {string} fileName
     * @return {string}
     */
    valuePipeImage(cell, fileName) {
        if (fs.existsSync(fileName)) {
            this.wsh.addImage(fileName, cell);
            return fileName;
        }
        return `File "${fileName}" not found`;
    }

    /**
     * @param {Cell} cell
     * @param {object[]} dataArray
     * @param {number} countRows
     * @return {number} count of inserted rows
     */
    blockPipeRepeatRows(cell, dataArray, countRows) {
        if (!Array.isArray(dataArray) || !dataArray.length) {
            console.warn('The data must be array, but got:', dataArray);
            return 0;
        }
        countRows = +countRows > 0 ? +countRows : 1;
        const startRow = cell.row;
        const endRow = startRow + countRows - 1;
        if (dataArray.length > 1) {
            this.wsh.cloneRows(startRow, endRow, dataArray.length - 1);
        }

        const wsDimension = this.wsh.getSheetDimension();
        let sectionRange = new CellRange(startRow, wsDimension.left, endRow, wsDimension.right);

        dataArray.forEach(data => {
            sectionRange = this.processBlocks(sectionRange, data);
            this.processValues(sectionRange, data);
            sectionRange.move(countRows, 0);
        });
        return (dataArray.length - 1) * countRows;
    }
}

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