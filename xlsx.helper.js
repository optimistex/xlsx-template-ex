const fs = require('fs');

const _ = require('lodash');
const xlsx = require('xlsx-populate');
const Excel = require('exceljs');

/**
 * Build an xlsx document from a template
 * @param templateFileName - the template file
 * @param data - a source data
 * @returns {Promise<Buffer>}
 */
module.exports.xlsxBuildByTemplate = (data, templateFileName = null) => {
    if (!templateFileName) {
        templateFileName = __dirname + '/xlsx.helper.template.xlsx';
    }
    if (!data) {
        return Promise.reject('Undefined data');
    }
    return xlsx.fromFileAsync(templateFileName).then(function (wb) {
        for (let name in data) {
            if (data.hasOwnProperty(name)) {
                if (typeof data[name] === 'string') {
                    wb.find(`%${name}%`, data[name]);
                } else if (typeof data[name] === 'object') {
                    const curSheet = wb.find(`%${name}:begin%`)[0].row().sheet();

                    const firstRowNum = wb.find(`%${name}:begin%`)[0].row().rowNumber() + 1;

                    const lastRowNum = wb.find(`%${name}:end%`)[0].row().rowNumber();
                    const repeatOffset = lastRowNum - firstRowNum;
                    const totalObjCnt = data[name].length;
                    const totalOffset = (totalObjCnt - 1) * repeatOffset;
                    const maxColumn = curSheet.usedRange()._maxColumnNumber;
                    const maxRow = curSheet.usedRange()._maxRowNumber;
                    for (let r = maxRow; r > lastRowNum; r--) {
                        for (let c = 1; c <= maxColumn; c++) {
                            const curVal = curSheet.row(r).cell(c).value();
                            const curStyle = curSheet.row(r).cell(c)._styleId;
                            curSheet.row(r + totalOffset).cell(c).value(curVal);
                            curSheet.row(r + totalOffset).cell(c)._styleId = curStyle;
                        }
                    }
                    for (let i = 0; i < totalObjCnt; i++) {
                        const curFirstRowNum = firstRowNum + i * repeatOffset;
                        const curLastRowNum = lastRowNum + i * repeatOffset;
                        let defRcnt = 0;
                        for (let r = curFirstRowNum; r < curLastRowNum; r++) {
                            for (let c = 1; c <= maxColumn; c++) {
                                const tpltVal = curSheet.row(firstRowNum + defRcnt).cell(c).value();
                                const tpltStyleId = curSheet.row(firstRowNum + defRcnt).cell(c)._styleId;
                                let newTemplated = '';
                                if (tpltVal !== undefined) {
                                    if ((/%.*\..*%/gi).test(tpltVal)) {
                                        newTemplated = '';
                                        if (i === 0) {
                                            newTemplated = tpltVal;
                                        } else {
                                            newTemplated = `%${tpltVal.slice(1, -1)}.${i}%`;
                                            let re = new RegExp(`%${name}\.[^%]*%`, 'gi');
                                            newTemplated = tpltVal.replace(re, function (match) {
                                                return `%${match.slice(1, -1)}.${i}%`;
                                            });
                                        }
                                    } else {
                                        newTemplated = tpltVal;
                                    }
                                } else {
                                    newTemplated = undefined;
                                }
                                curSheet.row(r).cell(c).value(newTemplated);
                                curSheet.row(r).cell(c)._styleId = tpltStyleId;

                            }
                            defRcnt++;
                        }
                    }
                    for (let i = 0; i < totalObjCnt; i++) {
                        for (let attr in data[name][i]) {
                            if (data[name][i].hasOwnProperty(attr)) {
                                if (i === 0) {
                                    wb.find(`%${name}.${attr}%`, data[name][i][attr]);
                                } else {
                                    wb.find(`%${name}.${attr}.${i}%`, data[name][i][attr]);
                                }
                            }
                        }
                    }
                    wb.find(`%${name}:begin%`, '');
                }
            }
        }
        return wb.outputAsync('buffer');
    });
};

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

        // wsh.cloneRows(34, 38, 5);
        // wsh.cloneRows(26, 28, 5);
        //
        // wsh.addImage(__dirname + '/alex.jpg', 29, 2);
        // wsh.addImage(__dirname + '/alex.jpg', 32, 2);
        // wsh.addImage(__dirname + '/alex.jpg', 35, 2);
        // wsh.addImage(__dirname + '/alex.jpg', 31, 3);
        // wsh.addImage(__dirname + '/alex.jpg', 35, 2);

        wsh.processTemplates(wsh.getSheetDimension(), data);

        return workbook.xlsx.writeBuffer();
    });
};

/**
 * @property {string} template
 * @property {object} data
 * @property {?string} replaceText
 */
class TemplateString {
    /**
     * @param {string} template
     * @param {object} data
     */
    constructor(template, data) {
        this.template = template;
        this.data = data;

        this.replaceText = null;
    }

    /**
     * @return {boolean}
     */
    parse() {
        const reg = new RegExp('{{.+?}}', 'g');
        const matches = this.template.match(reg);
        if (!Array.isArray(matches) || !matches.length) {
            return false;
        }

        this.replaceText = this.template;

        matches.forEach((rawExp) => {
            /** @const {string} expression */
            const expression = rawExp.slice(2, -2);
            /** @const {string[]} mainParts */
            const mainParts = expression.split('|');
            const valueName = mainParts[0];
            const pipes = mainParts.slice(1); //todo: make pipes

            this.replaceText = this.replaceText.replace(rawExp, this.data[valueName] || '');
        });

        return true;
    }
}

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
     * @param {number} row
     * @param {number} col
     */
    addImage(fileName, row, col) {
        const imgId = this.workbook.addImage({filename: fileName, extension: 'jpeg'});

        const cell = this.worksheet.getRow(row).getCell(col);
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
     * @param {number} srcRowStart
     * @param {number} srcRowEnd
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

    getSheetDimension() {
        const dm = this.worksheet.dimensions['model'];
        return new CellRange(dm.top, dm.left, dm.bottom, dm.right);
    }

    /**
     * @param {CellRange} cellRange
     * @param {object} data
     */
    processTemplates(cellRange, data) {
        for (let r = cellRange.bottom; r >= cellRange.top; r--) {
            const row = this.worksheet.findRow(r);
            for (let c = cellRange.right; c >= cellRange.left; c--) {
                const cell = row.findCell(c);
                if (cell) {
                    this.replaceTemplates(cell, data);
                }
            }
        }
    }

    /**
     * @param {Cell} cell
     * @param {object} data
     */
    replaceTemplates(cell, data) {
        const cVal = cell.value;
        if (typeof cVal !== "string") {
            return;
        }

        const te = new TemplateString(cVal, data);
        if (!te.parse()) {
            return;
        }

        if (te.replaceText !== null) {
            cell.value = te.replaceText;
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
     * @param {number} dRow
     * @param {number} dCol
     */
    move(dRow, dCol) {
        this.top += dRow;
        this.bottom += dRow;

        this.left += dCol;
        this.right += dCol;
    }
}