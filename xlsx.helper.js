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
        // addImage(workbook, __dirname + '/alex.jpg', 25, 1);

        const wsh = new WorkSheetHelper(workbook.worksheets[0]);

        // insertRow(workbook, 10);
        // rowsCloneToBottom(workbook, 26, 28);

        wsh.cloneRows(25, 27, 4);
        // wsh.cloneRows(25, 27, 1);
        // wsh.cloneRows(25, 27, 1);
        // wsh.cloneRows(25, 27, 1);

        // rowsCloneToBottom(workbook, 19, 20);
        // rowsCloneToBottom(workbook, 19, 20);
        // rowsCloneToBottom(workbook, 19, 20);
        // rowsCloneToBottom(workbook, 19, 20);

        return workbook.xlsx.writeBuffer();
    });
};

function addImage(workbook, fileName, row, col) {
    const imgId = workbook.addImage({filename: fileName, extension: 'jpeg'});
    const worksheet = workbook.worksheets[0];
    const cell = worksheet.findCell(row, col);
    if (cell && cell.isMerged) {
        worksheet.mergeCells()
    }
    worksheet.addImage(imgId, {
        tl: {col: col, row: row},
        br: {col: col + 1, row: row + 1}
    });
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
     * @param {number} srcRowStart
     * @param {number} srcRowEnd
     * @param {number} countClones
     */
    cloneRows(srcRowStart, srcRowEnd, countClones = 1) {
        const dxRow = (srcRowEnd - srcRowStart + 1) * countClones;
        const lastRow = this.worksheet.dimensions['model'].bottom + dxRow;

        for (let rowNumber = lastRow; rowNumber > srcRowStart; rowNumber--) {
            // for (let rowNumber = lastRow; rowNumber > srcRowEnd; rowNumber--) {
            const rowSrc = this.worksheet.getRow(rowNumber);
            const rowDest = this.worksheet.getRow(rowNumber + dxRow);
            this.copyRow(rowSrc, rowDest);
        }
    }

    /**
     * @param {Cell} cell
     * @return {(string|null)} e.g. `'A4:B5'` or null
     */
    getMergeRange(cell) {
        if (cell.isMerged && Array.isArray(this.worksheet.model['merges'])) {
            const address = cell.type === Excel.ValueType.Merge ? cell.master.address : cell.address;
            return this.worksheet.model['merges'].find(item => {
                return item.indexOf(address + ':') !== -1;
            });
        }
        return null;
    }

    /**
     * @param {Row} rowSrc
     * @param {Row} rowDest
     */
    copyRow(rowSrc, rowDest) {
        // rowDest.height = rowSrc.height;

        /** @var {RowModel} */
        const rowModel = _.cloneDeep(rowSrc.model);
        rowModel.number = rowDest.number;
        rowModel.cells = [];
        rowDest.model = rowModel;

        const lastCol = this.worksheet.dimensions['model'].right;
        for (let colNumber = lastCol; colNumber > 0; colNumber--) {
            const cell = rowSrc.getCell(colNumber);
            const newCell = rowDest.getCell(colNumber);
            this.copyCell(cell, newCell);
        }

        this.clearRow(rowSrc);
    }

    /**
     * @param {Cell} cellSrc
     * @param {Cell} cellDest
     */
    copyCell(cellSrc, cellDest) {
        // skip submerged cells
        if (cellSrc.isMerged && (cellSrc.type === Excel.ValueType.Merge)) {
            // this.clearCell(cellDest);
            return;
        }

        /** @var {CellModel} */
        const storeCellModel = _.cloneDeep(cellSrc.model);
        storeCellModel.address = cellDest.address;

        // Move a merge range
        const mergeRangeStr = this.getMergeRange(cellSrc);
        if (mergeRangeStr) {
            const endRangeCell = this.worksheet.getCell(mergeRangeStr.split(':')[1]);
            const dR = cellDest.row - cellSrc.row, dC = cellDest.col - cellSrc.col;
            const mergeRange = {
                top: cellSrc.row + dR,
                left: cellSrc.col + dC,
                right: endRangeCell.col + dC,
                bottom: endRangeCell.row + dR
            };
            this.worksheet.unMergeCells(mergeRange.top, mergeRange.left, mergeRange.bottom, mergeRange.right);
            this.worksheet.mergeCells(mergeRange.top, mergeRange.left, mergeRange.bottom, mergeRange.right);
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
}
