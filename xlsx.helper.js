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
     */
    processBlocks(cellRange, data) {
        let restart;
        do {
            restart = false;
            this.wsh.eachCellReverse(cellRange, (cell) => {
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
                    this.processBlockPipes(tplExp.pipes, data[tplExp.valueName]);
                });

                restart = true;
                return false;
            });
        } while (restart);
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
     * @param {Array<{pipeName: string, pipeParameters: string[]}>} pipes
     * @param {object} data
     */
    processBlockPipes(pipes, data) {
        // console.log('bp', pipes, data);
        pipes.forEach(pipe => {
            switch (pipe.pipeName) {
                case 'repeat-rows':
                    this.blockPipeRepeatRows.apply(this, [data].concat(pipe.pipeParameters));
                    break;
                case 'block':
                    this.blockPipeBlock.apply(this, [data].concat(pipe.pipeParameters));
                    break;
                case 'tile':

                    break;
            }
        });
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
        console.log('+++', fileName);

        this.wsh.addImage(fileName, cell);

        return fileName;
    }

    /**
     * @param {object[]} dataArray
     * @param {number} rowBeginAddr
     * @param {number} rowEndAddr
     */
    blockPipeRepeatRows(dataArray, rowBeginAddr, rowEndAddr) {
        if (!Array.isArray(dataArray) || !dataArray.length) {
            console.warn('The data must be array', dataArray);
            return;
        }
        const cellBegin = this.wsh.worksheet.getCell(rowBeginAddr);
        const cellEnd = this.wsh.worksheet.getCell(rowEndAddr);
        this.wsh.cloneRows(cellBegin.row, cellEnd.row, dataArray.length - 1);

        const dRow = cellEnd.row - cellBegin.row + 1;
        const wsDimension = this.wsh.getSheetDimension();
        const sectionRange = new CellRange(cellBegin.row, wsDimension.left, cellEnd.row, wsDimension.right);

        dataArray.forEach(data => {
            this.processValues(sectionRange, data);
            sectionRange.move(dRow, 0);
        });
    }

    /**
     * @param {object[]} data
     * @param {string} addrTopLeft
     * @param {string} addrBottomRight
     */
    blockPipeBlock(data, addrTopLeft, addrBottomRight) {
        console.log('TemplateEngine.blockPipeBlock', data, addrTopLeft, addrBottomRight);
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