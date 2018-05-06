const xlsx = require('xlsx-populate');

/**
 * Build an xlsx document from a template
 * @param templateFileName - the template file
 * @param data - a source data
 * @returns {Promise<Buffer>}
 */
exports.xlsxBuildByTemplate = (data, templateFileName = './xlsx.helper.template.xlsx') => {
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


