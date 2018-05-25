const fs = require('fs');
const Excel = require('exceljs');

const CellRange = require('./app/cell-range').CellRange;
const TemplateEngine = require('./app/template-engine').TemplateEngine;
const WorkSheetHelper = require('./app/worksheet-helper').WorkSheetHelper;
const TemplateExpression = require('./app/template-expression').TemplateExpression;

module.exports.CellRange = CellRange;
module.exports.TemplateEngine = TemplateEngine;
module.exports.WorkSheetHelper = WorkSheetHelper;
module.exports.TemplateExpression = TemplateExpression;

module.exports.xlsxBuildByTemplate = (data, templateFileName) => {
    if (!fs.existsSync(templateFileName)) {
        return Promise.reject(`File ${templateFileName} does not exist`);
    }
    if (typeof data !== "object") {
        return Promise.reject('The data must be an object');
    }

    const workbook = new Excel.Workbook();
    return workbook.xlsx.readFile(templateFileName).then(() => {

        workbook.worksheets.forEach(worksheet => {
            const wsh = new WorkSheetHelper(worksheet);
            const templateEngine = new TemplateEngine(wsh, data);
            templateEngine.execute();
        });

        return workbook.xlsx.writeBuffer();
    });
};
