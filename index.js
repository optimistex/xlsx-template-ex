const fs = require('fs');
const Excel = require('exceljs');

const CellRange = require('./src/cell-range');
const TemplateEngine = require('./src/template-engine');
const WorkSheetHelper = require('./src/worksheet-helper');
const TemplateExpression = require('./src/template-expression');

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
        const wsh = new WorkSheetHelper(workbook.worksheets[0]);

        const templateEngine = new TemplateEngine(wsh, data);
        templateEngine.execute();

        return workbook.xlsx.writeBuffer();
    });
};
