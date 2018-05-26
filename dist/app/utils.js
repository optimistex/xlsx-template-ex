"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const fs = require("fs");
const template_engine_1 = require("./template-engine");
const worksheet_helper_1 = require("./worksheet-helper");
const exceljs_1 = require("exceljs");
function xlsxBuildByTemplate(data, templateFileName) {
    if (!fs.existsSync(templateFileName)) {
        return Promise.reject(`File ${templateFileName} does not exist`);
    }
    if (typeof data !== "object") {
        return Promise.reject('The data must be an object');
    }
    const workbook = new exceljs_1.Workbook();
    return workbook.xlsx.readFile(templateFileName).then(() => {
        workbook.worksheets.forEach((worksheet) => {
            const wsh = new worksheet_helper_1.WorkSheetHelper(worksheet);
            const templateEngine = new template_engine_1.TemplateEngine(wsh, data);
            templateEngine.execute();
        });
        return workbook.xlsx.writeBuffer();
    });
}
exports.xlsxBuildByTemplate = xlsxBuildByTemplate;
