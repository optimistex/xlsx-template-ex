import * as fs from "fs";
import * as Excel from "exceljs";
import {TemplateEngine} from "./template-engine";
import {WorkSheetHelper} from "./worksheet-helper";

export function xlsxBuildByTemplate(data, templateFileName) {
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
}
