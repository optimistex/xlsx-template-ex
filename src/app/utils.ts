import * as fs from "fs";
import {TemplateEngine} from "./template-engine";
import {WorkSheetHelper} from "./worksheet-helper";
import {Workbook, Worksheet, Buffer} from "exceljs";

export function xlsxBuildByTemplate(data: any, templateFileName: string): Promise<Buffer> {
  if (!fs.existsSync(templateFileName)) {
    return Promise.reject(`File ${templateFileName} does not exist`);
  }
  if (typeof data !== "object") {
    return Promise.reject('The data must be an object');
  }

  const workbook = new Workbook();
  return workbook.xlsx.readFile(templateFileName).then(() => {

    workbook.worksheets.forEach((worksheet: Worksheet) => {
      const wsh = new WorkSheetHelper(worksheet);
      const templateEngine = new TemplateEngine(wsh, data);
      templateEngine.execute();
    });

    return workbook.xlsx.writeBuffer();
  });
}
