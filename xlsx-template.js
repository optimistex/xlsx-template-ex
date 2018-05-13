const fs = require('fs');
const Excel = require('exceljs');

/**
 * @property {Object} data
 * @property {string} templateFileName
 */
// class XlsxTemplate {
//
//     /**
//      * @param {Object} data
//      * @param {string} templateFileName
//      */
//     constructor(data, templateFileName) {
//         this.data = data;
//         this.templateFileName = templateFileName;
//
//     }
//
//     /** @return {boolean} */
//     validate() {
//         this.errors = [];
//         if (typeof this.data !== "object") {
//             this.errors.push('The data must be an object');
//         }
//         if (!fs.existsSync(this.templateFileName)) {
//             this.errors.push(`File ${this.templateFileName} does not exist`);
//         }
//         return !this.errors.length;
//     }
//
//     /** @return {Promise<Buffer>} */
//     writeBuffer() {
//         if (!this.validate()) {
//             return Promise.reject(this.errors);
//         }
//
//         this.workbook = new Excel.Workbook();
//         return this.workbook.xlsx.readFile(this.templateFileName)
//             .catch((error) => {
//                 console.log('1111111111', error);
//             })
//             .then(() => {
//                 // use workbook
//                 // console.log(workbook);
//
//                 const worksheet = this.workbook.worksheets[0];
//
//                 const r7 = worksheet.getRow(7);
//
//                 worksheet.spliceRows(10, 0, [], r7);
//
//                 // worksheet.getCell('B7').value = '1111111111';
//
//
//                 const imgId = this.workbook.addImage({filename: __dirname + '/152.jpg', extension: 'jpeg'});
//                 worksheet.addImage(imgId, {
//                     tl: {col: 1.1, row: 16.1},
//                     br: {col: 2.0, row: 18.0},
//                     // editAs: 'oneCell'
//                 });
//
//                 return this.workbook.xlsx.writeBuffer();
//             });
//     }
//
//
// }

module.exports.xlsxBuildByTemplate = (data, templateFileName) => {
    if (!templateFileName) {
        templateFileName = __dirname + '/xlsx.helper.template.xlsx';
    }
    if (!data) {
        return Promise.reject('Undefined data');
    }

    // read from a file
    const workbook = new Excel.Workbook();
    return workbook.xlsx.readFile(templateFileName)
        .catch((error) => {
            console.log('1111111111', error);
        })
        .then(function () {
            // use workbook
            // console.log(workbook);
            const imgId = workbook.addImage({filename: __dirname + '/152.jpg', extension: 'jpeg'});

            const worksheet = workbook.worksheets[0];
            worksheet.addImage(imgId, {
                tl: {col: 1.1, row: 16.1},
                br: {col: 2.0, row: 18.0},
                // editAs: 'oneCell'
            });

            return workbook.xlsx.writeBuffer();
        });
};
