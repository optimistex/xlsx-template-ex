const fs = require('fs');
const xlsx = require('xlsx-populate');

function xlsxBuildByTemplate(templateFileName, data) {
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
}

const testData = {
    'name': 'Отчёт',
    'date': '01.05.2018',
    'users': [
        {
            'name': 'Иванов Пётр Семёнович',
            'code': '00000001'
        },
        {
            'name': 'Прокофьев Павел Сергеевич',
            'code': '00000002'
        },
        {
            'name': 'Аллаберганов Мадиёр Фарходович',
            'code': '00000003'
        }
    ],
    'orders': [
        {
            'number': '1',
            'date': '25.04.2018',
            'position': 'Хлеб',
            'pts': '10'
        },
        {
            'number': '2',
            'date': '26.04.2018',
            'position': 'Кефир',
            'pts': '5'
        },
        {
            'number': '2',
            'date': '27.04.2018',
            'position': 'Колбаса',
            'pts': '2'
        }
    ]
};

xlsxBuildByTemplate('test.xlsx', testData).then(function (data) {
    fs.writeFileSync('./out.xlsx', data);
});

//let tpltFile = fs.readFileSync('./test.xlsx');


/*
let wb = xlsx.readFile('test.xlsx');
//console.log(wb);
xlsx.writeFile(wb, 'out.xlsx');
*/

/*
let ws = xlsx.parse(tpltFile);
ws[0].data[1].push('OLOLOLOLOLO');
console.log(ws[0].data[1]);
fs.writeFileSync('./out.xlsx', xlsx.build(ws));
*/

/*
let tpltUnzip = new zip();
tpltUnzip.loadAsync(tpltFile).then(function() {
  //console.log(tpltUnzip);
  tpltUnzip.file('xl/worksheets/sheet1.xml').async('string').then(function(a) {
    let out = a.replace('<? test ?>', '1234567890');
    fs.writeFileSync('./test.xml', out);
    tpltUnzip.file('xl/worksheets/sheet1.xml', out);
    tpltUnzip.generateNodeStream({type:'nodebuffer',streamFiles:true})
      .pipe(fs.createWriteStream('out.xlsx'))
    //fs.writeFileSync('./out.xlsx', );
  });
});*/