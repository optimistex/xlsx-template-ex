var fs = require('fs');
var xlsx = require('xlsx-populate');


var data = {
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


xlsx.fromFileAsync('test.xlsx').then(function(wb) {
    for(name in data) {
        if(typeof data[name] == 'string') {
            wb.find(`%${name}%`, data[name]);
        } else if(typeof data[name] == 'object') {
            var curSheet = wb.find(`%${name}:begin%`)[0].row().sheet();

            var firstRowNum = wb.find(`%${name}:begin%`)[0].row().rowNumber() + 1;

            var lastRowNum = wb.find(`%${name}:end%`)[0].row().rowNumber();
            var repeatOffset = lastRowNum - firstRowNum;
            var totalObjCnt = data[name].length;
            var totalOffset = (totalObjCnt-1) * repeatOffset;
            var maxColumn = curSheet.usedRange()._maxColumnNumber;
            var maxRow = curSheet.usedRange()._maxRowNumber;
            for(var r = maxRow; r > lastRowNum; r--) {
                for(var c = 1; c <= maxColumn; c++) {
                    var curVal = curSheet.row(r).cell(c).value();
                    var curStyle = curSheet.row(r).cell(c)._styleId;
                    curSheet.row(r + totalOffset).cell(c).value(curVal);
                    curSheet.row(r + totalOffset).cell(c)._styleId = curStyle;
                }
            }
            for(var i = 0; i < totalObjCnt; i++) {
                var curFirstRowNum = firstRowNum+i*repeatOffset;
                var curLastRowNum = lastRowNum+i*repeatOffset;
                var defRcnt = 0;
                for(var r = curFirstRowNum; r < curLastRowNum; r++) {
                    for(var c = 1; c <= maxColumn ; c++) {
                        var tpltVal = curSheet.row(firstRowNum+defRcnt).cell(c).value();
                        var tpltStyleId = curSheet.row(firstRowNum + defRcnt).cell(c)._styleId;
                        if(tpltVal !== undefined) {
                            if((/%.*\..*%/gi).test(tpltVal)) {
                                var newTemplated = '';
                                if (i === 0) {
                                    newTemplated = tpltVal;
                                } else {
                                    newTemplated = `%${tpltVal.slice(1, -1)}.${i}%`;
                                    var re = new RegExp(`%${name}\.[^%]*%`, 'gi');
                                    newTemplated = tpltVal.replace(re, function(match) {
                                        return `%${match.slice(1,-1)}.${i}%`;
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
            for(var i = 0; i < totalObjCnt; i++) {
                for(var attr in data[name][i]) {
                    if (i === 0) {
                        wb.find(`%${name}.${attr}%`, data[name][i][attr]);
                    } else {
                        wb.find(`%${name}.${attr}.${i}%`, data[name][i][attr]);
                    }
                }
            }
            wb.find(`%${name}:begin%`, '');
        }
    }
     return wb.outputAsync('buffer')
        .then(function(data) {
            fs.writeFileSync('./out.xlsx', data);
        });
});

//var tpltFile = fs.readFileSync('./test.xlsx');




/*
var wb = xlsx.readFile('test.xlsx');
//console.log(wb);
xlsx.writeFile(wb, 'out.xlsx');
*/

/*
var ws = xlsx.parse(tpltFile);
ws[0].data[1].push('OLOLOLOLOLO');
console.log(ws[0].data[1]);
fs.writeFileSync('./out.xlsx', xlsx.build(ws));
*/

/*
var tpltUnzip = new zip();
tpltUnzip.loadAsync(tpltFile).then(function() {
  //console.log(tpltUnzip);
  tpltUnzip.file('xl/worksheets/sheet1.xml').async('string').then(function(a) {
    var out = a.replace('<? test ?>', '1234567890');
    fs.writeFileSync('./test.xml', out);
    tpltUnzip.file('xl/worksheets/sheet1.xml', out);
    tpltUnzip.generateNodeStream({type:'nodebuffer',streamFiles:true})
      .pipe(fs.createWriteStream('out.xlsx'))
    //fs.writeFileSync('./out.xlsx', );
  });
});*/