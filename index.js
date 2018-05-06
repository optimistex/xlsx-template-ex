const fs = require('fs');
const xlsxHelper = require('./xlsx.helper');
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

xlsxHelper.xlsxBuildByTemplate('test.xlsx', testData).then((buffer) => {
    fs.writeFileSync('./out.xlsx', buffer);
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