# Excel шаблонизатор

Документация доступна на языках:
1. [English](https://github.com/optimistex/xlsx-template-ex#readme)
2. [Русский](README.ru.md)

Совместимые файлы: **xlsx**

Данный компонент реализует создание отчетов на основе шаблона.
Он имеет гибкий синтаксис, подобный шаблонным выражениям в Angular.

## Синтаксис

Поддерживаются 2 типа выражений:
* Вывод значения: `{{value|pipe:arg1:argN}}`
* Вывод массива: `[[value|pipe:arg1:argN]]`

Где:
* `value` - имя свйства которое содержит некоторое значение или массив
* `pipe` - некоторая функция дополнительной обработки значения
* `arg1`, `argN` - аргументы/параметры передаваемые в функцию

## Реализованные варианты выражений

* `{{propertyName}}` - вывод значения как есть
* `{{propertyName|date}}` - выводимое значение будет отформатировано как дата 
* `{{fileName|image}}` - поиск файла изображения по имени файла.
    Если изображение найдено, то оно встраивается в ячейку таблицы.
* `{{propertyArrayName|find:propertyName:propertyValue}}` - поиск значения в массиве `propertyArrayName` 
    который содержит свойство `propertyName` равное `propertyValue`
* `{{propertyObjectName|get:propertyName}}` - возвращает значение свойства `propertyArrayName` объекта `propertyObjectName`     

* `[[array|repeat-rows:3]]` - обработка массива значений и вывод содержимого
    в секцию из 3 строк начиная с текущей.
    Строки будут продублированы в соответствии с размером массива.
* `[[array|filter:propertyName:checkValue]]` - фильтр массива. 
    Если предоставлено только `propertyName`, тогда получим массив объектов содержащих это свойство.
    Если предоставлено `propertyName` и `checkValue`, тогда получим массив объектов которые содержат свойство 
    `propertyName` со значением `checkValue`.
* `[[array|tile:blockRows:blockColumns:tileColumns]]` - обработка массива значений и вывод по блокам.
    Исходный блок определяется по `blockRows` и `blockColumns`. 
    Данный блок будет выведен в таблице из `tileColumns` колонок. 
    
## Примеры

Будем выводить такие данные:
```javascript
let data = {
    reportBuildDate: 1526443275041,

    results: [
        { text: 'some text 1', answerText: 'a text of an answer 1'},
        { text: 'some text 2', answerText: 'a text of an answer 2'},
        { text: 'some text 3', answerText: 'a text of an answer 3'},
        { answerText: 'a text of an answer 3'},
    ],
};
```
    
Давайте сделаем шаблон:

**!!!** В данном примере / использовано вместо | из-за проблемы с синтаксисом Markdown 

| A | B |
|---|---|
|{{reportBuildDate/date}}| {{results/find:text:some text 2/get:answerText}} |
|[[results/filter:text/repeat-rows:1]] {{text}}| {{answerText}} |
| | |
| [[results/filter:text/tile:1:1:2]]{{answerText}} | |

Получим результат:

| A     | B     |
|-------|-------|
| 16.05.2018 | a text of an answer 2 |
| some text 1 | a text of an answer 1 |
| some text 2 | a text of an answer 2 |
| some text 3 | a text of an answer 3 |
| | |
| a text of an answer 1 | a text of an answer 2 |
| a text of an answer 3 | |

## Начало работы

```javascript
const fs = require("fs");
const XlsxTemplate = require('xlsx-template-ex');

const data = {
    reportBuildDate: 1526443275041,

    results: [
        { text: 'some text 1', answerText: 'a text of an answer 1'},
        { text: 'some text 2', answerText: 'a text of an answer 2'},
        { text: 'some text 3', answerText: 'a text of an answer 3'},
        { answerText: 'a text of an answer 3'},
    ],
};

XlsxTemplate.xlsxBuildByTemplate(data, 'template-file.xlsx')
    .then((buffer) => fs.writeFileSync('./out.xlsx', buffer))
    .catch((error) => console.log('xlsxHelper error:', error));
```

## Устранение неисправностей

Пожалуйста, следуйте данному руководству при сообщении о багах и запросе новых фич:

1. Используйте [GitHub Issues](https://github.com/optimistex/xlsx-template-ex/issues) раздел для отчетов об ошибках и запросах на новые фичи (не наш eMail)
2. Пожалуйста **всегда** описывайте шаги воспроизведения ошибки. В таком случае мы сможем сфокусироваться на решении проблемы, не ломая голову в попытках воспроизвести проблему.

Спасибо за понимание!

## Поддержка

- `npm start` - Запуск демо из сборки для релиза.
- `npm run start-dev` - Запуск демо в режиме разработки.
- `npm run build` - Сборка для релиза.

# Лицензия

Свободное использование (подробнее [Лицензия](https://github.com/optimistex/xlsx-template-ex/blob/master/LICENSE) в полнотекстовом файле)
