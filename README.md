# Excel Шаблонизатор

Поддерживаемые файлы: **xlsx**

## Синтаксис

Поддерживается 2 вида выражений:
* Вывод значения переменной: `{{value|pipe:arg1:argN}}`
* Вывод массива переменных: `[[value|pipe:arg1:argN]]`

Где:
* `value` - некоторое значение или массив
* `pipe` - некоторая функция дополнительной обработки значения
* `arg1`, `argN` - агрументы/параметры передаваемые в функцию обработки значения

## Реализованные варианты выражений

* `{{propertyName}}` - вывод значения без обработки
* `{{propertyName|date}}` - значение форматируется как дата 
* `{{fileName|image}}` - производится поиск картинки по имени файла. 
    Если картинка найдена, то она встраивается в ячейку таблицы
* `{{propertyArrayName|find:propertyName:propertyValue}}` - поиск объекта в массиве `propertyArrayName` 
    у которого есть свойство `propertyName` равное `propertyValue`
* `{{propertyObjectName|get:propertyName}}` - возвращает значение свойства `propertyArrayName` из объекта `propertyObjectName`     

* `[[array|repeat-rows:3]]` - обрабатать массив переменных и 
    вывести его содержимое в секцию из 3 строк начиная с текущей. 
    Строки будут продублированы в соответствии с размером массива.
    
## Примеры

Будем выводить в шаблонизаторе такие данные:
```javascript
let data = {
    reportBuildDate: 1526443275041,

    results: [
        { text: 'some text 1', answerText: 'a text of an answer 1'},
        { text: 'some text 2', answerText: 'a text of an answer 2'},
        { text: 'some text 3', answerText: 'a text of an answer 3'},
    ],
};
```
    
Составим шаблон:

**!!!** В данном примере / указано вместо |

| A | B |
|---|---|
|{{reportBuildDate/date}}| {{results/find:text:some text 2/get:answerText}} |
|[[results/repeat-rows:1]] {{text}}| {{answerText}} |

Получим результат:

| A     | B     |
|-------|-------|
| 16.05.2018 | a text of an answer 2 |
| some text 1 | a text of an answer 1 |
| some text 2 | a text of an answer 2 |
| some text 3 | a text of an answer 3 |
