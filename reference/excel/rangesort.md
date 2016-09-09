# Объект RangeSort (API JavaScript для Excel)

_Относится к: Excel 2016, Excel Online, Excel для iOS, Office 2016_

Управляет операциями сортировки для объектов Range.

## Свойства

Нет

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)](#applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string)|void|Выполняет сортировку.|

## Сведения о методе


### apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)
Выполняет сортировку.

#### Синтаксис
```js
rangeSortObject.apply(fields, matchCase, hasHeaders, orientation, method);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|fields|SortField[]|Список условий для сортировки.|
|matchCase|Bool|Необязательный параметр. Указывает, необходимо ли учитывать регистр при сортировке строк.|
|hasHeaders|Bool|Необязательный параметр. Указывает, есть ли у диапазона заголовок.|
|orientation|string|Необязательный параметр. Указывает ориентацию сортировки — строки или столбцы.  Возможные значения: Rows, Columns|
|method|string|Необязательный. Метод сортировки, используемый для китайских символов.  Возможные значения: PinYin, StrokeCount|

#### Возвращаемое значение
void

#### Примеры
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D4:G6";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.sort.apply([ 
            {
                key: 2,
                ascending: true
            },
        ], true);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```