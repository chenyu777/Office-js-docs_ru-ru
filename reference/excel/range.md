# Объект Range (API JavaScript для Excel)

_Относится к: Excel 2016, Excel Online, Excel для iOS, Office 2016_

Объект Range представляет собой набор из одной или нескольких смежных ячеек, например ячейку, строку, столбец, блок ячеек и т. д.

## Properties

| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|address|string|Представляет ссылку на диапазон в стиле A1. Значение адреса будет содержать ссылку на лист (например, Лист1!A1:B4). Только для чтения.|
|addressLocal|string|Представляет ссылку на указанный диапазон на языке пользователя. Только для чтения.|
|cellCount|int|Количество ячеек в диапазоне. Только для чтения.|
|columnCount|int|Представляет общее количество столбцов в диапазоне. Только для чтения.|
|columnHidden|bool|Указывает, скрыты ли все столбцы текущего диапазона.|
|columnIndex|int|Представляет номер столбца первой ячейки диапазона. Используется нулевой индекс. Только для чтения.|
|formulas|object[]|Представляет формулу в нотации стиля A1.|
|formulasLocal|object[][]|Представляет формулу в нотации стиля A1 на языке пользователя и в соответствии с его языковым стандартом. Например, английская формула =SUM(A1, 1.5) превратится в "=СУММ(A1; 1,5)" на русском языке.|
|formulasR1C1|object[][]|Представляет формулу в нотации стиля R1C1.|
|hidden|Bool|Указывает, скрыты ли все ячейки текущего диапазона. Только для чтения.|
|numberFormat|object[][]|Представляет код числового формата для данной ячейки.|
|rowCount|int|Возвращает общее количество строк в диапазоне. Только для чтения.|
|rowHidden|bool|Указывает, скрыты ли все строки текущего диапазона.|
|rowIndex|int|Возвращает номер строки первой ячейки диапазона. Используется нулевой индекс. Только для чтения.|
|text|object[][]|Текстовые значения указанного диапазона. Текстовое значение не зависит от ширины ячейки. Замена знака #, которая происходит в пользовательском интерфейсе Excel, не повлияет на текстовое значение, возвращаемое API-интерфейсом. Только для чтения.|
|valueTypes|string|Представляет тип данных каждой ячейки. Только для чтения. Возможные значения: Unknown, Empty, String, Integer, Double, Boolean, Error.|
|values|object[][]|Представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейка, которая содержит ошибку, возвращает строку ошибки.|

_Ознакомьтесь с [примерами](#примерами) доступа к свойствам._

## Связи
| Связь | Тип   |Описание|
|:---------------|:--------|:----------|
|format|[RangeFormat](rangeformat.md)|Возвращает объект формата, в который включены шрифт, заливка, границы, выравнивание и другие свойства диапазона. Только для чтения.|
|sort|[RangeSort](rangesort.md)|Представляет конфигурацию сортировки для диапазона. Только для чтения.|
|лист|[Таблица](worksheet.md)|Лист, содержащий текущий диапазон. Только для чтения.|

## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[clear(applyTo: string)](#clearapplyto-string)|void|Очищает значения, формат, заливку, границу диапазона и т. д.|
|[delete(shift: string)](#deleteshift-string)|void|Удаляет ячейки, связанные с диапазоном.|
|[getBoundingRect(anotherRange: Range или string)](#getboundingrectanotherrange-range-или-string)|[Range](range.md)|Возвращает наименьший объект диапазона, включающий в себя заданные диапазоны. Например, getBoundingRect для "B2:C5" и "D10:E15" возвращает значение "B2:E15".|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|Получает объект диапазона, содержащий одну ячейку, по номеру строки и столбца. Ячейка может находиться вне родительского диапазона, если она расположена в таблице листа. Возвращаемая ячейка располагается относительно верхней левой ячейки диапазона.|
|[getColumn(column: number)](#getcolumncolumn-number)|[Range](range.md)|Возвращает столбец в диапазоне.|
|[getEntireColumn()](#getentirecolumn)|[Range](range.md)|Возвращает объект, представляющий весь столбец диапазона.|
|[getEntireRow()](#getentirerow)|[Range](range.md)|Возвращает объект, представляющий всю строку диапазона.|
|[getIntersection(anotherRange: Range или string)](#getintersectionanotherrange-range-или-string)|[Range](range.md)|Возвращает объект диапазона, представляющий собой прямоугольное пересечение заданных диапазонов.|
|[getLastCell()](#getlastcell)|[Range](range.md)|Возвращает последнюю ячейку в диапазоне. Например, последняя ячейка диапазона B2:D5 — D5.|
|[getLastColumn()](#getlastcolumn)|[Range](range.md)|Возвращает последний столбец в диапазоне. Например, последний столбец диапазона B2:D5 — D2:D5.|
|[getLastRow()](#getlastrow)|[Range](range.md)|Возвращает последнюю строку в диапазоне. Например, последняя строка в диапазоне "B2:D5" — "B5:D5".|
|[getOffsetRange(rowOffset: number, columnOffset: number)](#getoffsetrangerowoffset-number-columnoffset-number)|[Range](range.md)|Возвращает объект, представляющий диапазон, который смещен от указанного диапазона. Измерение возвращаемого диапазона будет соответствовать этому диапазону. Если полученный диапазон выходит за пределы таблицы листа, вызывается исключение.|
|[getRow(row: number)](#getrowrow-number)|[Range](range.md)|Возвращает строку в диапазоне.|
|[getUsedRange(valuesOnly: bool)](#getusedrangevaluesonly-bool)|[Range](range.md)|Возвращает используемый поддиапазон объекта диапазона.|
|[insert(shift: string)](#insertshift-string)|[Range](range.md)|Вставляет ячейку или диапазон ячеек на лист вместо этого диапазона, а также сдвигает другие ячейки, чтобы освободить место. Возвращает новый объект Range в пустом месте.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|
|[merge(across: bool)](#mergeacross-bool)|void|Объединяет ячейки диапазона в одну область на листе.|
|[select()](#select)|void|Выбирает указанный диапазон в пользовательском интерфейсе Excel.|
|[unmerge()](#unmerge)|void|Разъединяет ячейки диапазона в отдельные ячейки.|

## Сведения о методе


### clear(applyTo: string)
Очищает значения, формат, заливку, границу диапазона и т. д.

#### Синтаксис
```js
rangeObject.clear(applyTo);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|applyTo|string|Необязательный параметр. Определяет тип действия очистки. Возможные значения: `All` Default-option, `Formats`, `Contents`.|

#### Возвращаемое значение
void

#### Примеры

Пример ниже очищает формат и содержимое диапазона. 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.clear();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### delete(shift: string)
Удаляет ячейки, связанные с диапазоном.

#### Синтаксис
```js
rangeObject.delete(shift);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|shift|string|Определяет способ сдвига ячеек. Возможные значения: Up, Left.|

#### Возвращаемое значение
void

#### Примеры

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getBoundingRect(anotherRange: Range или string)
Возвращает наименьший объект диапазона, включающий в себя заданные диапазоны. Например, GetBoundingRect для "B2:C5" и "D10:E15" возвращает значение "B2:E15".

#### Синтаксис
```js
rangeObject.getBoundingRect(anotherRange);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|anotherRange|Range или string|Объект диапазона либо адрес или имя диапазона.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D4:G6";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var range = range.getBoundingRect("G4:H8");
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // Prints Sheet1!D4:H8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getCell(row: number, column: number)
Получает объект диапазона, содержащий одну ячейку, по номеру строки и столбца. Ячейка может находиться вне родительского диапазона, если она расположена в таблице листа. Возвращаемая ячейка располагается относительно верхней левой ячейки диапазона.

#### Синтаксис
```js
rangeObject.getCell(row, column);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|row|number|Номер строки ячейки, которую требуется извлечь. Используется нулевой индекс.|
|column|number|Номер столбца ячейки, которую требуется извлечь. Используется нулевой индекс.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var cell = range.getCell(0,0);
    cell.load('address');
    return ctx.sync().then(function() {
        console.log(cell.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getColumn(column: number)
Возвращает столбец в диапазоне.

#### Синтаксис
```js
rangeObject.getColumn(column);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|column|number|Номер столбца диапазона, который требуется извлечь. Используется нулевой индекс.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet19";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getColumn(1);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!B1:B8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getEntireColumn()
Возвращает объект, представляющий весь столбец диапазона.

#### Синтаксис
```js
rangeObject.getEntireColumn();
```

#### Параметры
Нет

#### Возвращаемое значение
[Range](range.md)

#### Примеры

Примечание. Свойства сетки Range (values, numberFormat, formulas) содержат `null`, так как данный диапазон не ограничен.

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeEC = range.getEntireColumn();
    rangeEC.load('address');
    return ctx.sync().then(function() {
        console.log(rangeEC.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### getEntireRow()
Возвращает объект, представляющий всю строку диапазона.

#### Синтаксис
```js
rangeObject.getEntireRow();
```

#### Параметры
Нет

#### Возвращаемое значение
[Range](range.md)

#### Примеры
```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "D:F"; 
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeER = range.getEntireRow();
    rangeER.load('address');
    return ctx.sync().then(function() {
        console.log(rangeER.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
Свойства сетки Range (values, numberFormat, formulas) содержат `null`, так как данный диапазон не ограничен.

### getIntersection(anotherRange: Range или string)
Возвращает объект диапазона, представляющий собой прямоугольное пересечение заданных диапазонов.

#### Синтаксис
```js
rangeObject.getIntersection(anotherRange);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|anotherRange|Range или string|Объект диапазона или адрес диапазона, который будет использоваться для определения пересечения диапазонов.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getIntersection("D4:G6");
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!D4:F6
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getLastCell()
Возвращает последнюю ячейку в диапазоне. Например, последняя ячейка диапазона B2:D5 — D5.

#### Синтаксис
```js
rangeObject.getLastCell();
```

#### Параметры
Нет

#### Возвращаемое значение
[Range](range.md)

#### Примеры

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastCell();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getLastColumn()
Возвращает последний столбец в диапазоне. Например, последний столбец диапазона B2:D5 — D2:D5.

#### Синтаксис
```js
rangeObject.getLastColumn();
```

#### Параметры
Нет

#### Возвращаемое значение
[Range](range.md)

#### Примеры

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastColumn();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!F1:F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getLastRow()
Возвращает последнюю строку в диапазоне. Например, последняя строка в диапазоне "B2:D5" — "B5:D5".

#### Синтаксис
```js
rangeObject.getLastRow();
```

#### Параметры
Нет

#### Возвращаемое значение
[Range](range.md)

#### Примеры

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastRow();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!A8:F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```



### getOffsetRange(rowOffset: number, columnOffset: number)
Возвращает объект, представляющий диапазон, который смещен от указанного диапазона. Измерение возвращаемого диапазона будет соответствовать этому диапазону. Если полученный диапазон выходит за пределы таблицы листа, вызывается исключение.

#### Синтаксис
```js
rangeObject.getOffsetRange(rowOffset, columnOffset);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|rowOffset|number|Количество строк (положительное, отрицательное или нулевое), на которое необходимо сместить диапазон. Положительные значения соответствуют смещению вниз, а отрицательные — вверх.|
|columnOffset|number|Количество столбцов (положительное, отрицательное или 0), на который нужно сместить диапазон. Положительные значения соответствуют смещению вправо, а отрицательные — влево.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D4:F6";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getOffsetRange(-1,4);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!H3:K5
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getRow(row: number)
Возвращает строку в диапазоне.

#### Синтаксис
```js
rangeObject.getRow(row);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|row|number|Номер строки диапазона, который требуется извлечь. Используется нулевой индекс.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getRow(1);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!A2:F2
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getUsedRange(valuesOnly: bool)
Возвращает используемый диапазон заданного объекта диапазона.

#### Синтаксис
```js
rangeObject.getUsedRange(valuesOnly);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|valuesOnly|Bool|Необязательный параметр. Если задано значение true, используемыми считаются только ячейки, для которых установлены значения. Если задано значение по умолчанию false, используемыми считаются все ячейки, для которых когда-либо устанавливалось значение.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeUR = range.getUsedRange();
    rangeUR.load('address');
    return ctx.sync().then(function() {
        console.log(rangeUR.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### insert(shift: string)
Вставляет ячейку или диапазон ячеек на лист вместо этого диапазона, а также сдвигает другие ячейки, чтобы освободить место. Возвращает новый объект Range в пустом месте.

#### Синтаксис
```js
rangeObject.insert(shift);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|shift|string|Определяет способ сдвига ячеек. Возможные значения: Down, Right.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры

```js
    
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F5:F10";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.insert();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### load(param: object)
Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.

#### Синтаксис
```js
object.load(param);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|param|object|Необязательный параметр. Принимает имена параметров и связей в виде строки с разделителями или массива. Либо укажите объект [loadOption](loadoption.md).|

#### Возвращаемое значение
void

### merge(across: bool)
Объединяет ячейки диапазона в одну область на листе.

#### Синтаксис
```js
rangeObject.merge(across);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|across|Bool|Необязательный параметр. Установите значение true, чтобы объединить ячейки в каждой строке заданного диапазона как отдельные объединенные ячейки. Значение по умолчанию — false.|

#### Возвращаемое значение
void

#### Примеры
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.merge(true);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### select()
Выбирает указанный диапазон в пользовательском интерфейсе Excel.

#### Синтаксис
```js
rangeObject.select();
```

#### Параметры
Нет

#### Возвращаемое значение
void

#### Примеры

```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "F5:F10"; 
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.select();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### unmerge()
Разъединяет диапазон объединенных ячеек в отдельные ячейки.

#### Синтаксис
```js
rangeObject.unmerge();
```

#### Параметры
Нет

#### Возвращаемое значение
void

#### Примеры
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.unmerge();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### Примеры доступа к свойствам

Этот пример использует адрес диапазона, чтобы получить соответствующий объект.

```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8"; 
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.load('cellCount');
    return ctx.sync().then(function() {
        console.log(range.cellCount);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Этот пример использует именованный диапазон, чтобы получить объект диапазона.

```js

Excel.run(function (ctx) { 
    var rangeName = 'MyRange';
    var range = ctx.workbook.names.getItem(rangeName).range;
    range.load('cellCount');
    return ctx.sync().then(function() {
        console.log(range.cellCount);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Следующий пример задает numberFormat, значения и формулы для таблицы, которая содержит таблицу 2 x 3.

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F5:G7";
    var numberFormat = [[null, "d-mmm"], [null, "d-mmm"], [null, null]]
    var values = [["Today", 42147], ["Tomorrow", "5/24"], ["Difference in days", null]];
    var formulas = [[null,null], [null,null], [null,"=G6-G5"]];
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.numberFormat = numberFormat;
    range.values = values;
    range.formulas= formulas;
    range.load('text');
    return ctx.sync().then(function() {
        console.log(range.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
Он не отличается от приведенного выше примера за исключением того, что для формул используется формат R1C1.

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F5:G7";
    var numberFormat = [[null, "d-mmm"], [null, "d-mmm"], [null, null]]
    var values = [["Today", 42147], ["Tomorrow", "5/24"], ["Difference in days", null]];
    var formulasR1C1 = [[null,null], [null,null], [null,"=R[-1]C-R[-2]C"]];
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.numberFormat = numberFormat;
    range.values = values;
    range.formulasR1C1= formulasR1C1;
    range.load('text');
    return ctx.sync().then(function() {
        console.log(range.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
Возвращает лист, содержащий диапазон. 

```js
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var namedItem = names.getItem('MyRange');
    range = namedItem.range;
    var rangeWorksheet = range.worksheet;
    rangeWorksheet.load('name');
    return ctx.sync().then(function() {
            console.log(rangeWorksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

