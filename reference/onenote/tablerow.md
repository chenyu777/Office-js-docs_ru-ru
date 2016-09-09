# Объект TableRow (API JavaScript для OneNote)

_Относится к: OneNote Online_  


Представляет строку в таблице.

## Properties

| Свойство     | Тип   |Описание|Отзыв|
|:---------------|:--------|:----------|:-------|
|cellCount|int|Получает количество ячеек в строке. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-cellCount)|
|id|строка|Получает идентификатор строки. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-id)|
|rowIndex|int|Получает индекс строки в родительской таблице. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-rowIndex)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
| Связь | Тип   |Описание| Отзывы|
|:---------------|:--------|:----------|:-------|
|cells|[TableCellCollection](tablecellcollection.md)|Получает ячейки в строке. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-cells)|
|parentTable|[Table](table.md)|Получает родительскую таблицу. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-parentTable)|

## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[clear()](#clear)|void|Очищает содержимое строки.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-clear)|
|[insertRowAsSibling(insertLocation: строка, values: строка[])](#insertrowassiblinginsertlocation-string-values-string)|[TableRow](tablerow.md)|Вставляет строку перед текущей строкой или после нее.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-insertRowAsSibling)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-load)|
|[setShadingColor(colorCode: string)](#setshadingcolorcolorcode-string)|void|Задает цвет заливки всех ячеек в строке.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-setShadingColor)|

## Сведения о методе


### clear()
Очищает содержимое строки.

#### Синтаксис
```js
tableRowObject.clear();
```

#### Параметры
Нет

#### Возвращаемое значение
void

### insertRowAsSibling(insertLocation: string, values: string[])
Вставляет строку перед текущей строкой или после нее.

#### Синтаксис
```js
tableRowObject.insertRowAsSibling(insertLocation, values);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|insertLocation|строка|Определяет, куда необходимо вставить новые строки относительно текущей строки.  Возможные значения: Before, After|
|values|строка[]|Необязательный. Строки, которые необходимо вставить в новую строку, заданные в виде массива. Ячеек в этом параметре не должно быть больше, чем в текущей строке. Необязательный.|

#### Возвращаемое значение
[TableRow](tablerow.md)

#### Примеры
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get table rows.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                
                // Queue a command to load table.rows.
                ctx.load(table, "rows");
                
                // Run the queued commands
                return ctx.sync().then(function() {
                    var rows = table.rows;
                    rows.items[1].insertRowAsSibling("Before", ["cell0", "cell1"]);
                    return ctx.sync();
                });
            }
        }
    })
})
.catch(function(error) {
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

### setShadingColor(colorCode: string)
Задает цвет заливки всех ячеек в строке.

#### Синтаксис
```js
tableRowObject.setShadingColor(colorCode);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|colorCode|string|Код цвета, который нужно задать ячейкам.|

#### Возвращаемое значение
void
### Примеры доступа к свойствам
**id, cellCount, rowIndex**
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get table rows.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                
                // Queue a command to load table.rows.
                ctx.load(table, "rows");
                return ctx.sync().then(function() {
                    var rows = table.rows;
                    
                    // for each table row, log cell count and row index.
                    for (var i = 0; i < rows.items.length; i++) {
                        console.log("Row " + i + " Id: " + rows.items[i].id);
                        console.log("Row " + i + " Cell Count: " + rows.items[i].cellCount);
                        console.log("Row " + i + " Row Index: " + rows.items[i].rowIndex);
                    }
                    return ctx.sync();
                });
            }
        }
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**parentTable, cells**
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get table rows.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                
                // Queue a command to load parentTable and cells of each row in the table.
                ctx.load(table, "rows/parentTable, rows/cells");
                return ctx.sync().then(function() {
                    var rows = table.rows;
                    
                    // for each row, log parentTable and cells
                    for (var i = 0; i < rows.items.length; i++) {
                        console.log("Row " + i + " Parent Table Id: " + rows.items[i].parentTable.id);
                        var cells = rows.items[i].cells;
                        for (var j = 0 ; j < cells.items.length; j++) {
                            console.log("Row " + i + " Cell " + j + " Id: " + cells.items[j].id);
                        }
                    }
                    return ctx.sync();
                });
            }
        }
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

