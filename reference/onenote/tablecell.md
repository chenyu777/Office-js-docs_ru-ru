# Объект TableCell (API JavaScript для OneNote)

_Относится к: OneNote Online_  


Представляет ячейку в таблице OneNote.

## Свойства

| Свойство     | Тип   |Описание|Отзыв|
|:---------------|:--------|:----------|:-------|
|cellIndex|int|Получает индекс ячейки в строке. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-cellIndex)|
|id|строка|Получает идентификатор ячейки. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-id)|
|rowIndex|int|Получает индекс строки ячейки в таблице. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-rowIndex)|
|shadingColor|string|Возвращает и задает цвет заливки ячейки|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-shadingColor)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
| Связь | Тип   |Описание| Отзывы|
|:---------------|:--------|:----------|:-------|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Получает коллекцию объектов Paragraph в объекте TableCell. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-paragraphs)|
|parentRow|[TableRow](tablerow.md)|Получает родительскую строку ячейки. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-parentRow)|

## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[appendHtml(html: string)](#appendhtmlhtml-string)|void|Добавляет указанную строку HTML в нижнюю часть объекта TableCell.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendHtml)|
|[appendImage(base64EncodedImage: string, width: double, height: double)](#appendimagebase64encodedimage-string-width-double-height-double)|[Image](image.md)|Добавляет указанное изображение в ячейку таблицы.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendImage)|
|[appendRichText(paragraphText: string)](#appendrichtextparagraphtext-string)|[RichText](richtext.md)|Добавляет указанный текст в ячейку таблицы.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendRichText)|
|[appendTable(rowCount: number, columnCount: number, values: string[][])](#appendtablerowcount-number-columncount-number-values-string)|[Table](table.md)|Добавляет таблицу с указанным количеством строк и столбцов в ячейку таблицы.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendTable)|
|[clear()](#clear)|void|Очищает содержимое ячейки.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-clear)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-load)|

## Сведения о методе


### appendHtml(html: string)
Добавляет указанную строку HTML в нижнюю часть объекта TableCell.

#### Синтаксис
```js
tableCellObject.appendHtml(html);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|html|строка|Строка HTML, которую необходимо добавить. Сведения об API JavaScript для надстроек OneNote см. в разделе [Поддерживаемые элементы HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html).|

#### Возвращаемое значение
void

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
        
        // for each table, get a table cell at row one and column two and add "Hello".
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
                cell.appendHtml("<p>Hello</p>");
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});


### appendImage(base64EncodedImage: string, width: double, height: double)
Adds the specified image to table cell.

#### Syntax
```js
tableCellObject.appendImage(base64EncodedImage, width, height);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|base64EncodedImage|строка|Строка HTML, которую необходимо добавить.|
|width|double|Необязательный. Ширина в точках. Значение по умолчанию — null, ширина изображения имеет приоритет.|
|height|double|Необязательный. Высота в точках. Значение по умолчанию — null, высота изображения имеет приоритет.|

#### Возвращаемое значение
[Image](image.md)

### appendRichText(paragraphText: string)
Добавляет указанный текст в ячейку таблицы.

#### Синтаксис
```js
tableCellObject.appendRichText(paragraphText);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|paragraphText|строка|Строка HTML, которую необходимо добавить.|

#### Возвращаемое значение
[RichText](richtext.md)

#### Примеры
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    var appendedRichText = null;
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get a table cell at row one and column two and add "Hello".
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
                appendedRichText = cell.appendRichText("Hello");
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### appendTable(rowCount: number, columnCount: number, values: string[][])
Добавляет таблицу с указанным количеством строк и столбцов в ячейку таблицы.

#### Синтаксис
```js
tableCellObject.appendTable(rowCount, columnCount, values);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|rowCount|number|Обязательный. Количество строк в таблице.|
|columnCount|number|Обязательный. Количество столбцов в таблице.|
|values|string[][]|Необязательный. Необязательный двухмерный массив. Ячейки заполняются, если в массиве указаны соответствующие строки.|

#### Возвращаемое значение
[Table](table.md)

### clear()
Очищает содержимое ячейки.

#### Синтаксис
```js
tableCellObject.clear();
```

#### Параметры
Нет

#### Возвращаемое значение
void

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
### Примеры доступа к свойствам
**id, cellIndex, rowIndex**
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get a table cell at row one and column two.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
                
                // Queue a command to load the table cell.
                ctx.load(cell);
                ctx.sync().then(function() {
                    console.log("Cell Id: " + cell.id);
                    console.log("Cell Index: " + cell.cellIndex);
                    console.log("Cell's Row Index: " + cell.rowIndex);
                });
            }
        }
        return ctx.sync();
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
ParentTable, ParentRow, Paragraphs
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get a table cell at row one and column two.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
                
                // Queue a command to load parentTable, parentRow and paragraphs of the table cell.
                ctx.load(cell, "parentTable, parentRow, paragraphs");
                
                ctx.sync().then(function() {
                    console.log("Parent Table Id: " + cell.parentTable.id);
                    console.log("Parent Row Id: " + cell.parentRow.id);
                    var paragraphs = cell.paragraphs;
                    
                    for (var i = 0; i < paragraphs.items.length; i++) {
                        console.log("Paragraph Id: " + paragraphs.items[i].id);
                    }
                });
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

