# Объект Table (API JavaScript для OneNote)

_Относится к: OneNote Online_  


Представляет таблицу на странице OneNote.

## Свойства

| Свойство     | Тип   |Описание|Отзыв|
|:---------------|:--------|:----------|:-------|
|borderVisible|bool|Задает отображение границ или возвращает сведения об отображении границ. Значение true, если они отображаются, значение false — если нет.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-borderVisible)|
|columnCount|int|Получает количество столбцов в таблице. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-columnCount)|
|id|string|Получает идентификатор таблицы. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-id)|
|rowCount|int|Получает количество строк в таблице. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-rowCount)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
| Связь | Тип   |Описание| Отзывы|
|:---------------|:--------|:----------|:-------|
|paragraph|[Paragraph](paragraph.md)|Получает объект Paragraph, содержащий объект Table. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-paragraph)|
|rows|[TableRowCollection](tablerowcollection.md)|Получает все строки таблицы. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-rows)|

## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[appendColumn(values: string[])](#appendcolumnvalues-string)|void|Добавляет столбец в конец таблицы. Значения указываются в новом столбце. В противном случае столбец будет пустым.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-appendColumn)|
|[appendRow(values: string[])](#appendrowvalues-string)|[TableRow](tablerow.md)|Добавляет строку в конец таблицы. Значения указываются в новой строке. В противном случае строка будет пустой.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-appendRow)|
|[clear()](#clear)|void|Очищает содержимое таблицы.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-clear)|
|[getCell(rowIndex: число, cellIndex: число)](#getcellrowindex-число-cellindex-число)|[TableCell](tablecell.md)|Получает ячейку таблицы в указанной строке и указанном столбце.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-getCell)|
|[insertColumn(index: number, values: string[])](#insertcolumnindex-number-values-string)|void|Вставляет столбец в положении заданного индекса в таблице. Значения указываются в новом столбце. В противном случае столбец будет пустым.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-insertColumn)|
|[insertRow(index: number, values: string[])](#insertrowindex-number-values-string)|[TableRow](tablerow.md)|Вставляет строку в положение заданного индекса в таблице. Значения указываются в новой строке. В противном случае строка будет пустой.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-insertRow)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-load)|
|[setShadingColor(colorCode: string)](#setshadingcolorcolorcode-string)|void|Задает цвет заливки всех ячеек в таблице.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-setShadingColor)|

## Сведения о методах


### appendColumn(values: string[])
Добавляет столбец в конец таблицы. Значения указываются в новом столбце. В противном случае столбец будет пустым.

#### Синтаксис
```js
tableObject.appendColumn(values);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|values|string[]|Необязательный. Необязательный. Строки, которые необходимо вставить в новый столбец, заданные в виде массива. Значений в этом параметре не должно быть больше, чем строк в таблице.|

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
        
        // for each table, append a column.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                table.appendColumn(["cell0", "cell1"]);
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


### appendRow(values: string[])
Добавляет строку в конец таблицы. Значения указываются в новой строке. В противном случае строка будет пустой.

#### Синтаксис
```js
tableObject.appendRow(values);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|values|строка[]|Необязательный. Необязательный. Строки, которые необходимо вставить в новую строку, заданные в виде массива. Значений в этом параметре не должно быть больше, чем столбцов в таблице.|

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
        
        // for each table, append a column.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var row = table.appendRow(["cell0", "cell1"]);
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


### clear()
Очищает содержимое таблицы.

#### Синтаксис
```js
tableObject.clear();
```

#### Параметры
Нет

#### Возвращаемое значение
void

### getCell(rowIndex: число, cellIndex: число)
Получает ячейку таблицы в указанной строке и указанном столбце.

#### Синтаксис
```js
tableObject.getCell(rowIndex, cellIndex);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|rowIndex|число|Индекс строки.|
|cellIndex|число|Индекс ячейки в строке.|

#### Возвращаемое значение
[TableCell](tablecell.md)

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
        
        // for each table, get a cell in the second row and third column.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(2 /*Row Index*/, 3 /*Column Index*/);
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


### insertColumn(index: число, values: строка[])
Вставляет столбец в положении заданного индекса в таблице. Значения указываются в новом столбце. В противном случае столбец будет пустым.

#### Синтаксис
```js
tableObject.insertColumn(index, values);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|число|Индекс в таблице, в положении которого будет вставлен столбец.|
|values|строка[]|Необязательный. Необязательный. Строки, которые необходимо вставить в новый столбец, заданные в виде массива. Значений в этом параметре не должно быть больше, чем строк в таблице.|

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
        
        // for each table, insert a column at index two.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                table.insertColumn(2, ["cell0", "cell1"]);
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


### insertRow(index: число, values: строка[])
Вставляет строку в положение заданного индекса в таблице. Значения указываются в новой строке. В противном случае строка будет пустой.

#### Синтаксис
```js
tableObject.insertRow(index, values);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|число|Индекс в таблице, в положении которого будет вставлена строка.|
|values|строка[]|Необязательный. Необязательный. Строки, которые необходимо вставить в новую строку, заданные в виде массива. Значений в этом параметре не должно быть больше, чем столбцов в таблице.|

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
        
        // for each table, insert a row at index two.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var row = table.insertRow(2, ["cell0", "cell1"]);
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
Задает цвет заливки всех ячеек в таблице.

#### Синтаксис
```js
tableObject.setShadingColor(colorCode);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|colorCode|string|Код цвета, который нужно задать ячейкам.|

#### Возвращаемое значение
void
### Примеры доступа к свойствам
**columnCount, rowCount, id**
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // For each table, log properties.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                ctx.load(table);
                return ctx.sync().then(function() {
                    console.log("Table Id: " + table.id);
                    console.log("Row Count: " + table.rowCount);
                    console.log("Column Count: " + table.columnCount);
                    return ctx.sync();
                });
            }
        }
    });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**paragraph, rows**
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, log its paragraph id.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                ctx.load(table, "paragraph/id, rows/id");
                return ctx.sync().then(function() {
                    console.log("Paragraph Id: " + table.paragraph.id);
                    var rows = table.rows;
                    
                    // for each rows in the table, log row index and id.
                    for (var i = 0; i < rows.items.length; i++) {
                        console.log("Row " + i + " Id: " + rows.items[i].id);
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

