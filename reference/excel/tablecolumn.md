# Объект TableColumn (API JavaScript для Excel)

Представляет столбец в таблице.

## Properties

| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|id|int|Возвращает уникальный ключ, идентифицирующий столбец в таблице. Только для чтения.|
|index|int|Возвращает номер индекса столбца в коллекции столбцов таблицы. Используется нулевой индекс. Только для чтения.|
|name|string|Возвращает имя столбца таблицы. Только для чтения.|
|values|object[][]|Представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейка, которая содержит ошибку, возвращает строку ошибки.|

_Ознакомьтесь с [примерами](#примерами) доступа к свойствам._

## Связи
| Связь | Тип   |Описание|
|:---------------|:--------|:----------|
|фильтр|[Фильтр](filter.md)|Возвращает фильтр, применяемый к столбцу. Только для чтения.|

## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Удаляет столбец из таблицы.|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Получает объект диапазона, связанный с текстом данных столбца.|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|Получает объект диапазона, связанный со строкой заголовков столбца.|
|[getRange()](#getrange)|[Range](range.md)|Получает объект диапазона, связанный со всем столбцом.|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|Получает объект диапазона, связанный со строкой итогов столбца.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе


### delete()
Удаляет столбец из таблицы.

#### Синтаксис
```js
tableColumnObject.delete();
```

#### Параметры
Нет

#### Возвращаемое значение
void

#### Примеры

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(2);
    column.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getDataBodyRange()
Получает объект диапазона, связанный с текстом данных столбца.

#### Синтаксис
```js
tableColumnObject.getDataBodyRange();
```

#### Параметры
Нет

#### Возвращаемое значение
[Range](range.md)

#### Примеры

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
    var dataBodyRange = column.getDataBodyRange();
    dataBodyRange.load('address');
    return ctx.sync().then(function() {
        console.log(dataBodyRange.address);
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getHeaderRowRange()
Получает объект диапазона, связанный со строкой заголовков столбца.

#### Синтаксис
```js
tableColumnObject.getHeaderRowRange();
```

#### Параметры
Нет

#### Возвращаемое значение
[Range](range.md)

#### Примеры

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
    var headerRowRange = columns.getHeaderRowRange();
    headerRowRange.load('address');
    return ctx.sync().then(function() {
        console.log(headerRowRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### getRange()
Получает объект диапазона, связанный со всем столбцом.

#### Синтаксис
```js
tableColumnObject.getRange();
```

#### Параметры
Нет

#### Возвращаемое значение
[Range](range.md)

#### Примеры

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
    var columnRange = columns.getRange();
    columnRange.load('address');
    return ctx.sync().then(function() {
        console.log(columnRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getTotalRowRange()
Получает объект диапазона, связанный со строкой итогов столбца.

#### Синтаксис
```js
tableColumnObject.getTotalRowRange();
```

#### Параметры
Нет

#### Возвращаемое значение
[Range](range.md)

#### Примеры

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
    var totalRowRange = columns.getTotalRowRange();
    totalRowRange.load('address');
    return ctx.sync().then(function() {
        console.log(totalRowRange.address);
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
### Примеры доступа к свойствам

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItem(0);
    column.load('index');
    return ctx.sync().then(function() {
        console.log(column.index);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    var newValues = [["New"], ["Values"], ["For"], ["New"], ["Column"]];
    var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(2);
    column.values = newValues;
    column.load('values');
    return ctx.sync().then(function() {
        console.log(column.values);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
