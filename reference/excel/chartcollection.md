# Объект ChartCollection (API JavaScript для Excel)

Коллекция всех объектов диаграмм на листе.

## Свойства

| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|count|int|Возвращает количество диаграмм на листе. Только для чтения.|
|items|[Chart[]](chart.md)|Коллекция объектов диаграммы. Только для чтения.|

_Ознакомьтесь с [примерами](#примерами) доступа к свойствам._

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[add(type: string, sourceData: Range, seriesBy: string)](#addtype-string-sourcedata-range-seriesby-string)|[Chart](chart.md)|Создает диаграмму.|
|[getItem(name: string)](#getitemname-string)|[Chart](chart.md)|Возвращает диаграмму по ее имени. Если существует несколько диаграмм с таким именем, возвращается первая из них.|
|[getItemAt(index: number)](#getitematindex-number)|[Chart](chart.md)|Возвращает диаграмму с учетом ее положения в коллекции.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе


### add(type: string, sourceData: Range, seriesBy: string)
Создает диаграмму.

#### Синтаксис
```js
chartCollectionObject.add(type, sourceData, seriesBy);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|type|string|Представляет тип диаграммы. Возможные значения: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie и т. д.|
|sourceData|Range|Объект диапазона, содержащий исходные данные.|
|seriesBy|string|Необязательный параметр. Определяет способ использования столбцов или строк в качестве рядов данных на диаграмме. Возможные значения: Auto, Columns, Rows.|

#### Возвращаемое значение
[Chart](chart.md)

#### Примеры

Добавление диаграммы со значением ColumnClustered класса `chartType` на лист Charts, где в качестве параметра `sourceData` задан диапазон "A1:B4", а в качестве параметра `seriesBy` — значение "auto".

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var sourceData = sheetName + "!" + "A1:B4";
    var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("ColumnClustered", sourceData, "auto");
    return ctx.sync().then(function() {
            console.log("New Chart Added");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getItem(name: string)
Возвращает диаграмму по ее имени. Если существует несколько диаграмм с таким именем, возвращается первая из них.

#### Синтаксис
```js
chartCollectionObject.getItem(name);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|name|string|Имя получаемой диаграммы.|

#### Возвращаемое значение
[Chart](chart.md)

#### Примеры

```js
Excel.run(function (ctx) { 
    var chartname = 'Chart1';
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartname);
    return ctx.sync().then(function() {
            console.log(chart.height);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


#### Примеры

```js
Excel.run(function (ctx) { 
    var chartId = 'SamplChartId';
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartId);
    return ctx.sync().then(function() {
            console.log(chart.height);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```



#### Примеры

```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
    return ctx.sync().then(function() {
            console.log(chart.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getItemAt(index: number)
Возвращает диаграмму с учетом ее положения в коллекции.

#### Синтаксис
```js
chartCollectionObject.getItemAt(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number|Значение индекса получаемого объекта. Используется нулевой индекс.|

#### Возвращаемое значение
[Chart](chart.md)

#### Примеры

```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
    return ctx.sync().then(function() {
            console.log(chart.name);
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
    var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
    charts.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < charts.items.length; i++)
        {
            console.log(charts.items[i].name);
            console.log(charts.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Получение количества диаграмм.

```js
Excel.run(function (ctx) { 
    var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
    charts.load('count');
    return ctx.sync().then(function() {
        console.log("charts: Count= " + charts.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

