# Объект Chart (API JavaScript для Excel)

Представляет объект диаграммы в книге.

## Свойства

| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|height|double|Обозначает высоту объекта диаграммы (в пунктах).|
|id|string|Возвращает диаграмму с учетом ее положения в коллекции. Только для чтения.|
|left|double|Расстояние в пунктах от левого края диаграммы до начала листа.|
|name|string|Обозначает имя объекта диаграммы.|
|top|double|Представляет расстояние в пунктах от верхнего края объекта до верхнего края первой строки (на листе) или до верхнего края области диаграммы (на диаграмме).|
|width|double|Представляет ширину объекта диаграммы (в пунктах).|

_Ознакомьтесь с [примерами](#примерами) доступа к свойствам._

## Связи
| Связь | Тип   |Описание|
|:---------------|:--------|:----------|
|axes|[ChartAxes](chartaxes.md)|Представляет оси диаграммы. Только для чтения.|
|dataLabels|[ChartDataLabels](chartdatalabels.md)|Представляет метки данных на диаграмме. Только для чтения.|
|format|[ChartAreaFormat](chartareaformat.md)|Инкапсулирует свойства формата для области диаграммы. Только для чтения.|
|legend|[ChartLegend](chartlegend.md)|Представляет условные обозначения для диаграммы. Только для чтения.|
|series|[ChartSeriesCollection](chartseriescollection.md)|Представляет один ряд данных или коллекцию рядов данных в диаграмме. Только для чтения.|
|должности.|[ChartTitle](charttitle.md)|Представляет заголовок указанной диаграммы, включая его текст, видимость, положение и форматирование. Только для чтения.|

## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Удаляет объект диаграммы.|
|[getImage(height: number, width: number, fittingMode: string)](#getimageheight-number-width-number-fittingmode-string)|System.IO.Stream|Отрисовывает диаграмму в виде изображения с кодировкой base64, масштабируя ее в соответствии с указанным размером.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|
|[setData(sourceData: Range, seriesBy: string)](#setdatasourcedata-range-seriesby-string)|void|Сбрасывает исходные данные для диаграммы.|
|[setPosition(startCell: Range или string, endCell: Range или string)](#setpositionstartcell-range-или-string-endcell-range-или-string)|void|Располагает диаграмму относительно ячеек на листе.|

## Сведения о методе


### delete()
Удаляет объект диаграммы.

#### Синтаксис
```js
chartObject.delete();
```

#### Параметры
Нет

#### Возвращаемое значение
void

#### Примеры
```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### getImage(height: number, width: number, fittingMode: string)
Отрисовывает диаграмму в виде изображения с кодировкой base64, масштабируя ее в соответствии с указанным размером.

#### Синтаксис
```js
chartObject.getImage(height, width, fittingMode);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|height|number|Необязательный. (Необязательно) Нужная высота создаваемого изображения.|
|width|number|Необязательный. (Необязательно) Нужная ширина создаваемого изображения.|
|fittingMode|строка|Необязательный. (Необязательно) Метод, используемый для масштабирования диаграммы до указанного размера (если указаны и высота, и ширина).  Возможные значения: Fit, FitAndCenter, Fill|

#### Возвращаемое значение
System.IO.Stream

#### Примеры
```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    var image = chart.getImage();
    return ctx.sync(); 
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

### setData(sourceData: Range, seriesBy: string)
Сбрасывает исходные данные для диаграммы.

#### Синтаксис
```js
chartObject.setData(sourceData, seriesBy);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|sourceData|Range|Объект Range, соответствующий исходным данным.|
|seriesBy|string|Необязательный параметр. Определяет способ использования столбцов или строк в качестве рядов данных на диаграмме. Возможные значения: Auto, Columns, Rows. Если выбрано значение auto, классическое приложение изучает исходные данные и автоматически определяет, расположены ли они по строкам или по столбцам. В Excel Online при этом по умолчанию выбирается значение columns.|

#### Возвращаемое значение
void

#### Примеры

Указание значения "A1:B4" для параметра `sourceData` и "Columns" — для `seriesBy`

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    var sourceData = "A1:B4";
    chart.setData(sourceData, "Columns");
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### setPosition(startCell: Range или string, endCell: Range или string)
Располагает диаграмму относительно ячеек на листе.

#### Синтаксис
```js
chartObject.setPosition(startCell, endCell);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|startCell|Range или string|Начальная ячейка. Место, куда будет перемещена диаграмма. Начальная ячейка — это верхняя левая или верхняя правая ячейка (это зависит от того, использует ли пользователь параметры отображения слева направо).|
|endCell|Range или string|Необязательный параметр. Конечная ячейка. Если указан данный параметр, значения ширины и высоты диаграммы устанавливаются так, чтобы полностью покрыть эту ячейку или диапазон.|

#### Возвращаемое значение
void

#### Примеры


```js
Excel.run(function (ctx) { 
    var sheetName = "Charts";
    var sourceData = sheetName + "!" + "A1:B4";
    var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("pie", sourceData, "auto");
    chart.width = 500;
    chart.height = 300;
    chart.setPosition("C2", null);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### Примеры доступа к свойствам

Получение диаграммы Chart1

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.load('name');
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

Обновление диаграммы, включая переименование, размещение и изменение размера.

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.name="New Name";
    chart.top = 100;
    chart.left = 100;
    chart.height = 200;
    chart.weight = 200;
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Переименование диаграммы, изменение размера до 200 пунктов по высоте и по ширине. Перемещение Chart1 на 100 пунктов вверх и влево. 

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");
    chart.name="New Name";  
    chart.top = 100;
    chart.left = 100;
    chart.height =200;
    chart.width =200;
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

