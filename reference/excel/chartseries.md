# Объект ChartSeries (API JavaScript для Excel)

Представляет ряд в диаграмме.

## Properties

| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|name|string|Представляет имя ряда в диаграмме.|

_Ознакомьтесь с [примерами](#примерами) доступа к свойствам._

## Связи
| Связь | Тип   |Описание|
|:---------------|:--------|:----------|
|format|[ChartSeriesFormat](chartseriesformat.md)|Представляет форматирование ряда диаграммы, включая формат заливки и линий. Только для чтения.|
|points|[ChartPointsCollection](chartpointscollection.md)|Представляет коллекцию всех точек в ряду. Только для чтения.|

## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе


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

Изменение имени первого ряда диаграммы Chart1 на "New Series Name".

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.series.getItemAt(0).name = "New Series Name";
    return ctx.sync().then(function() {
            console.log("Series1 Renamed");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
