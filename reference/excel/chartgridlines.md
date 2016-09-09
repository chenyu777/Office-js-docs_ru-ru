# Объект ChartGridlines (API JavaScript для Excel)

Представляет основные или вспомогательные линии сетки на оси диаграммы.

## Properties

| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|visible|bool|Логическое значение, определяющее, отображаются ли линии сетки оси.|

_Ознакомьтесь с [примерами](#примерами) доступа к свойствам._

## Связи
| Связь | Тип   |Описание|
|:---------------|:--------|:----------|
|format|[ChartGridlinesFormat](chartgridlinesformat.md)|Представляет форматирование линий сетки диаграммы. Только для чтения.|

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

Получение свойства `visible` основных линий сетки на оси значений диаграммы Chart1.

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    var majGridlines = chart.axes.valueaxis.majorGridlines;
    majGridlines.load('visible');
    return ctx.sync().then(function() {
            console.log(majGridlines.visible);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Отображение основных линий сетки на оси значений диаграммы Chart1.

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.axes.valueaxis.majorgridlines.visible = true;
    return ctx.sync().then(function() {
            console.log("Axis Gridlines Added ");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
