# Объект ChartFill (API JavaScript для Excel)

Представляет форматирование заливки для элемента диаграммы.

## Свойства

Нет

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Очищает цвет заливки элемента диаграммы.|
|[setSolidColor(color: string)](#setsolidcolorcolor-string)|void|Устанавливает форматирование заливки элемента диаграммы на единый цвет.|

## Сведения о методе


### clear()
Очищает цвет заливки элемента диаграммы.

#### Синтаксис
```js
chartFillObject.clear();
```

#### Параметры
Нет

#### Возвращаемое значение
void

#### Примеры

Очищает формат основных линий сетки на оси значений диаграммы Chart1.

```js
Excel.run(function (ctx) { 
    var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueaxis.majorGridlines;   
    gridlines.format.line.clear();
    return ctx.sync().then(function() {
            console.log("Chart Major Gridlines Format Cleared");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### setSolidColor(color: string)
Устанавливает форматирование заливки элемента диаграммы на единый цвет.

#### Синтаксис
```js
chartFillObject.setSolidColor(color);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|color|string|HTML-код, представляющий цвет линии границы в формате #RRGGBB (например, "FFA500") или в виде ключевого слова в HTML (например, "orange").|

#### Возвращаемое значение
void

#### Примеры

Установка красного в качестве фонового цвета диаграммы Chart1.

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 

    chart.format.fill.setSolidColor("#FF0000");

    return ctx.sync().then(function() {
            console.log("Chart1 Background Color Changed.");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
