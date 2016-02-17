# Объект ChartTitle (API JavaScript для Excel)

_Область применения: Excel 2016, Excel Online, Office 2016_

Представляет объект заголовка диаграммы.

## Свойства

| Свойство   | Тип|Описание
|:---------------|:--------|:----------|
|overlay|bool|Логическое значение, определяющее, отображается ли заголовок диаграммы поверх нее.|
|text|string|Представляет текст заголовка диаграммы.|
|visible|bool|Логическое значение, представляющее видимость объекта заголовка диаграммы.|

_Ознакомьтесь с [примерами](#property-access-examples) доступа к свойствам._

## Связи
| Связь | Тип|Описание|
|:---------------|:--------|:----------|
|format|[ChartTitleFormat](charttitleformat.md)|Представляет форматирование названия диаграммы, включая формат заливки и шрифта. Только для чтения.|

## Методы

| Метод   | Возвращаемый тип|Описание|
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
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|param|object|Необязательный параметр. Принимает имена параметров и связей в виде строки с разделителями или массива. Либо укажите объект [loadOption](loadoption.md).|

#### Возвращаемое значение
void
### Примеры доступа к свойствам

Получение свойства `text` для заголовка диаграммы Chart1.

```js
Excel.run(function (ctx) { 
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

var title = chart.title;
title.load('text');
return ctx.sync().then(function() {
		console.log(title.text);
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Свойству `text` заголовка диаграммы присваивается значение "My Chart". Заголовок отображается над диаграммой, но не поверх нее.

```js
Excel.run(function (ctx) { 
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.title.text= "My Chart"; 
chart.title.visible=true;
chart.title.overlay=true;

return ctx.sync().then(function() {
		console.log("Char Title Changed");
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

