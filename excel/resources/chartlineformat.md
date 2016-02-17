# Объект ChartLineFormat (API JavaScript для Excel)

_Область применения: Excel 2016, Excel Online, Office 2016_

Инкапсулирует параметры форматирования для элементов линий.

## Свойства

| Свойство   | Тип|Описание
|:---------------|:--------|:----------|
|color|string|HTML-код цвета, представляющий цвет линий в диаграмме.|

_Ознакомьтесь с [примерами](#property-access-examples) доступа к свойствам._

## Связи
Нет


## Методы

| Метод   | Возвращаемый тип|Описание|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Очищает формат линий элемента диаграммы.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе

### clear()
Очищает формат линий элемента диаграммы.

#### Синтаксис
```js
chartLineFormatObject.clear();
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

Установка красного цвета для основных линий сетки на оси значений.

```js
Excel.run(function (ctx) { 
	var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.axes.valueaxis.majorGridlines;
	gridlines.format.line.color = "#FF0000";
	return ctx.sync().then(function() {
			console.log("Chart Gridlines Color Updated");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

