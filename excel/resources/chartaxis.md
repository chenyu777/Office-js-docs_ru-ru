# Объект ChartAxis (API JavaScript для Excel)

_Область применения: Excel 2016, Excel Online, Office 2016_

Представляет одну ось на диаграмме.

## Свойства

| Свойство   | Тип|Описание
|:---------------|:--------|:----------|
|majorUnit|object|Обозначает интервал между двумя основными делениями. Его можно указать в виде числового значения или пустой строки. Возвращаемое значение всегда является числом.|
|maximum|object|Представляет максимальное значение на оси значений. Его можно указать в виде числового значения или пустой строки (для автоматически заданных значений оси). Возвращаемое значение всегда является числом.|
|minimum|object|Представляет минимальное значение на оси значений. Его можно указать в виде числового значения или пустой строки (для автоматически заданных значений оси). Возвращаемое значение всегда является числом.|
|minorUnit|object|Представляет интервал между двумя промежуточными делениями. Его можно указать в виде числового значения или пустой строки (для автоматически заданных значений оси). Возвращаемое значение всегда является числом.|

_Ознакомьтесь с [примерами](#property-access-examples) доступа к свойствам._

## Связи
| Связь | Тип|Описание|
|:---------------|:--------|:----------|
|format|[ChartAxisFormat](chartaxisformat.md)|Представляет форматирование объекта диаграммы, в том числе форматирование линий и шрифта. Только для чтения.|
|majorGridlines|[ChartGridlines](chartgridlines.md)|Возвращает объект линии сетки, который представляет основные линии сетки для указанной оси. Только для чтения.|
|minorGridlines|[ChartGridlines](chartgridlines.md)|Возвращает объект Gridlines, который представляет вспомогательные линии сетки для указанной оси. Только для чтения.|
|title|[ChartAxisTitle](chartaxistitle.md)|Представляет название оси. Только для чтения.|

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
Получение значения свойства `maximum` для оси диаграммы из Chart1

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var axis = chart.axes.valueaxis;
	axis.load('maximum');
	return ctx.sync().then(function() {
			console.log(axis.maximum);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Настройка свойства `maximum`, `minimum`, `majorunit` или `minorunit` оси значений. 

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.axes.valueaxis.maximum = 5;
	chart.axes.valueaxis.minimum = 0;
	chart.axes.valueaxis.majorunit = 1;
	chart.axes.valueaxis.minorunit = 0.2;
	return ctx.sync().then(function() {
			console.log("Axis Settings Changed");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

