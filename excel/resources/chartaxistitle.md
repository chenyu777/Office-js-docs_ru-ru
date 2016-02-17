# Объект ChartAxisTitle (API JavaScript для Excel)

_Область применения: Excel 2016, Excel Online, Office 2016_

Представляет название оси диаграммы.

## Свойства

| Свойство   | Тип|Описание
|:---------------|:--------|:----------|
|text|string|Обозначает название оси.|
|visible|Bool|Логическое значение, которое определяет видимость названия оси.|

_Ознакомьтесь с [примерами](#property-access-examples) доступа к свойствам._

## Связи
| Связь | Тип|Описание|
|:---------------|:--------|:----------|
|format|[ChartAxisTitleFormat](chartaxistitleformat.md)|Представляет форматирование для названия оси диаграммы. Только для чтения.|

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
Получение значения свойства `text` для названия оси диаграммы из оси значений диаграммы Chart1.

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var title = chart.axes.valueaxis.title;
	title.load('text');
	return ctx.sync().then(function() {
			console.log(title.text);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Добавление значения "Values" в качестве названия для оси значений.

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.axes.valueaxis.title.text = "Values";
	return ctx.sync().then(function() {
			console.log("Axis Title Added ");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

