# Объект RangeFormat (API JavaScript для Excel)

_Область применения: Excel 2016, Excel Online, Office 2016_

Объект формата, в который включены шрифт, заливка, границы, выравнивание и другие свойства диапазона.

## Свойства

| Свойство   | Тип|Описание
|:---------------|:--------|:----------|
|horizontalAlignment|string|Представляет выравнивание по горизонтали для указанного объекта. Возможные значения: General, Left, Center, Right, Fill, Justify, CenterAcrossSelection, Distributed.|
|verticalAlignment|string|Представляет выравнивание по вертикали для указанного объекта. Возможные значения: Top, Center, Bottom, Justify, Distributed.|
|wrapText|bool|Указывает, что текстовый элемент управления Excel переносит текст в объекте. Значение null указывает, что для диапазона в целом не используется согласованный параметр переноса текста.|

_Ознакомьтесь с [примерами](#property-access-examples) доступа к свойствам._

## Связи
| Связь | Тип|Описание|
|:---------------|:--------|:----------|
|borders|[RangeBorderCollection](rangebordercollection.md)|Коллекция объектов границы, которые применяются к общему выделенному диапазону. Только для чтения.|
|fill|[RangeFill](rangefill.md)|Возвращает объект заливки, определенный для всего диапазона. Только для чтения.|
|font|[RangeFont](rangefont.md)|Возвращает объект шрифта, определенный для общего выбранного диапазона. Только для чтения.|

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

Этот пример распечатывает все свойства форматирования диапазона. 

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F:G";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.load(["format/*", "format/fill", "format/borders", "format/font"]);
	return ctx.sync().then(function() {
		console.log(range.format.wrapText);
		console.log(range.format.fill.color);
		console.log(range.format.font.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Пример ниже задает имя шрифта и цвет заливки диапазона, а также применяет перенос текста. 

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F:G";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.format.wrapText = true;
	range.format.font.name = 'Times New Roman';
	range.format.fill.color = '0000FF';
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Следующий пример добавляет границу сетки вокруг диапазона.

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F:G";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.format.borders('InsideHorizontal').lineStyle = 'Continuous';
	range.format.borders('InsideVertical').lineStyle = 'Continuous';
	range.format.borders('EdgeBottom').lineStyle = 'Continuous';
	range.format.borders('EdgeLeft').lineStyle = 'Continuous';
	range.format.borders('EdgeRight').lineStyle = 'Continuous';
	range.format.borders('EdgeTop').lineStyle = 'Continuous';
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
