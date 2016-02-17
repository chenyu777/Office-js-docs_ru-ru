# Объект RangeFont (API JavaScript для Excel)

_Область применения: Excel 2016, Excel Online, Office 2016_

Этот объект представляет атрибуты шрифта (имя, размер, цвет и т. д.) для объекта.

## Свойства

| Свойство   | Тип|Описание
|:---------------|:--------|:----------|
|bold|bool|Представляет параметр полужирного шрифта.|
|color|string|HTML-код цвета текста. Например, #FF0000 обозначает красный.|
|italic|bool|Представляет параметр курсива.|
|name|string|Имя шрифта (например, Calibri)|
|size|double|font-size|
|underline|string|Тип подчеркивания, применяемый для шрифта. Возможные значения: None, Single, Double, SingleAccountant, DoubleAccountant.|

_Ознакомьтесь с [примерами](#property-access-examples) доступа к свойствам._

## Связи
Нет


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

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F:G";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	var rangeFont = range.format.font;
	rangeFont.load('name');
	return ctx.sync().then(function() {
		console.log(rangeFont.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
Следующий пример задает имя шрифта. 

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F:G";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.format.font.name = 'Times New Roman';
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
