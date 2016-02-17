# Объект RangeBorderCollection (API JavaScript для Excel)

_Область применения: Excel 2016, Excel Online, Office 2016_

Представляет объекты границ, из которых состоит граница диапазона.

## Свойства

| Свойство   | Тип|Описание
|:---------------|:--------|:----------|
|count|int|Количество объектов границы в коллекции. Только для чтения.|
|items|[RangeBorder[]](rangeborder.md)|Коллекция объектов rangeBorder. Только для чтения.|

_Ознакомьтесь с [примерами](#property-access-examples) доступа к свойствам._

## Связи
Нет


## Методы

| Метод   | Возвращаемый тип|Описание|
|:---------------|:--------|:----------|
|[getItem(index: string)](#getitemindex-string)|[RangeBorder](rangeborder.md)|Возвращает объект границы по его имени.|
|[getItemAt(index: number)](#getitematindex-number)|[RangeBorder](rangeborder.md)|Возвращает объект границы по его индексу.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе

### getItem(index: string)
Возвращает объект границы по его имени. 

#### Синтаксис
```js
rangeBorderCollectionObject.getItem(index);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|index|string|Значение индекса получаемого объекта границы. Возможные значения: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight, InsideVertical, InsideHorizontal, DiagonalDown, DiagonalUp.|

#### Возвращаемое значение
[RangeBorder](rangeborder.md)

#### Примеры
```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	var borderName = 'EdgeTop';
	var border = range.format.borders.getItem(borderName);
	border.load('style');
	return ctx.sync().then(function() {
			console.log(border.style);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


#### Примеры
```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	var border = ctx.workbook.borders.getItemAt(0);
	border.load('sideIndex');
	return ctx.sync().then(function() {
			console.log(border.sideIndex);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getItemAt(index: number)
Возвращает объект границы по его индексу.

#### Синтаксис
```js
rangeBorderCollectionObject.getItemAt(index);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|index|number|Значение индекса получаемого объекта. Используется нулевой индекс.|

#### Возвращаемое значение
[RangeBorder](rangeborder.md)

#### Примеры
```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	var border = ctx.workbook.borders.getItemAt(0);
	border.load('sideIndex');
	return ctx.sync().then(function() {
			console.log(border.sideIndex);
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

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	var borders = range.format.borders;
	border.load('items');
	return ctx.sync().then(function() {
		console.log(borders.count);
		for (var i = 0; i < borders.items.length; i++)
		{
			console.log(borders.items[i].sideIndex);
		}
	});
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
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.format.borders.getItem('InsideHorizontal').style = 'Continuous';
	range.format.borders.getItem('InsideVertical').style = 'Continuous';
	range.format.borders.getItem('EdgeBottom').style = 'Continuous';
	range.format.borders.getItem('EdgeLeft').style = 'Continuous';
	range.format.borders.getItem('EdgeRight').style = 'Continuous';
	range.format.borders.getItem('EdgeTop').style = 'Continuous';
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
