# Объект Worksheet (API JavaScript для Excel)

_Область применения: Excel 2016, Excel Online, Office 2016_

Лист Excel представляет собой сетку ячеек. Он может содержать данные, таблицы, диаграммы и т. д.

## Свойства

| Свойство   | Тип|Описание
|:---------------|:--------|:----------|
|id|string|Возвращает значение, однозначно идентифицирующее лист в данной книге. Значение идентификатора остается прежним, даже если переименовать или переместить лист. Только для чтения.|
|name|string|Отображаемое имя листа.|
|position|int|Положение листа (начиная с нуля) в книге.|
|visibility|string|Видимость листа. Возможные значения: Visible, Hidden, VeryHidden. Только для чтения.|

_Ознакомьтесь с [примерами](#property-access-examples) доступа к свойствам._

## Связи
| Связь | Тип|Описание|
|:---------------|:--------|:----------|
|charts|[ChartCollection](chartcollection.md)|Возвращает коллекцию диаграмм, имеющихся на листе. Только для чтения.|
|tables|[TableCollection](tablecollection.md)|Коллекция таблиц, имеющихся на листе. Только для чтения.|

## Методы

| Метод   | Возвращаемый тип|Описание|
|:---------------|:--------|:----------|
|[activate()](#activate)|void|Активация листа в пользовательском интерфейсе Excel.|
|[delete()](#delete)|void|Удаляет лист из книги.|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|Получает объект диапазона, содержащий одну ячейку, по номеру строки и столбца. Ячейка может находиться вне родительского диапазона, если она расположена в таблице листа.|
|[getRange(address: string)](#getrangeaddress-string)|[Range](range.md)|Возвращает объект диапазона по адресу или имени.|
|[getUsedRange()](#getusedrange)|[Range](range.md)|Используемый диапазон — это наименьший диапазон, включающий все ячейки, которые содержат значение или форматирование. Если лист пустой, эта функция возвращает верхнюю левую ячейку.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе

### activate()
Активация листа в пользовательском интерфейсе Excel.

#### Синтаксис
```js
worksheetObject.activate();
```

#### Параметры
Нет

#### Возвращаемое значение
void

#### Примеры

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	worksheet.activate();
	return ctx.sync(); 
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### delete()
Удаляет лист из книги.

#### Синтаксис
```js
worksheetObject.delete();
```

#### Параметры
Нет

#### Возвращаемое значение
void

#### Примеры

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	worksheet.delete();
	return ctx.sync(); 
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getCell(row: number, column: number)
Получает объект диапазона, содержащий одну ячейку, по номеру строки и столбца. Ячейка может находиться вне родительского диапазона, если она расположена в таблице листа.

#### Синтаксис
```js
worksheetObject.getCell(row, column);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|row|number|Номер строки ячейки, которую требуется извлечь. Используется нулевой индекс.|
|column|number|Номер столбца ячейки, которую требуется извлечь. Используется нулевой индекс.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var cell = worksheet.getCell(0,0);
	cell.load('address');
	return ctx.sync().then(function() {
		console.log(cell.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getRange(address: string)
Возвращает объект диапазона по адресу или имени.

#### Синтаксис
```js
worksheetObject.getRange(address);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|address|string|Необязательный параметр. Адрес или имя диапазона. Если аргумент не указан, возвращается весь диапазон листа.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры
Этот пример использует адрес диапазона, чтобы получить соответствующий объект.

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.load('cellCount');
	return ctx.sync().then(function() {
		console.log(range.cellCount);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Этот пример использует именованный диапазон, чтобы получить соответствующий объект.

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeName = 'MyRange';
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeName);
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### getUsedRange()
Используемый диапазон — это наименьший диапазон, включающий в себя все ячейки с определенным значением или форматированием. Если лист пустой, эта функция вернет верхнюю левую ячейку.

#### Синтаксис
```js
worksheetObject.getUsedRange();
```

#### Параметры
Нет

#### Возвращаемое значение
[Range](range.md)

#### Примеры

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	var usedRange = worksheet.getUsedRange();
	usedRange.load('address');
	return ctx.sync().then(function() {
			console.log(usedRange.address);
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

Получение свойств листа на основе его имени.

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	worksheet.load('position')
	return ctx.sync().then(function() {
			console.log(worksheet.position);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Настройка положения листа. 

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	worksheet.position = 2;
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


