# Объект TableRow (API JavaScript для Excel)

_Область применения: Excel 2016, Excel Online, Office 2016_

Представляет строку в таблице.

## Свойства

| Свойство   | Тип|Описание
|:---------------|:--------|:----------|
|index|int|Возвращает номер индекса строки в коллекции строк таблицы. Используется нулевой индекс. Только для чтения.|
|values|object[][]|Представляет необработанные значения указанного диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейка, которая содержит ошибку, возвращает строку ошибки.|

_Ознакомьтесь с [примерами](#property-access-examples) доступа к свойствам._

## Связи
Нет


## Методы

| Метод   | Возвращаемый тип|Описание|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Удаляет строку из таблицы.|
|[getRange()](#getrange)|[Range](range.md)|Получает объект диапазона, связанный со всей строкой.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе

### delete()
Удаляет строку из таблицы.

#### Синтаксис
```js
tableRowObject.delete();
```

#### Параметры
Нет

#### Возвращаемое значение
void

#### Примеры

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var row = ctx.workbook.tables.getItem(tableName).tableRows.getItemAt(2);
	row.delete();
	return ctx.sync(); 
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getRange()
Получает объект диапазона, связанный со всей строкой.

#### Синтаксис
```js
tableRowObject.getRange();
```

#### Параметры
Нет

#### Возвращаемое значение
[Range](range.md)

#### Примеры

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var row = ctx.workbook.tables.getItem(tableName).tableRows.getItemAt(0);
	var rowRange = row.getRange();
	rowRange.load('address');
	return ctx.sync().then(function() {
		console.log(rowRange.address);
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
	var tableName = 'Table1';
	var row = ctx.workbook.tables.getItem(tableName).tableRows.getItem(0);
	row.load('index');
	return ctx.sync().then(function() {
		console.log(row.index);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	var newValues = [["New", "Values", "For", "New", "Row"]];
	var row = ctx.workbook.tables.getItem(tableName).tableRows.getItemAt(2);
	row.values = newValues;
	row.load('values');
	return ctx.sync().then(function() {
		console.log(row.values);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
