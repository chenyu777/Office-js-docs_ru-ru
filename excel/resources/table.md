# Объект Table (API JavaScript для Excel)

_Область применения: Excel 2016, Excel Online, Office 2016_

Представляет таблицу Excel.

## Свойства

| Свойство   | Тип|Описание
|:---------------|:--------|:----------|
|id|int|Возвращает значение, однозначно идентифицирующее таблицу в данной книге. Значение идентификатора остается прежним, даже если переименовать таблицу. Только для чтения.|
|name|string|Имя таблицы.|
|showHeaders|bool|Указывает, отображается ли строка заголовков. Можно задать это значение, чтобы отобразить или скрыть строку заголовков.|
|showTotals|bool|Указывает, отображается ли строка итогов. Можно задать это значение, чтобы отобразить или скрыть строку итогов.|
|style|string|Постоянное значение, представляющее стиль таблицы. Возможные значения: с TableStyleLight1 по TableStyleLight21, с TableStyleMedium1 по TableStyleMedium28, с TableStyleStyleDark1 по TableStyleStyleDark11. Кроме того, вы можете указать настраиваемый пользовательский стиль, присутствующий в книге.|

_Ознакомьтесь с [примерами](#property-access-examples) доступа к свойствам._

## Связи
| Связь | Тип|Описание|
|:---------------|:--------|:----------|
|столбцы|[TableColumnCollection](tablecolumncollection.md)|Представляет коллекцию всех столбцов в таблице. Только для чтения.|
|rows|[TableRowCollection](tablerowcollection.md)|Представляет коллекцию всех строк в таблице. Только для чтения.|

## Методы

| Метод   | Возвращаемый тип|Описание|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Удаляет таблицу.|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Получает объект диапазона, связанный с телом данных таблицы.|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|Получает объект диапазона, связанный со строкой заголовков таблицы.|
|[getRange()](#getrange)|[Range](range.md)|Получает объект диапазона, связанный со всей таблицей.|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|Получает объект диапазона, связанный со строкой итогов таблицы.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе

### delete()
Удаляет таблицу.

#### Синтаксис
```js
tableObject.delete();
```

#### Параметры
Нет

#### Возвращаемое значение
void

#### Примеры
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	table.delete();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getDataBodyRange()
Получает объект диапазона, связанный с телом данных таблицы.

#### Синтаксис
```js
tableObject.getDataBodyRange();
```

#### Параметры
Нет

#### Возвращаемое значение
[Range](range.md)

#### Примеры
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	var tableDataRange = table.getDataBodyRange();
	tableDataRange.load('address')
	return ctx.sync().then(function() {
			console.log(tableDataRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### getHeaderRowRange()
Получает объект диапазона, связанный со строкой заголовка таблицы.

#### Синтаксис
```js
tableObject.getHeaderRowRange();
```

#### Параметры
Нет

#### Возвращаемое значение
[Range](range.md)

#### Примеры
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	var tableHeaderRange = table.getHeaderRowRange();
	tableHeaderRange.load('address');
	return ctx.sync().then(function() {
		console.log(tableHeaderRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getRange()
Получает объект диапазона, связанный со всей таблицей.

#### Синтаксис
```js
tableObject.getRange();
```

#### Параметры
Нет

#### Возвращаемое значение
[Range](range.md)

#### Примеры
```js
Excel.run(function (ctx) { 
	var table = ctx.workbook.tables.getItem(tableName);
	var tableRange = table.getRange();
	tableRange.load('address');	
	return ctx.sync().then(function() {
			console.log(tableRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getTotalRowRange()
Получает объект диапазона, связанный со строкой итогов таблицы.

#### Синтаксис
```js
tableObject.getTotalRowRange();
```

#### Параметры
Нет

#### Возвращаемое значение
[Range](range.md)

#### Примеры
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	var tableTotalsRange = table.getTotalRowRange();
	tableTotalsRange.load('address');	
	return ctx.sync().then(function() {
			console.log(tableTotalsRange.address);
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

Получение таблицы по имени. 

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	table.load('index')
	return ctx.sync().then(function() {
			console.log(table.index);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Получение таблицы по индексу.

```js
Excel.run(function (ctx) { 
	var index = 0;
	var table = ctx.workbook.tables.getItemAt(0);
	table.name('name')
	return ctx.sync().then(function() {
			console.log(table.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Настройка стиля таблицы. 

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	table.name = 'Table1-Renamed';
	table.showTotals = false;
	table.tableStyle = 'TableStyleMedium2';
	table.load('tableStyle');
	return ctx.sync().then(function() {
			console.log(table.tableStyle);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
