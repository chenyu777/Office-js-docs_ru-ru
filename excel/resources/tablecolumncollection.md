# Объект TableColumnCollection (API JavaScript для Excel)

_Область применения: Excel 2016, Excel Online, Office 2016_

Представляет коллекцию всех столбцов, включенных в таблицу.

## Свойства

| Свойство   | Тип|Описание
|:---------------|:--------|:----------|
|count|int|Возвращает количество столбцов в таблице. Только для чтения.|
|items|[TableColumn[]](tablecolumn.md)|Коллекция объектов tableColumn. Только для чтения.|

_Ознакомьтесь с [примерами](#property-access-examples) доступа к свойствам._

## Связи
Нет


## Методы

| Метод   | Возвращаемый тип|Описание|
|:---------------|:--------|:----------|
|[add(index: number, values: (boolean, string или number)[][])](#addindex-number-values-boolean-or-string-or-number)|[TableColumn](tablecolumn.md)|Добавляет новый столбец в таблицу.|
|[getItem(key: number или string)](#getitemkey-number-or-string)|[TableColumn](tablecolumn.md)|Возвращает объект столбца по имени или идентификатору.|
|[getItemAt(index: number)](#getitematindex-number)|[TableColumn](tablecolumn.md)|Возвращает столбец на основании его позиции в коллекции.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе

### add(index: number, values: (boolean, string или number)[][])
Добавляет новый столбец в таблицу.

#### Синтаксис
```js
tableColumnCollectionObject.add(index, values);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|index|number|Определяет относительную позицию нового столбца. Предыдущий столбец на этой позиции сдвигается вправо. Значение индекса не должно превышать значение индекса последнего столбца, поэтому его невозможно использовать для добавления столбца в конце таблицы. Используется нулевой индекс.|
|values|(boolean, string или number)[][]|Необязательный параметр. Двухмерный массив неформатированных значений столбца таблицы.|

#### Возвращаемое значение
[TableColumn](tablecolumn.md)

#### Примеры

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	var values = [["Sample"], ["Values"], ["For"], ["New"], ["Column"]];
	var column = tables.getItem("Table1").columns.add(null, values);
	column.load('name');
	return ctx.sync().then(function() {
		console.log(column.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getItem(key: number или string)
Возвращает объект столбца по имени или идентификатору.

#### Синтаксис
```js
tableColumnCollectionObject.getItem(key);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|key|number или string| Имя или ИД столбца.|

#### Возвращаемое значение
[TableColumn](tablecolumn.md)

#### Примеры

```js
Excel.run(function (ctx) { 
	var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItem(0);
	tablecolumn.load('name');
	return ctx.sync().then(function() {
			console.log(tablecolumn.name);
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
	var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
	tablecolumn.load('name');
	return ctx.sync().then(function() {
			console.log(tablecolumn.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### getItemAt(index: number)
Возвращает столбец на основании его позиции в коллекции.

#### Синтаксис
```js
tableColumnCollectionObject.getItemAt(index);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|index|number|Значение индекса получаемого объекта. Используется нулевой индекс.|

#### Возвращаемое значение
[TableColumn](tablecolumn.md)

#### Примеры
```js
Excel.run(function (ctx) { 
	var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
	tablecolumn.load('name');
	return ctx.sync().then(function() {
			console.log(tablecolumn.name);
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
	var tablecolumns = ctx.workbook.tables.getItem['Table1'].columns;
	tablecolumns.load('items');
	return ctx.sync().then(function() {
		console.log("tablecolumns Count: " + tablecolumns.count);
		for (var i = 0; i < tablecolumns.items.length; i++)
		{
			console.log(tablecolumns.items[i].name);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
