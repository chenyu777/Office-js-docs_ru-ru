# Объект BindingCollection (API JavaScript для Excel)

_Область применения: Excel 2016, Excel Online, Office 2016_

Представляет коллекцию всех объектов привязки, включенных в книгу.

## Свойства

| Свойство   | Тип|Описание
|:---------------|:--------|:----------|
|count|int|Возвращает число привязок в коллекции. Только для чтения.|
|items|[Binding[]](binding.md)|Коллекция объектов привязки. Только для чтения.|

_Ознакомьтесь с [примерами](#property-access-examples) доступа к свойствам._

## Связи
Нет


## Методы

| Метод   | Возвращаемый тип|Описание|
|:---------------|:--------|:----------|
|[getItem(id: string)](#getitemid-string)|[Binding](binding.md)|Возвращает объект привязки по идентификатору.|
|[getItemAt(index: number)](#getitematindex-number)|[Binding](binding.md)|Возвращает объект привязки с учетом его положения в массиве элементов.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе

### getItem(id: string)
Возвращает объект привязки по идентификатору.

#### Синтаксис
```js
bindingCollectionObject.getItem(id);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|id|string|Идентификатор получаемого объекта привязки.|

#### Возвращаемое значение
[Binding](binding.md)

#### Примеры

Создайте привязку таблицы, чтобы отслеживать изменения данных в этой таблице. При изменении данных фон таблицы станет оранжевым.

```js
function addEventHandler() {
	//Create Table1
Excel.run(function (ctx) { 
	ctx.workbook.tables.add("Sheet1!A1:C4", true);
	return ctx.sync().then(function() {
			 console.log("My Diet Data Inserted!");
	})
	.catch(function (error) {
			 console.log(JSON.stringify(error));
	});
});
	//Create a new table binding for Table1
Office.context.document.bindings.addFromNamedItemAsync("Table1", Office.CoercionType.Table, { id: "myBinding" }, function (asyncResult) {
	if (asyncResult.status == "failed") {
		console.log("Action failed with error: " + asyncResult.error.message);
	}
	else {
		// If successful, add the event handler to the table binding.
		Office.select("bindings#myBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
	}
});
}
	
// When data in the table is changed, this event is triggered.
function onBindingDataChanged(eventArgs) {
Excel.run(function (ctx) { 
	// Highlight the table in orange to indicate data changed.
	ctx.workbook.bindings.getItem(eventArgs.binding.id).getTable().getDataBodyRange().format.fill.color = "Orange";
	return ctx.sync().then(function() {
			console.log("The value in this table got changed!");
	})
	.catch(function (error) {
			console.log(JSON.stringify(error));
	});
});
}

```



#### Примеры
```js
Excel.run(function (ctx) { 
	var lastPosition = ctx.workbook.bindings.count - 1;
	var binding = ctx.workbook.bindings.getItemAt(lastPosition);
	binding.load('type')
	return ctx.sync().then(function() {
			console.log(binding.type); 
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getItemAt(index: number)
Возвращает объект привязки с учетом его положения в массиве элементов.

#### Синтаксис
```js
bindingCollectionObject.getItemAt(index);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|index|number|Значение индекса получаемого объекта. Используется нулевой индекс.|

#### Возвращаемое значение
[Binding](binding.md)

#### Примеры
```js
Excel.run(function (ctx) { 
	var lastPosition = ctx.workbook.bindings.count - 1;
	var binding = ctx.workbook.bindings.getItemAt(lastPosition);
	binding.load('type')
	return ctx.sync().then(function() {
			console.log(binding.type); 
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
|param|object|Необязательный параметр. Принимает имена параметров и связей в виде строки с разделителями или массива. Этот параметр также принимает объект [loadOption](loadoption.md).|

#### Возвращаемое значение
void

### Примеры доступа к свойствам

```js
Excel.run(function (ctx) { 
	var bindings = ctx.workbook.bindings;
	bindings.load('items');
	return ctx.sync().then(function() {
		for (var i = 0; i < bindings.items.length; i++)
		{
			console.log(bindings.items[i].id);
			console.log(bindings.items[i].index);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
Получение количества привязок.

```js
Excel.run(function (ctx) { 
	var bindings = ctx.workbook.bindings;
	bindings.load('count');
	return ctx.sync().then(function() {
		console.log("Bindings: Count= " + bindings.count);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

