# Объект Binding (API JavaScript для Excel)

_Область применения: Excel 2016, Excel Online, Office 2016_

Представляет привязку Office.js, которая определена в книге.

## Свойства

| Свойство   | Тип|Описание
|:---------------|:--------|:----------|
|id|string|Представляет идентификатор привязки. Только для чтения.|
|type|string|Возвращает тип привязки. Только для чтения. Возможные значения: диапазон, Table, Text.|

_Ознакомьтесь с [примерами](#property-access-examples) доступа к свойствам._

## Связи
Нет


## Методы

| Метод   | Возвращаемый тип|Описание|
|:---------------|:--------|:----------|
|[getRange()](#getrange)|[Range](range.md)|Возвращает представленный привязкой диапазон. Если тип привязки неправильный, выдается ошибка.|
|[getTable()](#gettable)|[Table](table.md)|Возвращает представленную привязкой таблицу. Если тип привязки неправильный, выдается ошибка.|
|[getText()](#gettext)|string|Возвращает представленный привязкой текст. Если тип привязки неправильный, выдается ошибка.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе

### getRange()
Возвращает представленный привязкой диапазон. Если тип привязки неправильный, выдается ошибка.

#### Синтаксис
```js
bindingObject.getRange();
```

#### Параметры
Нет

#### Возвращаемое значение
[Range](range.md)

#### Примеры
Приведенный ниже пример получает связанный диапазон с помощью объекта привязки.

```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	var range = binding.getRange();
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

### getTable()
Возвращает представленную привязкой таблицу. Если тип привязки неправильный, выдается ошибка.

#### Синтаксис
```js
bindingObject.getTable();
```

#### Параметры
Нет

#### Возвращаемое значение
[Table](table.md)

#### Примеры
```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	var table = binding.getTable();
	table.load('name');
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

### getText()
Возвращает представленный привязкой текст. Если тип привязки неправильный, выдается ошибка.

#### Синтаксис
```js
bindingObject.getText();
```

#### Параметры
Нет

#### Возвращаемое значение
string

#### Примеры

```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	var text = binding.getText();
	ctx.load('text');
	return ctx.sync().then(function() {
		console.log(text);
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
	var binding = ctx.workbook.bindings.getItemAt(0);
	binding.load('type');
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

