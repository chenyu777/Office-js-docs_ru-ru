# Объект NamedItemCollection (API JavaScript для Excel)

_Область применения: Excel 2016, Excel Online, Office 2016_

Коллекция всех объектов namedItem, которые включены в книгу.

## Свойства

| Свойство   | Тип|Описание
|:---------------|:--------|:----------|
|items|[NamedItem[]](nameditem.md)|Коллекция объектов namedItem. Только для чтения.|

_Ознакомьтесь с [примерами](#property-access-examples) доступа к свойствам._

## Связи
Нет


## Методы

| Метод   | Возвращаемый тип|Описание|
|:---------------|:--------|:----------|
|[getItem(name: string)](#getitemname-string)|[NamedItem](nameditem.md)|Возвращает объект namedItem по его имени.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе

### getItem(name: string)
Возвращает объект namedItem по его имени.

#### Синтаксис
```js
namedItemCollectionObject.getItem(name);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|name|string|Имя nameditem.|

#### Возвращаемое значение
[NamedItem](nameditem.md)

#### Примеры

```js
Excel.run(function (ctx) { 
	var nameditem = ctx.workbook.names.getItem(wSheetName);
	nameditem.load('type');
	return ctx.sync().then(function() {
			console.log(nameditem.type);
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
	var nameditems = ctx.workbook.names;
	nameditems.load('items');
	return ctx.sync().then(function() {
		for (var i = 0; i < nameditems.items.length; i++)
		{
			console.log(nameditems.items[i].name);
			console.log(nameditems.items[i].index);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Получение количества именованных элементов.

```js
Excel.run(function (ctx) { 
	var nameditems = ctx.workbook.names;
	nameditems.load('count');
	return ctx.sync().then(function() {
		console.log("nameditems: Count= " + nameditems.count);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


