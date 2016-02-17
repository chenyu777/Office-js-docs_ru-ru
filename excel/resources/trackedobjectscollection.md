# Объект TrackedObjectsCollection (API JavaScript для Office 2016)

_Область применения: Excel 2016, Excel Online, Office 2016_

Позволяет надстройкам управлять ссылками на объект диапазона в пакетах sync(). Обычно Excel.run() позволяет сохранять ссылки в пакетах автоматически без необходимости их явного отслеживания. Тем не менее, если для сценария надстройки требуется, чтобы объект диапазона отслеживался и корректировался вручную для отражения текущего состояния базового диапазона Excel, то с помощью этой коллекции можно пометить такие объекты для отслеживания. Обратите внимание, что если объект диапазона помечен для отслеживания, его необходимо явным образом удалить после использования, чтобы освободить память в Excel, даже в случае ошибки.

## Свойства
Нет

## Связи

Нет

## Методы

Для объекта trackedObjectsCollection определены следующие методы:

| Метод     | Возвращаемый тип    |Описание|
|:-----------------|:--------|:----------|
|[add(rangeObject: Range)](#addrangeobject-range)| Null             |Создает новую ссылку на диапазон.|
|[remove(rangeObject: Range)](#removerangeobject-range)| Null             |Удаление ссылки на диапазон.  |
|[removeAll()](#removeall)| Null|Удаляет все ссылки, созданные надстройкой на устройстве.|


## Спецификации API 

### add(rangeObject: range)
Добавление объекта диапазона к trackedObjectsCollection. Будут отслеживаться любые базовые изменения в пакетах, а все последующие обновления будут применены к текущему состоянию объекта диапазона. 

#### Синтаксис
```js
trackedObjectsCollection.add(rangeObject);
```

#### Параметры

Параметр       | Тип   | Описание
--------------- | ------ | ------------
`rangeObject`  | [Range](range.md)| Объект диапазона, который нужно добавить к trackedObjectCollection.

#### Возвращаемое значение
Null

#### Примеры

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.trackedObjectsCollection.add(range);
ctx.load(range);

Excel.run(function (ctx) { 
	range.insert("Down");
	Console.log(range.address); // Address should be updated to A3:B4
	return ctx.sync(); 
});
```


### remove(rangeObject: range)

Удаление объекта ссылки из коллекции. Высвобождает память и ресурсы, необходимые для сохранения состояния отслеживаемого объекта. Обратите внимание, что если объект диапазона помечен для отслеживания, его необходимо явным образом удалить, даже в случае ошибки.

#### Синтаксис
```js
trackedObjectsCollection.remove(rangeObject);
```

#### Параметры

Параметр       | Тип   | Описание
--------------- | ------ | ------------
`rangeObject`  | [Range](range.md)| Объект диапазона, который требуется удалить из trackedObjectCollection.

#### Возвращаемое значение
Null

#### Примеры


```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.trackedObjectsCollection.add(range);
ctx.load(range);

Excel.run(function (ctx) { 
	range.insert("Down");
	Console.log(range.address); // Address should be updated to A3:B4
	ctx.trackedObjectsCollection.remove(range); 
	return ctx.sync(); 
});
```

### removeAll(rangeObject: range)

Удаляет все ссылки, созданные надстройкой на устройстве.

#### Синтаксис
```js
trackedObjectsCollection.removeAll();
```

#### Параметры

Нет

#### Возвращаемое значение
Null

#### Примеры

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:B2";
	var ctx = new Excel.RequestContext();
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	ctx.trackedObjectsCollection.add(range);
	ctx.load(range);
	range.insert("Down");
	Console.log(range.address); // Address should be updated to A3:B4
	ctx.trackedObjectsCollection.removeAll(); 
	return ctx.sync(); 
});
```

