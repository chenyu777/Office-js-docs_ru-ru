# Объект RequestContext (API JavaScript для Excel)

_Область применения: Excel 2016, Excel Online, Office 2016_

Объект RequestContext упрощает отправку запросов в приложение Excel. Так как надстройка Office и приложение Excel выполняются в виде двух отдельных процессов, контекст запроса необходим для получения доступа из надстройки к Excel и связанным объектам, например листам, таблицам и т. д. 

## Свойства
Нет

## Методы

| Метод         | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[load(object: object, option: object)](#loadobject-object-option-object)  |void     |Заполняет прокси-объект, созданный на уровне JavaScript, свойством и настройками, которые указаны в параметре.|

## Спецификации API

### load(object: object, option: object)
Заполняет прокси-объект, созданный на уровне JavaScript, свойством и настройками, которые указаны в параметре.

#### Синтаксис
```js
requestContextObject.load(object, loadOption);
```

#### Параметры
| Параметр       | Тип    |Описание|
|:----------------|:--------|:----------|
|object|object|Необязательный параметр. Укажите имя объекта, который необходимо загрузить.|
|option|[loadOption](loadoption.md)|Необязательный параметр. Укажите параметры загрузки, например select, expand, skip и top. Дополнительные сведения см. в статье, посвященной объекту loadOption.|

#### Возвращаемое значение
void

##### Примеры

Приведенный ниже пример загружает значения свойств из одного диапазона и копирует их в другой.

```js
Excel.run(function (ctx) { 
	var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");
	ctx.load(range, "values");
	return ctx.sync().then(function() {
		var myvalues=range.values;
		ctx.workbook.worksheets.getActiveWorksheet().getRange("B1:B2").values = myvalues;
		console.log(range.values);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
})
```

