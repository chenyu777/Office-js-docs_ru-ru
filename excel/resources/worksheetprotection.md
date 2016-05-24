# Объект WorksheetProtection (API JavaScript для Excel)

_Относится к: Excel 2016, Excel Online, Excel для iOS, Office 2016_

Представляет защиту объекта Sheet.

## Свойства

| Свойство   | Тип|Описание
|:---------------|:--------|:----------|
|protected|bool|Указывает, защищен ли лист. Только для чтения.|

## Связи
| Связь | Тип|Описание|
|:---------------|:--------|:----------|
|options|[WorksheetProtectionOptions](worksheetprotectionoptions.md)|Параметры защиты листа. Только для чтения.|

## Методы

| Метод   | Возвращаемый тип|Описание|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект сведениями о защите листа.|
|[protect(options: WorksheetProtectionOptions)](#protectoptions-worksheetprotectionoption)|void|Защита листа. Выдает исключение, если лист защищен.|
|[unprotect()](#unprotect)|void|Снятие защиты с листа|

## Сведения о методе


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

#### Примеры
В этом примере загружаются сведения о защите активного листа.
```js
Excel.run(function (ctx) {
    var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
    worksheet.protection.load();            
    return ctx.sync()
        .then(function () {
            console.log("Active worksheet's protection status: " + worksheet.protection.protected);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### protect(options: WorksheetProtectionOptions)
Защищает лист с помощью необязательных политик защиты. Выдает исключение, если лист защищен. 

Если параметры указаны, можно включить или отключить отдельные политики. Если политика не указана, она включена по умолчанию. 

#### Синтаксис
```js
worksheetProtectionObject.protect(options);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|options|WorksheetProtectionOptions|Необязательные параметры защиты листа.|


#### Возвращаемое значение
void

#### Примеры
```js
Excel.run(function (ctx) { 
	var sheet = ctx.workbook.worksheets.getItem("Sheet1");
	var range = sheet.getRange("A1:B3").format.protection.locked = false;
	sheet.protection.protect({allowInsertRows:true});
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});

```
### unprotect()
Снятие защиты с листа. 

#### Синтаксис
```js
worksheetProtectionObject.unprotect();
```

#### Параметры
Нет

#### Возвращаемое значение
void

#### Примеры
```js
Excel.run(function (ctx) { 
	var sheet = ctx.workbook.worksheets.getItem("Sheet1");	
	sheet.protection.unprotect();
	return ctx.sync(); 
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
