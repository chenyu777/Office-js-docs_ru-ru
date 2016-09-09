# Объект WorksheetCollection (API JavaScript для Excel)

Представляет коллекцию объектов листа, включенных в книгу.

## Properties

| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|items|[Worksheet[]](worksheet.md)|Коллекция объектов листа. Только для чтения.|

_Ознакомьтесь с [примерами](#примерами) доступа к свойствам._

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[add(name: string)](#addname-string)|[Таблица](worksheet.md)|Добавляет новый лист в книгу. Лист будет добавлен после существующих листов. Чтобы активировать только что добавленный лист, вызовите метод activate().|
|[getActiveWorksheet()](#getactiveworksheet)|[Таблица](worksheet.md)|Получает текущий активный лист в книге.|
|[getItem(key: string)](#getitemkey-string)|[Таблица](worksheet.md)|Получает объект листа по его имени или ИД.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе


### add(name: string)
Добавляет новый лист в книгу. Лист будет добавлен после существующих листов. Чтобы активировать только что добавленный лист, вызовите метод activate().

#### Синтаксис
```js
worksheetCollectionObject.add(name);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|name|string|Необязательный параметр. Имя добавляемого листа. Если параметр используется, имя должно быть уникальным. В противном случае Excel определяет имя нового листа.|

#### Возвращаемое значение
[Таблица](worksheet.md)

#### Примеры

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sample Name';
    var worksheet = ctx.workbook.worksheets.add(wSheetName);
    worksheet.load('name');
    return ctx.sync().then(function() {
        console.log(worksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getActiveWorksheet()
Получает текущий активный лист в книге.

#### Синтаксис
```js
worksheetCollectionObject.getActiveWorksheet();
```

#### Параметры
Нет

#### Возвращаемое значение
[Таблица](worksheet.md)

#### Примеры

```js
Excel.run(function (ctx) {  
    var activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();
    activeWorksheet.load('name');
    return ctx.sync().then(function() {
            console.log(activeWorksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getItem(key: string)
Получает объект листа по его имени или ИД.

#### Синтаксис
```js
worksheetCollectionObject.getItem(key);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|key|string|Имя или ИД листа.|

#### Возвращаемое значение
[Таблица](worksheet.md)

### load(param: object)
Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.

#### Синтаксис
```js
object.load(param);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|param|object|Необязательный параметр. Принимает имена параметров и связей в виде строки с разделителями или массива. Либо укажите объект [loadOption](loadoption.md).|

#### Возвращаемое значение
void
### Примеры доступа к свойствам
```js
Excel.run(function (ctx) {
  var worksheets = ctx.workbook.worksheets;
  worksheets.load({"items" : "id, name"});
  return ctx.sync().then(function() {
    for (var i = 0; i < worksheets.items.length; i++)
    {
      console.log(worksheets.items[i].name);
      console.log(worksheets.items[i].id);
    }
  });
}).catch(function(error) {
  console.log("Error: " + error);
  if (error instanceof OfficeExtension.Error) {
    console.log("Debug info: " + JSON.stringify(error.debugInfo));
  }
});
```
