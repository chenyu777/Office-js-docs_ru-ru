# Объект TableCollection (API JavaScript для Excel)

Представляет коллекцию всех таблиц, включенных в книгу.

## Properties

| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|count|int|Возвращает количество таблиц в книге. Только для чтения.|
|items|[Table[]](table.md)|Коллекция объектов таблицы. Только для чтения.|

_Ознакомьтесь с [примерами](#примерами) доступа к свойствам._

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[add(address: string, hasHeaders: bool)](#addaddress-string-hasheaders-bool)|[Таблица](table.md)|Создание таблицы. Исходный адрес диапазона определяет лист, на который будет добавлена таблица. Если добавить таблицу не удается (например, если адрес недействителен или одна таблица будет перекрываться другой), выводится сообщение об ошибке.|
|[getItem(key: number или string)](#getitemkey-number-or-string)|[Таблица](table.md)|Получает таблицу по имени или ИД.|
|[getItemAt(index: number)](#getitematindex-number)|[Таблица](table.md)|Получает таблицу на основании ее позиции в коллекции.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе


### add(address: string, hasHeaders: bool)
Создает таблицу. Исходный адрес диапазона определяет лист, на который будет добавлена таблица. Если добавить таблицу не удается (например, если адрес недействителен или одна таблица будет перекрываться другой), выводится сообщение об ошибке.

#### Синтаксис
```js
tableCollectionObject.add(address, hasHeaders);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|address|string|Адрес или имя объекта диапазона, представляющего источник данных. Если адрес не содержит имя листа, используется текущий активный лист.|
|hasHeaders|bool|Логическое значение, указывающее, есть ли у импортируемых данных метки столбцов. Если источник не содержит заголовков (например, если этому свойству присвоено значение false), Excel автоматически создаст заголовок и сдвинет данные на одну строку вниз.|

#### Возвращаемое значение
[Таблица](table.md)

#### Примеры

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.add('Sheet1!A1:E7', true);
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

### getItem(key: number или string)
Получает таблицу по имени или ИД.

#### Синтаксис
```js
tableCollectionObject.getItem(key);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|key|number или string|Имя или ИД получаемой таблицы.|

#### Возвращаемое значение
[Таблица](table.md)

#### Примеры

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
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


#### Примеры

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItemAt(0);
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


### getItemAt(index: number)
Получает таблицу на основании ее позиции в коллекции.

#### Синтаксис
```js
tableCollectionObject.getItemAt(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number|Значение индекса получаемого объекта. Используется нулевой индекс.|

#### Возвращаемое значение
[Таблица](table.md)

#### Примеры

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItemAt(0);
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
    var tables = ctx.workbook.tables;
    tables.load('items');
    return ctx.sync().then(function() {
        console.log("tables Count: " + tables.count);
        for (var i = 0; i < tables.items.length; i++)
        {
            console.log(tables.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Получение количества таблиц

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    tables.load('count');
    return ctx.sync().then(function() {
        console.log(tables.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
