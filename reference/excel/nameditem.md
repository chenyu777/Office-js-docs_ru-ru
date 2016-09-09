# Объект NamedItem (API JavaScript для Excel)

Представляет определенное имя для диапазона ячеек или значения. Имена могут быть простыми именованными объектами (как показано ниже в столбце "Тип,"), объектом диапазона и ссылкой на диапазон. Этот объект может использоваться для получения объекта диапазона, связанного с именами.

## Свойства

| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|name|string|Имя объекта. Только для чтения.|
|type|string|Указывает тип ссылки, связанный с именем. Только для чтения. Возможные значения: String, Integer, Double, Boolean, Range.|
|value|object|Представляет формулу, на которую ссылается имя, например =Sheet14!$B$2:$H$12, =4.75 и т. д. Только для чтения.|
|visible|bool|Определяет, является ли объект видимым.|

_Ознакомьтесь с [примерами](#примерами) доступа к свойствам._

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[getRange()](#getrange)|[Range](range.md)|Возвращает объект Range, сопоставленный с именем. Вызывает исключение, если тип именованного элемента не является диапазоном.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе


### getRange()
Возвращает объект Range, сопоставленный с именем. Вызывает исключение, если тип именованного элемента не является диапазоном.

#### Синтаксис
```js
namedItemObject.getRange();
```

#### Параметры
Нет

#### Возвращаемое значение
[Range](range.md)

#### Примеры

Возвращает объект диапазона, связанный с именем. Если тип имени отличается от `null`, возвращается значение `Range`. Примечание: этот API в настоящее время поддерживает только элементы в области книги.

```js
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var range = names.getItem('MyRange').getRange();
    range.load('address');
    return ctx.sync().then(function() {
            console.log(range.address);
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
    var names = ctx.workbook.names;
    var namedItem = names.getItem('MyRange');
    namedItem.load('type');
    return ctx.sync().then(function() {
            console.log(namedItem.type);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
