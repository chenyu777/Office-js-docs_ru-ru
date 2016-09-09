# Объект Workbook (API JavaScript для Excel)

Workbook — объект верхнего уровня, содержащий связанные объекты книг, такие как листы, таблицы, диапазоны и т. д.

## Свойства

Нет

## Связи
| Связь | Тип   |Описание|
|:---------------|:--------|:----------|
|приложение|[Приложение](application.md)|Представляет экземпляр приложения Excel, содержащий эту книгу. Только для чтения.|
|bindings|[BindingCollection](bindingcollection.md)|Представляет коллекцию привязок, включенных в книгу. Только для чтения.|
|functions|[Функции](functions.md)|Представляет экземпляр приложения Excel, содержащий эту книгу. Только для чтения.|
|names|[NamedItemCollection](nameditemcollection.md)|Представляет коллекцию именованных элементов в книге (именованные диапазоны и константы). Только для чтения.|
|таблицы|[TableCollection](tablecollection.md)|Представляет коллекцию таблиц, сопоставленных с книгой. Только для чтения.|
|worksheets|[WorksheetCollection](worksheetcollection.md)|Представляет коллекцию листов, сопоставленных с книгой. Только для чтения.|

## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[getSelectedRange()](#getselectedrange)|[Range](range.md)|Получает текущий выделенный диапазон из книги.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе


### getSelectedRange()
Получает текущий выделенный диапазон из книги.

#### Синтаксис
```js
workbookObject.getSelectedRange();
```

#### Параметры
Нет

#### Возвращаемое значение
[Range](range.md)

#### Примеры

```js
Excel.run(function (ctx) { 
    var selectedRange = ctx.workbook.getSelectedRange();
    selectedRange.load('address');
    return ctx.sync().then(function() {
            console.log(selectedRange.address);
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
