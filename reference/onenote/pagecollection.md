# Объект PageCollection (API JavaScript для OneNote)

_Относится к: OneNote Online_  


Представляет коллекцию страниц.

## Свойства

| Свойство     | Тип   |Описание|Отзывы|
|:---------------|:--------|:----------|:-------|
|count|int|Возвращает количество страниц в коллекции. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-count)|
|items|[Page[]](page.md)|Коллекция объектов Page. Только для чтения. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-items)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[getByTitle(title: string)](#getbytitletitle-string)|[PageCollection](pagecollection.md)|Возвращает коллекцию объектов Page с указанным названием.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-getByTitle)|
|[getItem(index: числовой тип или string)](#getitemindex-number-or-string)|[Page](page.md)|Возвращает объект Page по его идентификатору или индексу в коллекции. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[Page](page.md)|Возвращает объект Page по позиции в коллекции.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageCollection-load)|

## Сведения о методе


### getByTitle(title: string)
Возвращает коллекцию объектов Page с указанным названием.

#### Синтаксис
```js
pageCollectionObject.getByTitle(title);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|title|string|Заголовок страницы.|

#### Возвращаемое значение
[PageCollection](pagecollection.md)

#### Примеры
```js
OneNote.run(function (context) {

    // Get all the pages in the current section.
    var allPages = context.application.getActiveSection().pages;

    // Queue a command to load the pages. 
    // For best performance, request specific properties.
    allPages.load("id"); 

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Get the sections with the specified name.
            var todoPages = allPages.getByTitle("Todo list");

            // Queue a command to load the section. 
            // For best performance, request specific properties.
            todoPages.load("id,title"); 

            return context.sync()
                .then(function () {

                    // Iterate through the collection or access items individually by index.
                    if (todoPages.items.length > 0) {
                        console.log("Page title: " + todoPages.items[0].title);
                        console.log("Page ID: " + todoPages.items[0].id);
                    }
                });
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### getItem(index: числовой тип или string)
Возвращает объект Page по его идентификатору или индексу в коллекции. Только для чтения.

#### Синтаксис
```js
pageCollectionObject.getItem(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number или string|Идентификатор объекта Page или расположение индекса этого объекта в коллекции.|

#### Возвращаемое значение
[Page](page.md)

### getItemAt(index: number)
Возвращает объект Page по его позиции в коллекции.

#### Синтаксис
```js
pageCollectionObject.getItemAt(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number|Значение индекса получаемого объекта. Используется нулевой индекс.|

#### Возвращаемое значение
[Page](page.md)

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

**items**
```js
OneNote.run(function (context) {
    
    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;
    
    // Queue a command to load the id and title for each page.            
    pages.load('id,title');
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Display the properties.
            $.each(pages.items, function(index, page) {
                console.log(page.title);
                console.log(page.id);
            });
        }); 
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

