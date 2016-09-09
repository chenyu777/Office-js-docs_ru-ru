# Объект Page (API JavaScript для OneNote)

_Относится к: OneNote Online_   


Представляет страницу OneNote.

## Свойства

| Свойство     | Тип   |Описание|Отзыв|
|:---------------|:--------|:----------|:-------|
|clientUrl|string|URL-адрес клиента страницы Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-clientUrl)|
|id|string|Получает идентификатор страницы. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-id)|
|pageLevel|int|Получает или задает уровень отступа для страницы.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-pageLevel)|
|title|string|Получает или задает название страницы.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-title)|
|webUrl|string|URL-адрес страницы. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-webUrl)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
| Связь | Тип   |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|contents|[PageContentCollection](pagecontentcollection.md)|Коллекция объектов PageContent на странице. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-contents)|
|inkAnalysisOrNull|[InkAnalysis](inkanalysis.md)|Интерпретация текста для рукописного ввода на странице. Возвращает значение null, если отсутствуют данные анализа рукописного фрагмента. Только для чтения. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-inkAnalysisOrNull)|
|parentSection|[Section](section.md)|Получает раздел, содержащий страницу. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-parentSection)|

## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[addOutline(left: double, top: double, html: String)](#addoutlineleft-double-top-double-html-string)|[Структура](outline.md)|Добавляет объект Outline на страницу в заданной позиции.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-addOutline)|
|[copyToSection(destinationSection: Section)](#copytosectiondestinationsection-section)|[Page](page.md)|Копирует эту страницу в указанный раздел.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-copyToSection)|
|[insertPageAsSibling(location: string, title: string)](#insertpageassiblinglocation-string-title-string)|[Page](page.md)|Вставляет новую страницу перед текущей страницей или после нее.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-insertPageAsSibling)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-load)|

## Сведения о методе


### addOutline(left: double, top: double, html: String)
Добавляет объект Outline на страницу в заданной позиции.

#### Синтаксис
```js
pageObject.addOutline(left, top, html);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|left|double|Левая позиция верхней части, левый угол объекта Outline.|
|top|double|Верхняя позиция верхней части, левый угол объекта Outline.|
|html|String|Строка HTML, описывающая визуальное представление объекта Outline. API JavaScript для надстроек OneNote см. в разделе [Поддерживаемые элементы HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html).|

#### Возвращаемое значение
[Outline](outline.md)

#### Примеры
```js
OneNote.run(function (context) {

    // Gets the active page.
    var page = context.application.getActivePage();

    // Queue a command to add an outline with given html. 
    var outline = page.addOutline(200, 200,
"<p>Images and a table below:</p> \
 <img src=\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==\"> \
 <img src=\"http://imagenes.es.sftcdn.net/es/scrn/6653000/6653659/microsoft-onenote-2013-01-535x535.png\"> \
 <table> \
   <tr> \
     <td>Jill</td> \
     <td>Smith</td> \
     <td>50</td> \
   </tr> \
   <tr> \
     <td>Eve</td> \
     <td>Jackson</td> \
     <td>94</td> \
   </tr> \
 </table>"     
        );

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
});
```


### copyToSection(destinationSection: Section)
Копирует эту страницу в указанный раздел.

#### Синтаксис
```js
pageObject.copyToSection(destinationSection);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|destinationSection|Section|Раздел, куда нужно скопировать эту страницу.|

#### Возвращаемое значение
[Page](page.md)

#### Примеры
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    
    // Gets the active notebook.
    var notebook = app.getActiveNotebook();
    
    // Gets the active page.
    var page = app.getActivePage();
    
    // Queue a command to load sections under the notebook.
    notebook.load('sections');
    
    var newPage;
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync()
        .then(function() {
            var section = notebook.sections.items[0];
            
            // copy page to the section.
            newPage = page.copyToSection(section);
            newPage.load('id');
            return ctx.sync();
        })
        .then(function() {
            console.log(newPage.id);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### insertPageAsSibling(location: string, title: string)
Вставляет новую страницу перед текущей страницей или после нее.

#### Синтаксис
```js
pageObject.insertPageAsSibling(location, title);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|location|string|Расположение новой страницы относительно текущей страницы.  Возможные значения: Before, After|
|должности.|string|Название новой страницы.|

#### Возвращаемое значение
[Page](page.md)

#### Примеры
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Queue a command to add a new page after the active page. 
    var newPage = activePage.insertPageAsSibling("After", "Next Page");

    // Queue a command to load the newPage to access its data.
    context.load(newPage);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log("page is created with title: " + newPage.title);
        });
})
.catch(function(error) {
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

**contents**
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Queue a command to add a new page after the active page. 
    var pageContents = activePage.contents;

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            for(var i=0; i < pageContents.items.length; i++)
            {
                var pageContent = pageContents.items[i];
                if (pageContent.type == "Outline")
                {
                    console.log("Found an outline");
                }
                else if (pageContent.type == "Image")
                {
                    console.log("Found an image");
                }
                else if (pageContent.type == "Other")
                {
                    console.log("Found a type not supported yet.");
                }
            }
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**webUrl**
```js
OneNote.run(function (context) {

    var app = context.application;
    
    // Gets the active page.
    var page = app.getActivePage();
    
    // Queue a command to load the webUrl of the page.
    page.load("webUrl");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log(page.webUrl);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**inkAnalysisOrNull**
```js
OneNote.run(function (ctx) {        
    var app = ctx.application;
    
    // Gets the active page.
    var page = app.getActivePage();
    
    // Load ink words
    page.load('inkAnalysisOrNull/paragraphs/lines/words');
    
    return ctx.sync()
        .then(function() {
            if (!page.inkAnalysisOrNull.isNull)
                console.log(page.inkAnalysisOrNull.paragraphs.length);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

