# Объект Application (API JavaScript для OneNote)

_Область применения: OneNote Online_


Представляет собой объект верхнего уровня и содержит все глобально адресуемые объекты OneNote, например записные книжки, активную записную книжку и активный раздел.

## Свойства

Нет

## Связи
| Связь | Тип   |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|notebooks|[NotebookCollection](notebookcollection.md)|Получает коллекцию записных книжек, открытых в экземпляре приложения OneNote. В OneNote Online в экземпляре приложения может быть открыто не более одной записной книжки одновременно. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-notebooks)|

## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[getActiveNotebook()](#getactivenotebook)|[Notebook](notebook.md)|Получает активную записную книжку, если она есть. Если такой записной книжки нет, создается исключение ItemNotFound.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveNotebook)|
|[getActiveNotebookOrNull()](#getactivenotebookornull)|[Notebook](notebook.md)|Получает активную записную книжку, если она есть. Если такой записной книжки нет, возвращается значение null.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveNotebookOrNull)|
|[getActiveOutline()](#getactiveoutline)|[Outline](outline.md)|Возвращает активную структуру, если она есть. Если такой структуры нет, создается исключение ItemNotFound.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveOutline)|
|[getActiveOutlineOrNull()](#getactiveoutlineornull)|[Outline](outline.md)|Возвращает активную структуру, если она есть. Если такой нет, возвращается значение null.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveOutlineOrNull)|
|[getActivePage()](#getactivepage)|[Page](page.md)|Возвращает активную страницу, если она есть. Если такой страницы нет, создается исключение ItemNotFound.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActivePage)|
|[getActivePageOrNull()](#getactivepageornull)|[Page](page.md)|Возвращает активную страницу, если она есть. Если активной страницы нет, возвращается значение null.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActivePageOrNull)|
|[getActiveSection()](#getactivesection)|[Section](section.md)|Возвращает активный раздел, если он есть. Если такого раздела нет, создается исключение ItemNotFound.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveSection)|
|[getActiveSectionOrNull()](#getactivesectionornull)|[Section](section.md)|Возвращает активный раздел, если он есть. Если такого раздела нет, возвращается значение null.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-getActiveSectionOrNull)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-load)|
|[navigateToPage(page: Page)](#navigatetopagepage-page)|void|Открывает указанную страницу в экземпляре приложения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-navigateToPage)|
|[navigateToPageWithClientUrl(url: string)](#navigatetopagewithclienturlurl-string)|[Page](page.md)|Возвращает указанную страницу и открывает ее в экземпляре приложения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-application-navigateToPageWithClientUrl)|

## Сведения о методе


### getActiveNotebook()
Получает активную записную книжку, если она есть. Если такой записной книжки нет, создается исключение ItemNotFound.

#### Синтаксис
```js
applicationObject.getActiveNotebook();
```

#### Параметры
Нет

#### Возвращаемое значение
[Notebook](notebook.md)

#### Примеры
```js
OneNote.run(function (context) {
        
    // Get the active notebook.
    var notebook = context.application.getActiveNotebook();
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Notebook name: " + notebook.name);
            console.log("Notebook ID: " + notebook.id);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### getActiveNotebookOrNull()
Получает активную записную книжку, если она есть. Если такой записной книжки нет, возвращается значение null.

#### Синтаксис
```js
applicationObject.getActiveNotebookOrNull();
```

#### Параметры
Нет

#### Возвращаемое значение
[Notebook](notebook.md)

#### Примеры
```js
OneNote.run(function (context) {

    // Get the active notebook.
    var notebook = context.application.getActiveNotebookOrNull();

    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('id,name');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // check if active notebook is set.
            if (!notebook.isNull) {
                console.log("Notebook name: " + notebook.name);
                console.log("Notebook ID: " + notebook.id);
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


### getActiveOutline()
Возвращает активную структуру, если она есть. Если такой структуры нет, создается исключение ItemNotFound.

#### Синтаксис
```js
applicationObject.getActiveOutline();
```

#### Параметры
Нет

#### Возвращаемое значение
[Outline](outline.md)

#### Примеры
```js
OneNote.run(function (context) {

    // get active outline.
    var outline = context.application.getActiveOutline();

    // Queue a command to load the id of the outline.         
    outline.load('id');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Show some properties.
            console.log("outline id: " + outline.id);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### getActiveOutlineOrNull()
Возвращает активную структуру, если она есть. Если такой нет, возвращается значение null.

#### Синтаксис
```js
applicationObject.getActiveOutlineOrNull();
```

#### Параметры
Нет

#### Возвращаемое значение
[Outline](outline.md)

#### Примеры
```js
OneNote.run(function (context) {

    // get active outline.
    var outline = context.application.getActiveOutlineOrNull();

    // Queue a command to load the id of the outline.         
    outline.load('id');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            if (!outline.isNull) {
                console.log("outline id: " + outline.id);
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


### getActivePage()
Возвращает активную страницу, если она есть. Если такой страницы нет, создается исключение ItemNotFound.

#### Синтаксис
```js
applicationObject.getActivePage();
```

#### Параметры
Нет

#### Возвращаемое значение
[Page](page.md)

#### Примеры
```js
OneNote.run(function (context) {
        
    // Get the active page.
    var page = context.application.getActivePage();
            
    // Queue a command to load the page. 
    // For best performance, request specific properties.           
    page.load('id,title');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Page title: " + page.title);
            console.log("Page ID: " + page.id);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### getActivePageOrNull()
Возвращает активную страницу, если она есть. Если активной страницы нет, возвращается значение null.

#### Синтаксис
```js
applicationObject.getActivePageOrNull();
```

#### Параметры
Нет

#### Возвращаемое значение
[Page](page.md)

#### Примеры
```js
OneNote.run(function (context) {

    // Get the active page.
    var page = context.application.getActivePageOrNull();

    // Queue a command to load the page. 
    // For best performance, request specific properties.           
    page.load('id,title');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            if (!page.isNull) {
                // Show some properties.
                console.log("Page title: " + page.title);
                console.log("Page ID: " + page.id);
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


### getActiveSection()
Возвращает активный раздел, если он есть. Если такого раздела нет, создается исключение ItemNotFound.

#### Синтаксис
```js
applicationObject.getActiveSection();
```

#### Параметры
Нет

#### Возвращаемое значение
[Section](section.md)

#### Примеры
```js
OneNote.run(function (context) {
        
    // Get the active section.
    var section = context.application.getActiveSection();
            
    // Queue a command to load the section. 
    // For best performance, request specific properties.           
    section.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // Show some properties.
            console.log("Section name: " + section.name);
            console.log("Section ID: " + section.id);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### getActiveSectionOrNull()
Возвращает активный раздел, если он есть. Если такого раздела нет, возвращается значение null.

#### Синтаксис
```js
applicationObject.getActiveSectionOrNull();
```

#### Параметры
Нет

#### Возвращаемое значение
[Section](section.md)

#### Примеры
```js
OneNote.run(function (context) {

    // Get the active section.
    var section = context.application.getActiveSectionOrNull();

    // Queue a command to load the section. 
    // For best performance, request specific properties.           
    section.load('id,name');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            if (!section.isNull) {
                // Show some properties.
                console.log("Section name: " + section.name);
                console.log("Section ID: " + section.id);
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

### navigateToPage(page: Page)
Открывает указанную страницу в экземпляре приложения.

#### Синтаксис
```js
applicationObject.navigateToPage(page);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|page|Page|Страница, которую необходимо открыть.|

#### Возвращаемое значение
void

#### Примеры
```js        
OneNote.run(function (context) {
        
    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;
            
    // Queue a command to load the pages. 
    // For best performance, request specific properties.           
    pages.load('id');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
                    
            // This example loads the first page in the section.
            var page = pages.items[0];
                        
            // Open the page in the application.                    
            context.application.navigateToPage(page);
                    
            // Run the queued command.
            return context.sync();
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### navigateToPageWithClientUrl(url: string)
Возвращает указанную страницу и открывает ее в экземпляре приложения.

#### Синтаксис
```js
applicationObject.navigateToPageWithClientUrl(url);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|url|string|URL-адрес клиента страницы, которую необходимо открыть.|

#### Возвращаемый объект
[Page](page.md)

#### Примеры
```js
OneNote.run(function (context) {

    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;

    // Queue a command to load the pages. 
    // For best performance, request specific properties.           
    pages.load('clientUrl');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // This example loads the first page in the section.
            var page = pages.items[0];

            // Open the page in the application.                    
            context.application.navigateToPageWithClientUrl(page.clientUrl);

            // Run the queued command.
            return context.sync();
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
