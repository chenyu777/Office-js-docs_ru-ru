# Объект Section (API JavaScript для OneNote)

_Относится к: OneNote Online_   


Представляет раздел в OneNote. Разделы могут содержать страницы.

## Свойства

| Свойство     | Тип   |Описание|Отзыв|
|:---------------|:--------|:----------|:-------|
|clientUrl|string|URL-адрес клиента раздела. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-clientUrl)|
|id|string|Получает идентификатор объекта Section. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-id)|
|name|string|Получает имя раздела. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-name)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
| Связь | Тип   |Описание| Отзывы|
|:---------------|:--------|:----------|:-------|
|notebook|[Notebook](notebook.md)|Получает записную книжку, содержащую раздел. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-notebook)|
|pages|[PageCollection](pagecollection.md)|Получает коллекцию страниц в разделе. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-pages)|
|parentSectionGroup|[SectionGroup](sectiongroup.md)|Получает группу объектов, содержащую раздел. Возвращает значение ItemNotFound, если раздел является прямым потомком записной книжки. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-parentSectionGroup)|
|parentSectionGroupOrNull|[SectionGroup](sectiongroup.md)|Получает группу объектов, содержащую раздел. Возвращает значение null, если объект Section является прямым потомком объекта Notebook. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-parentSectionGroupOrNull)|

## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[addPage(title: string)](#addpagetitle-string)|[Page](page.md)|Добавляет новую страницу в конец раздела.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-addPage)|
|[copyToNotebook(destinationNotebook: Notebook)](#copytonotebookdestinationnotebook-notebook)|[Section](section.md)|Копирует этот раздел в указанную записную книжку.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-copyToNotebook)|
|[copyToSectionGroup(destinationSectionGroup: SectionGroup)](#copytosectiongroupdestinationsectiongroup-sectiongroup)|[Section](section.md)|Копирует этот раздел в указанную группу разделов.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-copyToSectionGroup)|
|[insertSectionAsSibling(location: string, title: string)](#insertsectionassiblinglocation-string-title-string)|[Раздел](section.md)|Вставляет новый раздел перед текущим разделом или после него.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-insertSectionAsSibling)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-section-load)|

## Сведения о методе


### addPage(title: string)
Добавляет новую страницу в конец раздела.

#### Синтаксис
```js
sectionObject.addPage(title);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|title|string|Название новой страницы.|

#### Возвращаемое значение
[Page](page.md)

#### Примеры
```js
OneNote.run(function (context) {
            
    // Queue a command to add a page to the current section.
    var page = context.application.getActiveSection().addPage("Wish list");
            
    // Queue a command to load the id and title of the new page. 
    // This example loads the new page so it can read its properties later.           
    page.load('id,title');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
             
            // Display the properties.       
            console.log("Page name: " + page.title);
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


### copyToNotebook(destinationNotebook: Notebook)
Копирует этот раздел в указанную записную книжку.

#### Синтаксис
```js
sectionObject.copyToNotebook(destinationNotebook);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|destinationNotebook|Notebook|Записная книжка, куда нужно скопировать этот раздел.|

#### Возвращаемое значение
[Раздел](section.md)

#### Примеры
```js
OneNote.run(function (context) {
    var app = context.application;
    
    // Gets the active Notebook.
    var notebook = app.getActiveNotebook();
    
    // Gets the active Section.
    var section = app.getActiveSection();
    
    var newSection;
    
    return context.sync()
        .then(function() {
            newSection = section.copyToNotebook(notebook);
            newSection.load('id');
            return context.sync();
        })
        .then(function() {
            console.log(newSection.id);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### copyToSectionGroup(destinationSectionGroup: SectionGroup)
Копирует этот раздел в указанную группу разделов.

#### Синтаксис
```js
sectionObject.copyToSectionGroup(destinationSectionGroup);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|destinationSectionGroup|SectionGroup|Группа разделов, куда нужно скопировать этот раздел.|

#### Возвращаемое значение
[Раздел](section.md)

#### Примеры
```js
OneNote.run(function (ctx) {
    var app = ctx.application;
    
    // Gets the active Notebook.
    var notebook = app.getActiveNotebook();
    
    // Gets the active Section.
    var section = app.getActiveSection();
    
    var newSection;
    
    return ctx.sync()
        .then(function() {
            var firstSectionGroup = notebook.sectionGroups.items[0];
            newSection = section.copyToSectionGroup(firstSectionGroup);
            newSection.load('id');
            return ctx.sync();
        })
        .then(function() {
            console.log(newSection.id);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### insertSectionAsSibling(location: string, title: string)
Вставляет новый раздел перед текущим разделом или после него.

#### Синтаксис
```js
sectionObject.insertSectionAsSibling(location, title);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|location|string|Расположение нового раздела относительно текущего раздела.  Возможные значения: Before, After|
|должности.|string|Имя нового раздела.|

#### Возвращаемое значение
[Раздел](section.md)

#### Примеры
```js
OneNote.run(function (context) {
            
    // Queue a command to insert a section after the current section.
    var section = context.application.getActiveSection().insertSectionAsSibling("After", "New section");
            
    // Queue a command to load the id and name of the new section. 
    // This example loads the new section so it can read its properties later.           
    section.load('id,name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
             
            // Display the properties.       
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

**id**
```js
OneNote.run(function (context) {
        
    // Get the current section.
    var section = context.application.getActiveSection();
            
    // Queue a command to load the section. 
    // For best performance, request specific properties.           
    section.load("id");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
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

**name и notebook**
```js
OneNote.run(function (context) {
        
    // Get the current section.
    var section = context.application.getActiveSection();
            
    // Queue a command to load the section with the specified properties. 
    section.load("name,notebook/name");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("Section name: " + section.name);
            console.log("Parent notebook name: " + section.notebook.name);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**parentSectionGroupOrNull**
```js
OneNote.run(function (context) {
    // Queue a command to add a page to the current section.
    var section = context.application.getActiveSection();
    section.load('clientUrl,notebook');
    var sectionGroup = section.parentSectionGroupOrNull;
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            if(sectionGroup.isNull === false)
            {
                // If a parent section group exists, queue a command to add a section in it!
                sectionGroup.addSection("NewSectionInSectionGroup");
            }
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
    
