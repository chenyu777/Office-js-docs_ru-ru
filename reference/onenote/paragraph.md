# Объект Paragraph (API JavaScript для OneNote)

_Относится к: OneNote Online_  


Контейнер для контента, отображаемого на странице. Объект Paragraph может содержать контент с любым типом ParagraphType.

## Свойства

| Свойство     | Тип   |Описание|Отзывы|
|:---------------|:--------|:----------|:-------|
|id|string|Получает идентификатор объекта Paragraph. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-id)|
|type|string|Получает тип объекта Paragraph. Только для чтения. Возможные значения: RichText, Image, Table, Other.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-type)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
| Связь | Тип   |Описание| Отзывы|
|:---------------|:--------|:----------|:-------|
|image|[Image](image.md)|Получает объект Image в объекте Paragraph. Вызывает исключение, если ParagraphType не является Image. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-image)|
|inkWords|[InkWordCollection](inkwordcollection.md)|Возвращает коллекцию рукописных фрагментов в объекте Paragraph. Вызывает исключение, если ParagraphType не является Ink. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-inkWords)|
|outline|[Outline](outline.md)|Получает объект Outline, содержащий объект Paragraph. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-outline)|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Коллекция абзацев под этим абзацем. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-paragraphs)|
|parentParagraph|[Paragraph](paragraph.md)|Получает родительский объект Paragraph. Возвращает исключение, если родительского объекта абзаца не существует. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentParagraph)|
|parentParagraphOrNull|[Paragraph](paragraph.md)|Получает родительский объект Paragraph. Возвращает значение null, если родительского объекта абзаца не существует. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentParagraphOrNull)|
|parentTableCell|[TableCell](tablecell.md)|Получает объект TableCell, содержащий объект Paragraph, если он существует. Если родительский объект не TableCell, возвращает ItemNotFound. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentTableCell)|
|parentTableCellOrNull|[TableCell](tablecell.md)|Получает объект TableCell, содержащий объект Paragraph, если он существует. Если родительский объект не TableCell, возвращает значение null. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentTableCellOrNull)|
|richText|[RichText](richtext.md)|Получает объект RichText в объекте Paragraph. Вызывает исключение, если ParagraphType не является RichText. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-richText)|
|table|[Table](table.md)|Получает объект Table в объекте Paragraph. Вызывает исключение, если ParagraphType не является Table. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-table)|

## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[delete()](#delete)|void|Удаляет объект Paragraph.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-delete)|
|[insertHtmlAsSibling(insertLocation: string, html: string)](#inserthtmlassiblinginsertlocation-string-html-string)|void|Вставляет указанный HTML-контент|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertHtmlAsSibling)|
|[insertImageAsSibling(insertLocation: string, base64EncodedImage: string, width: double, height: double)](#insertimageassiblinginsertlocation-string-base64encodedimage-string-width-double-height-double)|[Image](image.md)|Вставляет изображение в указанное место.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertImageAsSibling)|
|[insertRichTextAsSibling(insertLocation: string, paragraphText: string)](#insertrichtextassiblinginsertlocation-string-paragraphtext-string)|[RichText](richtext.md)|Вставляет текст абзаца в указанное место.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertRichTextAsSibling)|
|[insertTableAsSibling(insertLocation: string, rowCount: number, columnCount: number, values: string[][])](#inserttableassiblinginsertlocation-string-rowcount-number-columncount-number-values-string)|[Table](table.md)|Добавляет таблицу с указанным количеством строк и столбцов перед текущим абзацем или после него.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertTableAsSibling)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-load)|

## Сведения о методе


### delete()
Удаляет объект Paragraph.

#### Синтаксис
```js
paragraphObject.delete();
```

#### Параметры
Нет

#### Возвращаемое значение
void

#### Примеры
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
    
    var paragraphs = pageContent.outline.paragraphs;
    
    var firstParagraph = paragraphs.getItemAt(0);
    
    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Queue a command to delete the first paragraph                 
            firstParagraph.delete();
            
            // Run the command to delete it
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


### insertHtmlAsSibling(insertLocation: string, html: string)
Вставляет указанный HTML-контент

#### Синтаксис
```js
paragraphObject.insertHtmlAsSibling(insertLocation, html);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|insertLocation|string|Расположение нового содержимого относительно текущего объекта Paragraph.  Возможные значения: Before, After|
|html|string|Строка HTML, описывающая визуальное представление содержимого. API JavaScript для надстроек OneNote см. в разделе [Поддерживаемые элементы HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html).|

#### Возвращаемое значение
void

#### Примеры
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
    var paragraphs = pageContent.outline.paragraphs;
    var firstParagraph = paragraphs.getItemAt(0);

    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Queue commands to insert before and after the first paragraph
            firstParagraph.insertHtmlAsSibling("Before", "<p>ContentBeforeFirstParagraph</p>");
            firstParagraph.insertHtmlAsSibling("After", "<p>ContentAfterFirstParagraph</p>");
            
            // Run the command to run inserts
            return context.sync();
        });
))
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### insertImageAsSibling(insertLocation: string, base64EncodedImage: string, width: double, height: double)
Вставляет изображение в указанное место.

#### Синтаксис
```js
paragraphObject.insertImageAsSibling(insertLocation, base64EncodedImage, width, height);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|insertLocation|string|Расположение таблицы относительно текущего объекта Paragraph.  Возможные значения: Before, After|
|base64EncodedImage|string|Строка HTML, которую необходимо добавить.|
|width|double|Необязательный. Ширина в точках. Значение по умолчанию — null, ширина изображения имеет приоритет.|
|height|double|Необязательный. Высота в точках. Значение по умолчанию — null, высота изображения имеет приоритет.|

#### Возвращаемое значение
[Image](image.md)

#### Примеры
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
    var paragraphs = pageContent.outline.paragraphs;
    var firstParagraph = paragraphs.getItemAt(0);

    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Queue commands to insert before and after the first paragraph
            firstParagraph.insertImageAsSibling("Before", "R0lGODlhDwAPAKECAAAAzMzM/////wAAACwAAAAADwAPAAACIISPeQHsrZ5ModrLlN48CXF8m2iQ3YmmKqVlRtW4MLwWACH+H09wdGltaXplZCBieSBVbGVhZCBTbWFydFNhdmVyIQAAOw==");
            firstParagraph.insertImageAsSibling("After", "R0lGODlhDwAPAKECAAAAzMzM/////wAAACwAAAAADwAPAAACIISPeQHsrZ5ModrLlN48CXF8m2iQ3YmmKqVlRtW4MLwWACH+H09wdGltaXplZCBieSBVbGVhZCBTbWFydFNhdmVyIQAAOw==");
            
            // Run the command to insert images
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


### insertRichTextAsSibling(insertLocation: string, paragraphText: string)
Вставляет текст абзаца в указанное место.

#### Синтаксис
```js
paragraphObject.insertRichTextAsSibling(insertLocation, paragraphText);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|insertLocation|string|Расположение таблицы относительно текущего объекта Paragraph.  Возможные значения: Before, After|
|paragraphText|string|Строка HTML, которую необходимо добавить.|

#### Возвращаемое значение
[RichText](richtext.md)

#### Примеры
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
    var paragraphs = pageContent.outline.paragraphs;
    var firstParagraph = paragraphs.getItemAt(0);

    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Queue commands to insert before and after the first paragraph
            firstParagraph.insertRichTextAsSibling("Before", "Text Appears Before Paragraph");
            firstParagraph.insertRichTextAsSibling("After", "Text Appears After Paragraph");
            
            // Run the command to insert text contents
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


### insertTableAsSibling(insertLocation: string, rowCount: number, columnCount: number, values: string[][])
Добавляет таблицу с указанным количеством строк и столбцов перед текущим абзацем или после него.

#### Синтаксис
```js
paragraphObject.insertTableAsSibling(insertLocation, rowCount, columnCount, values);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|insertLocation|string|Расположение таблицы относительно текущего объекта Paragraph.  Возможные значения: Before, After|
|rowCount|number|Количество строк в таблице.|
|columnCount|number|Количество столбцов в таблице.|
|values|string[][]|Необязательный. Необязательный двухмерный массив. Ячейки заполняются, если в массиве указаны соответствующие строки.|

#### Возвращаемое значение
[Table](table.md)

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

**id и type**
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;
    
    // Queue a command to load the outline property of each pageContent.
    pageContents.load("outline");
        
    // Get the first PageContent on the page, and then get its Outline.
    var pageContent = pageContents._GetItem(0);
    var paragraphs = pageContent.outline.paragraphs;
            
    // Queue a command to load the id and type of each paragraph.
    paragraphs.load("id,type");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Write the text.                  
            $.each(paragraphs.items, function(index, paragraph) {
                console.log("Paragraph type: " + paragraph.type);
                console.log("Paragraph ID: " + paragraph.id);
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

**paragraphs**
```js
OneNote.run(function(context) {
    var app = context.application;
    
    // Gets the active outline
    var outline = app.getActiveOutline();
    
    // load nested paragraphs and their types.
    outline.load("paragraphs/type");
    
    return context.sync().then(function () {
        var paragraphs = outline.paragraphs.items;
        
        var promise;
        // for each nested paragraphs, load tables only
        for (var i = 0; i < paragraphs.length; i++) {
            var paragraph = paragraphs[i];
            if (paragraph.type == "Table") {
                paragraph.load("table/id");
                promise =  context.sync().then(function() {
                    console.log(paragraph.table.id);
                });
            }
        }
        return promise;
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

