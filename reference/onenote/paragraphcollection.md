# Объект ParagraphCollection (API JavaScript для OneNote)

_Относится к: OneNote Online_  


Представляет коллекцию объектов Paragraph.

## Свойства

| Свойство     | Тип   |Описание|Отзывы|
|:---------------|:--------|:----------|:-------|
|count|int|Возвращает количество объектов Paragraph в объекте Page. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraphCollection-count)|
|items|[Paragraph[]](paragraph.md)|Коллекция объектов paragraph. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraphCollection-items)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[getItem(index: числовой тип или string)](#getitemindex-number-or-string)|[Paragraph](paragraph.md)|Получает объект Paragraph по его идентификатору или индексу в коллекции. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraphCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[Paragraph](paragraph.md)|Получает абзац по позиции в коллекции.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraphCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraphCollection-load)|

## Сведения о методе


### getItem(index: числовой тип или string)
Получает объект Paragraph по его идентификатору или индексу в коллекции. Только для чтения.

#### Синтаксис
```js
paragraphCollectionObject.getItem(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number или string|Идентификатор объекта Paragraph или расположение индекса объекта Paragraph в коллекции.|

#### Возвращаемое значение
[Paragraph](paragraph.md)

### getItemAt(index: number)
Получает объект Paragraph по его позиции в коллекции.

#### Синтаксис
```js
paragraphCollectionObject.getItemAt(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number|Значение индекса получаемого объекта. Используется нулевой индекс.|

#### Возвращаемое значение
[Paragraph](paragraph.md)

#### Примеры
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page, and then get its Outline's first paragraph.
    var pageContent = pageContents.getItemAt(0);
    var paragraphs = pageContent.outline.paragraphs;

    var firstParagraph = paragraphs.getItemAt(0);

    // Queue a command to load the type and richText.text property of this paragraph.
    firstParagraph.load("id,type");


    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            // Write text from paragraph to console
            console.log("First Paragraph found with id : " + firstParagraph.id + " and type " + firstParagraph.type);
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

**items**
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page, and then get its Outline's first paragraph.
    var pageContent = pageContents.getItem(0);
    var paragraphs = pageContent.outline.paragraphs;
    
    // Queue a command to load the id and type of each paragraph.
    paragraphs.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            var firstParagraph = paragraphs.items[0];
            // Write text from first paragraph to console
            console.log("First Paragraph found with id : " + firstParagraph.id + " and type " + firstParagraph.type);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**просмотрите richText**
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page, and then get its outline's paragraphs.
    var outlinePageContents = [];
    var paragraphs = [];
    var richTextParagraphs = [];
    // Queue a command to load the id and type of each page content in the outline.
    pageContents.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            // Load all page contents of type Outline
            $.each(pageContents.items, function(index, pageContent) {
                if(pageContent.type == 'Outline')
                {
                    pageContent.load('outline,outline/paragraphs,outline/paragraphs/type');
                    outlinePageContents.push(pageContent);
                }
            });
            return context.sync();
        })
        .then(function () {
            // Load all rich text paragraphs across outlines
            $.each(outlinePageContents, function(index, outlinePageContent) {
                var outline = outlinePageContent.outline;
                paragraphs = paragraphs.concat(outline.paragraphs.items);
            });
            $.each(paragraphs, function(index, paragraph) {
                if(paragraph.type == 'RichText')
                {
                    richTextParagraphs.push(paragraph);
                    paragraph.load("id,richText/text");
                }
            });
            return context.sync();
        })
        .then(function () {
            // Display all rich text paragraphs to the console
            $.each(richTextParagraphs, function(index, richTextParagraph) {
                var richText = richTextParagraph.richText;
                console.log("Paragraph found with richtext content : " + richText.text + " and richtext id : " + richText.id);
            });
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

