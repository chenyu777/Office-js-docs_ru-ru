# Объект Image (API JavaScript для OneNote)

_Относится к: OneNote Online_  


Представляет изображение. Объект Image может быть прямым потомком объекта PageContent или объекта Paragraph.

## Свойства

| Свойство     | Тип   |Описание|Отзывы|
|:---------------|:--------|:----------|:-------|
|description|string|Получает или задает описание объекта Image.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-description)|
|height|double|Получает или задает высоту объекта Image.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-height)|
|hyperlink|string|Получает или задает гиперссылку для объекта Image.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-hyperlink)|
|id|string|Получает идентификатор объекта Image. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-id)|
|width|double|Получает или задает ширину объекта Image.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-width)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
| Связь | Тип   |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|ocrData|[ImageOcrData](imageocrdata.md)|Получает данные, собранные с помощью распознавания текста (OCR) для этого объекта Image, например текст и язык. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-ocrData)|
|pageContent|[PageContent](pagecontent.md)|Получает объект PageContent, содержащий объект Image. Выдает исключение, если объект Image не является непосредственным дочерним элементом для объекта PageContent. Этот объект определяет положение объекта Image на странице. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-pageContent)|
|paragraph|[Paragraph](paragraph.md)|Получает объект Paragraph, содержащий объект Image. Выдает исключение, если объект Image не является непосредственным дочерним элементом для объекта Paragraph. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-paragraph)|

## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[getBase64Image()](#getbase64image)|строка|Получает двоичное представление объекта Image в кодировке Base64.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-getBase64Image)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-load)|

## Сведения о методе


### getBase64Image()
Получает двоичное представление объекта Image в кодировке Base64.

#### Синтаксис
```js
imageObject.getBase64Image();
```

#### Параметры
Нет

#### Возвращаемое значение
string

#### Примеры
```js

var image = null;
var imageString;

OneNote.run(function(ctx){
    // Get the current outline.         
    var outline = ctx.application.getActiveOutline();
    
    // Queue a command to load paragraphs and their types. 
    outline.load("paragraphs/type")
    return ctx.sync().
        then(function(){
            for (var i=0; i < outline.paragraphs.items.length; i++)
            {
                var paragraph = outline.paragraphs.items[i];
                if (paragraph.type == "Image")
                {
                    image = paragraph.image;
                }
            }
        })
        .then(function(){
            if (image != null)
            {
                imageString = image.getBase64Image();
                return ctx.sync();
            }
        })
        .then(function(){
            console.log(imageString);
        });
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
**id, width, height, description и hyperlink**
```js
OneNote.run(function(ctx){
    // Get the current outline.         
    var outline = ctx.application.getActiveOutline();
    var image = null;
    
    // Queue a command to load paragraphs and their types. 
    outline.load("paragraphs/type")
    return ctx.sync().
        then(function(){
            for (var i=0; i < outline.paragraphs.items.length; i++)
            {
                var paragraph = outline.paragraphs.items[i];
                if (paragraph.type == "Image")
                {
                    image = paragraph.image;
                }
            }
        })
        .then(function(){
            if (image != null)
            {
                // load every properties and relationships
                ctx.load(image);
                return ctx.sync();
            }
        })
        .then(function(){
            if (image != null)
            {                   
                console.log("image " + image.id + " width is " + image.width + " height is " + image.height);
                console.log("description: " + image.description);                   
                console.log("hyperlink: " + image.hyperlink);
            }
        });
});
```

**ocrData**
```js
var image = null;

OneNote.run(function(ctx){
    // Get the current outline.
    var outline = ctx.application.getActiveOutline();

    // Queue a command to load paragraphs and their types.
    outline.load("paragraphs")
    return ctx.sync().
        then(function(){
            for (var i=0; i < outline.paragraphs.items.length; i++)
            {
                var paragraph = outline.paragraphs.items[i];
                if (paragraph.type == "Image")
                {
                    image = paragraph.image;
                }
            }
            if (image != null)
            {
               image.load("ocrData");
            }
            return ctx.sync();
        })
        .then(function(){
            console.log(image.ocrData);
        });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**paragraph**
```js
OneNote.run(function(ctx){
    // Get the current outline.         
    var outline = ctx.application.getActiveOutline();
    var searchedParagraph = null;
    
    // Queue a command to load paragraphs and their types. 
    outline.load("paragraphs/type")
    return ctx.sync().
        then(function() {
            for (var i=0; i < outline.paragraphs.items.length; i++)
            {
                var paragraph = outline.paragraphs.items[i];
                if (paragraph.type == "Image")
                {
                    searchedParagraph = paragraph;
                    break;
                }
            }
        })
        .then(function() {
            if (searchedParagraph != null)
            {
                // load every properties and relationships
                searchedParagraph.image.load('paragraph');
                return ctx.sync();
            }
        })
        .then(function() {
            if (searchedParagraph != null)
            {                   
                if (searchedParagraph.id != searchedParagraph.image.paragraph.id)
                {
                    console.log("id must match");
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

