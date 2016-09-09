# Объект Outline (API JavaScript для OneNote)

_Относится к: OneNote Online_  


Представляет контейнер для объектов Paragraph.

## Свойства

| Свойство     | Тип   |Описание|Отзывы|
|:---------------|:--------|:----------|:-------|
|id|string|Получает идентификатор объекта Outline. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-id)|

## Связи
| Связь | Тип   |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|pageContent|[PageContent](pagecontent.md)|Получает объект PageContent, содержащий объект Outline. Этот объект определяет положение объекта Outline на странице. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-pageContent)|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Получает коллекцию объектов Paragraph в объекте Outline. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-paragraphs)|

## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[appendHtml(html: string)](#appendhtmlhtml-string)|void|Добавляет указанный HTML в нижнюю часть объекта Outline.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendHtml)|
|[appendImage(base64EncodedImage: string, width: double, height: double)](#appendimagebase64encodedimage-string-width-double-height-double)|[Image](image.md)|Добавляет указанное изображение в нижнюю часть объекта Outline.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendImage)|
|[appendRichText(paragraphText: string)](#appendrichtextparagraphtext-string)|[RichText](richtext.md)|Добавляет указанный текст в нижнюю часть объекта Outline.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendRichText)|
|[appendTable(rowCount: number, columnCount: number, values: string[][])](#appendtablerowcount-number-columncount-number-values-string)|[Table](table.md)|Добавляет таблицу с указанным количеством строк и столбцов в нижнюю часть объекта Outline.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendTable)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-load)|

## Сведения о методе


### appendHtml(html: string)
Добавляет указанный HTML в нижнюю часть объекта Outline.

#### Синтаксис
```js
outlineObject.appendHtml(html);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|html|string|Строка HTML, которую необходимо добавить. Сведения об API JavaScript для надстроек OneNote см. в разделе [Поддерживаемые элементы HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html).|

#### Возвращаемое значение
void

#### Примеры
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Get pageContents of the activePage. 
    var pageContents = activePage.contents;

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            if (pageContents.items.length != 0 && pageContents.items[0].type == "Outline")
            {
                // First item is an outline.
                outline = pageContents.items[0].outline;

                // Queue a command to append a paragraph to the outline.
                outline.appendHtml("<p>new paragraph</p>");

                // Run the queued commands.
                return context.sync();
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


### appendImage(base64EncodedImage: string, width: double, height: double)
Добавляет указанное изображение в нижнюю часть объекта Outline.

#### Синтаксис
```js
outlineObject.appendImage(base64EncodedImage, width, height);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Строка HTML, которую необходимо добавить.|
|width|double|Необязательный. Ширина в точках. Значение по умолчанию — null, ширина изображения имеет приоритет.|
|height|double|Необязательный. Высота в точках. Значение по умолчанию — null, высота изображения имеет приоритет.|

#### Возвращаемое значение
[Image](image.md)

### appendRichText(paragraphText: string)
Добавляет указанный текст в нижнюю часть объекта Outline.

#### Синтаксис
```js
outlineObject.appendRichText(paragraphText);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|paragraphText|string|Строка HTML, которую необходимо добавить.|

#### Возвращаемое значение
[RichText](richtext.md)

### appendTable(rowCount: number, columnCount: number, values: string[][])
Добавляет таблицу с указанным количеством строк и столбцов в нижнюю часть объекта Outline.

#### Синтаксис
```js
outlineObject.appendTable(rowCount, columnCount, values);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|rowCount|number|Обязательный. Количество строк в таблице.|
|columnCount|number|Обязательный. Количество столбцов в таблице.|
|values|string[][]|Необязательный. Необязательный двухмерный массив. Ячейки заполняются, если в массиве указаны соответствующие строки.|

#### Возвращаемое значение
[Таблица](table.md)

#### Примеры
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Get pageContents of the activePage. 
    var pageContents = activePage.contents;

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            if (pageContents.items.length != 0 && pageContents.items[0].type == "Outline") {
                // First item is an outline.
                var outline = pageContents.items[0].outline;

                // Queue a command to append a paragraph to the outline.
                outline.appendTable(2, 2, [[1, 2],[3, 4]]);

                // Run the queued commands.
                return context.sync();
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
