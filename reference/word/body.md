# Объект Body (API JavaScript для Word)

Представляет содержимое документа или раздела.

_Область применения: Word 2016, Word для iPad, Word для Mac._

## Свойства
| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|style|string|Возвращает или задает стиль для содержимого. Это имя предустановленного или пользовательского стиля.|
|text|string|Возвращает текст содержимого. Для вставки текста используется метод insertText. Только для чтения.|

_Ознакомьтесь с [примерами](#примерами) доступа к свойствам._

## Связи
| Связь | Тип   |Описание|
|:---------------|:--------|:----------|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Возвращает коллекцию из объектов управления форматированным текстом, находящихся в содержимом документа или раздела. Только для чтения.|
|font|[Font](font.md)|Возвращает формат текста, указанный для содержимого документа или раздела. Используйте эту связь, чтобы получить и задать имя, размер, цвет и другие свойства шрифта. Только для чтения.|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Возвращает коллекцию объектов inlinePicture, находящихся в содержимом документа или раздела. Коллекция не содержит плавающие изображения. Только для чтения.|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Возвращает коллекцию объектов абзаца, находящихся в содержимом документа или раздела. Только для чтения.|
|parentContentControl|[ContentControl](contentcontrol.md)|Возвращает элемент управления содержимым, включающий объект body. Возвращает значение null, если родительского элемента управления содержимым не существует. Только для чтения.|

## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Очищает объект содержимого. Пользователь может отменить очистку содержимого.|
|[getHtml()](#gethtml)|string|Возвращает HTML-представление объекта body.|
|[getOoxml()](#getooxml)|string|Возвращает OOXML-представление (Office Open XML) объекта body.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Вставляет разрыв в заданном расположении. Разрыв можно вставить только в основной текст документа. Но при этом разрыв строки вставляется в любой объект содержимого. Возможные значения insertLocation: Start и End.|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Включает объект body в элемент управления форматированным текстом.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Вставляет документ в содержимое документа или раздела в заданном расположении. Возможные значения insertLocation: Replace, Start и End.|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Вставляет HTML-код в заданном расположении. Возможные значения insertLocation: Replace, Start и End.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Вставляет изображение в содержимое документа или раздела в заданном расположении. Возможные значения insertLocation: Start и End. |
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Вставляет OOXML-код в заданном расположении.  Возможные значения insertLocation: Replace, Start и End.|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Вставляет абзац в указанном расположении. Возможные значения insertLocation: Start и End.|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Вставляет текст в содержимое документа или раздела в заданном расположении. Возможные значения insertLocation: Replace, Start и End.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[SearchResultCollection](searchresultcollection.md)|Выполняет поиск с помощью указанного параметра searchOptions в области объекта body. Результат поиска — коллекция объектов range.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Выбирает содержимое и переходит к нему в пользовательском интерфейсе Word. Возможные значения selectionMode: Select, Start и End.|

## Сведения о методе

### clear()
Очищает объект содержимого. Пользователь может отменить операцию очищения для содержимого.

#### Синтаксис
```js
bodyObject.clear();
```

#### Параметры
Нет

#### Возвращаемое значение
void

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to clear the contents of the body.
    body.clear();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Cleared the body contents.');
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

В примере надстройки [Silly stories](https://aka.ms/sillystorywordaddin) показано, как можно использовать метод **clear** для очистки содержимого документа.

### getHtml()
Возвращает HTML-представление объекта содержимого

#### Синтаксис
```js
bodyObject.getHtml();
```

#### Параметры
Нет

#### Возвращаемое значение
string

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to get the HTML contents of the body.
    var bodyHTML = body.getHtml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body HTML contents: " + bodyHTML.value);
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### getOoxml()
Возвращает OOXML-представление (Office Open XML) объекта содержимого.

#### Синтаксис
```js
bodyObject.getOoxml();
```

#### Параметры
Нет

#### Возвращаемое значение
string

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to get the OOXML contents of the body.
    var bodyOOXML = body.getOoxml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body OOXML contents: " + bodyOOXML.value);
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)
Вставляет разрыв в заданном расположении. Разрыв можно вставить только в основной текст документа. Но при этом разрыв строки вставляется в любой объект содержимого. Возможные значения insertLocation: Start и End.

#### Синтаксис
```js
bodyObject.insertBreak(breakType, insertLocation);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|breakType|BreakType|Обязательный параметр. Тип разрыва, который необходимо добавить в содержимое.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Start и End.|

#### Возвращаемое значение
void

#### Дополнительные сведения
За исключением разрывов строк, вы не можете вставлять разрывы в верхние и нижние колонтитулы, сноски, концевые сноски, примечания и текстовые поля.

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (ctx) {

    // Create a proxy object for the document body.
    var body = ctx.document.body;

    // Queue a commmand to insert a page break at the start of the document body.
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        console.log('Added a page break at the start of the document body.');
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
### insertContentControl()
Включает объект содержимого в элемент управления форматированным текстом.

#### Синтаксис
```js
bodyObject.insertContentControl();
```

#### Параметры
Нет

#### Возвращаемое значение
[ContentControl](contentcontrol.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to wrap the body in a content control.
    body.insertContentControl();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Wrapped the body in a content control.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```
### insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Вставляет документ в содержимое в заданном расположении. Возможные значения insertLocation: Replace, Start и End.

#### Синтаксис
```js
bodyObject.insertFileFromBase64(base64File, insertLocation);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|base64File|string|Обязательный параметр. Содержимое файла в кодировке base64 для вставки.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Replace, Start и End.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert base64 encoded .docx at the beginning of the content body.
    // You will need to implement getBase64() to pass in a string of a base64 encoded docx file.
    body.insertFileFromBase64(getBase64(), Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Added base64 encoded text to the beginning of the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

В примере надстройки [Silly stories](https://aka.ms/sillystorywordaddin) показано, как можно использовать метод **insertFileFromBase64**, чтобы вставить DOCX-файлы из службы.

### insertHtml(html: string, insertLocation: InsertLocation)
Вставляет HTML-код в заданном расположении. Возможные значения insertLocation: Replace, Start и End.

#### Синтаксис
```js
bodyObject.insertHtml(html, insertLocation);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|html|string|Обязательный параметр. HTML-код, который необходимо вставить в документ.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Replace, Start и End.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert HTML in to the beginning of the body.
    body.insertHtml('<strong>This is text inserted with body.insertHtml()</strong>', Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('HTML added to the beginning of the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Вставляет рисунок в содержимое в заданном расположении. Возможные значения insertLocation: Start и End.

#### Синтаксис
bodyObject.insertInlinePictureFromBase64(image, insertLocation);

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Обязательный параметр. Вставляемое в основной текст изображение в кодировке base64.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Start и End.|

#### Возвращаемое значение
[InlinePicture](inlinepicture.md)

### insertOoxml(ooxml: string, insertLocation: InsertLocation)
Вставляет OOXML-код в заданном расположении. Возможные значения insertLocation: Replace, Start и End.

#### Синтаксис
```js
bodyObject.insertOoxml(ooxml, insertLocation);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|"ooxml"|string|Обязательный параметр. OOXML или wordProcessingML, которые необходимо вставить.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Replace, Start и End.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert OOXML in to the beginning of the body.
    body.insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('OOXML added to the beginning of the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### Дополнительные сведения
Рекомендации по работе с OOXML см. в статье [Создание улучшенных надстроек для Word с использованием Office Open XML](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx). В примере [Word-Add-in-DocumentAssembly][body.insertOoxml] показано, как можно использовать этот API для сборки документа.

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Вставляет абзац в заданном расположении. Возможные значения insertLocation: Start и End.

#### Синтаксис
```js
bodyObject.insertParagraph(paragraphText, insertLocation);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|paragraphText|string|Обязательный параметр. Текст абзаца, который необходимо вставить.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Start и End.|

#### Возвращаемое значение
[Paragraph](paragraph.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert the paragraph at the end of the document body.
    body.insertParagraph('Content of a new paragraph', Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Paragraph added at the end of the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### Дополнительные сведения
В примере [Word-Add-in-DocumentAssembly][body.insertParagraph] показано, как можно использовать метод insertParagraph для сборки документа.

### insertText(text: string, insertLocation: InsertLocation)
Вставляет текст в содержимое в заданном расположении. Возможные значения insertLocation: Replace, Start и End.

#### Синтаксис
```js
bodyObject.insertText(text, insertLocation);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|text|string|Обязательный параметр. Текст, который необходимо вставить.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Replace, Start и End.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert text in to the beginning of the body.
    body.insertText('This is text inserted with body.insertText()', Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Text added to the beginning of the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
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

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to load font and style information for the document body.
    context.load(body, 'font/size, font/name, font/color, style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Show the results of the load method. Here we show the
        // property values on the body object.
        var results = 'Font size: ' + body.font.size +
                      '; Font name: ' + body.font.name +
                      '; Font color: ' + body.font.color +
                      '; Body style: ' + body.style;

        console.log(results);
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```
### search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)
Выполняет поиск с помощью указанных параметров поиска в области объекта содержимого. Результат поиска — коллекция объектов диапазона.

#### Синтаксис
```js
bodyObject.search(searchText, searchOptions);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|searchText|string|Обязательный параметр. Текст для поиска.|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|Необязательный параметр. Параметры поиска.|

#### Возвращаемое значение
[SearchResultCollection](searchresultcollection.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to search the document.
    var searchResults = context.document.body.search('video', {matchCase: false});

    // Queue a commmand to load the results.
    context.load(searchResults, 'text, font');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        var results = 'Found count: ' + searchResults.items.length +
                      '; we highlighted the results.';

        // Queue a command to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
          searchResults.items[i].font.color = '#FF0000'    // Change color to Red
          searchResults.items[i].font.highlightColor = '#FFFF00';
          searchResults.items[i].font.bold = true;
        }

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log(results);
        });
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### Дополнительные сведения
[Word-Add-in-DocumentAssembly][body.search] — еще один пример поиска в документе.

### select(selectionMode: SelectionMode)
Выбирает содержимое и переходит к нему в пользовательском интерфейсе Word. Возможные значения selectionMode: Select, Start и End.

#### Синтаксис
```js
bodyObject.select(selectionMode);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|Необязательный параметр. Возможные режимы выбора: Select, Start и End. Значение по умолчанию — Select.|

#### Возвращаемое значение
void

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to select the document body. The Word UI will
    // move to the selected document body.
    body.select();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Selected the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## Примеры доступа к свойствам

### Получение свойства текста в объекте содержимого
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to load the text in document body.
    context.load(body, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
### Получение свойств стиля и размера шрифта, его названия и цвета в объекте содержимого

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to load font and style information for the document body.
    context.load(body, 'font/size, font/name, font/color, style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Show the results of the load method. Here we show the
        // property values on the body object.
        var results = 'Font size: ' + body.font.size +
                      '; Font name: ' + body.font.name +
                      '; Font color: ' + body.font.color +
                      '; Body style: ' + body.style;

        console.log(results);
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## Сведения о поддержке

Используйте [набор требований](../office-add-in-requirement-sets.md) в проверках среды выполнения, чтобы обеспечить поддержку ведущей версии Word для своего приложения. Дополнительные сведения о требованиях ведущих приложений и серверов Office см. в статье [Требования для запуска надстроек Office](../../docs/overview/requirements-for-running-office-add-ins.md).


[body.insertOoxml]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L127 "вставка OOXML"
[body.insertParagraph]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L153 "вставка абзаца"
[body.search]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L261 "поиск в содержимом документа или раздела"
