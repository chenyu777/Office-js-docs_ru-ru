# Объект Range (API JavaScript для Word)

Представляет непрерывную область в документе.

_Область применения: Word 2016, Word для iPad, Word для Mac_

## Свойства
| Свойство   | Тип|Описание
|:---------------|:--------|:----------|
|style|string|Возвращает или задает стиль диапазона. Это имя предустановленного или пользовательского стиля.|
|text|string|Возвращает текст диапазона. Только для чтения.|

_Ознакомьтесь с [примерами](#property-access-examples) доступа к свойствам._

## Связи
| Связь | Тип|Описание|
|:---------------|:--------|:----------|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Возвращает коллекцию объектов элементов управления содержимым в диапазоне. Только для чтения.|
|font|[Font](font.md)|Возвращает формат текста диапазона. Используйте это свойство, чтобы получать и задавать имея, размер, цвет и другие свойства шрифта. Только для чтения.|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Возвращает коллекцию объектов inlinePicture, включенных в диапазон. Только для чтения.|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Возвращает коллекцию объектов paragraph в диапазоне. Только для чтения.|
|parentContentControl|[ContentControl](contentcontrol.md)|Возвращает элемент управления содержимым, содержащий диапазон. Возвращает значение null, если родительского элемента управления содержимым не существует. Только для чтения.|

## Методы

| Метод   | Возвращаемый тип|Описание|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Очищает содержимое объекта диапазона. Пользователь может отменить операцию для очищенного содержимого.|
|[delete()](#delete)|void|Удаляет диапазон и его содержимое из документа.|
|[getHtml()](#gethtml)|string|Возвращает HTML-представление объекта диапазона.|
|[getOoxml()](#getooxml)|string|Возвращает OOXML-представление объекта диапазона.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Вставляет разрыв в указанном расположении. Разрыв можно вставить только в объекты диапазона, содержащиеся в основном тексте документа, за исключением разрыва строки, который можно вставить в любой объект в тексте. Возможные значения InsertLocation: Replace, Before и After.|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Включает объект диапазона в элемент управления содержимым форматированного текста.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Вставляет документ в диапазон в указанном расположении. Возможные значения InsertLocation: Replace, Start и End.|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Вставляет HTML в диапазон в указанном расположении. Возможные значения InsertLocation: Replace, Start и End.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Вставляет рисунок в содержимое в заданном расположении. Возможные значения insertLocation: Replace, Start, End, Before и After.
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Вставляет OOXML или wordProcessingML в диапазон в указанном расположении. Возможные значения InsertLocation: Replace, Start и End.|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Вставляет абзац в диапазон в указанном расположении. Возможные значения InsertLocation: Before и After.|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Вставляет текст в диапазон в указанном расположении. Возможные значения InsertLocation: Replace, Start и End.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestringssearchoptions)|[SearchResultCollection](searchresultcollection.md)|Выполняет поиск с помощью указанного объекта searchOptions в области объекта диапазона. Результат поиска — это коллекция объектов диапазона.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Выбор диапазона и переход к нему в пользовательском интерфейсе Word. Возможные значения selectionMode: Select, Start и End.|

## Сведения о методе

### clear()
Очищает содержимое объекта диапазона. Пользователь может отменить операцию для очищенного содержимого.

#### Синтаксис
```js
rangeObject.clear();
```

#### Параметры
Нет

#### Возвращаемое значение
void

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to clear the contents of the proxy range object.
    range.clear();
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Cleared the selection (range object)');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```
### delete()
Удаляет диапазон и его содержимое из документа.

#### Синтаксис
```js
rangeObject.delete();
```

#### Параметры
Нет

#### Возвращаемое значение
void

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to delete the range object.
    range.delete();
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Deleted the selection (range object)');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### getHtml()
Возвращает HTML-представление объекта диапазона.

#### Синтаксис
```js
rangeObject.getHtml();
```

#### Параметры
Нет

#### Возвращаемое значение
string

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to get the HTML of the current selection. 
    var html = range.getHtml();
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The HTML read from the document was: ' + html.value);
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### getOoxml()
Возвращает OOXML-представление объекта диапазона.

#### Синтаксис
```js
rangeObject.getOoxml();
```

#### Параметры
Нет

#### Возвращаемое значение
string

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to get the OOXML of the current selection. 
    var ooxml = range.getOoxml();
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The OOXML read from the document was:  ' + ooxml.value);
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)
Вставляет разрыв в указанном расположении. Разрыв можно вставить только в объекты диапазона, содержащиеся в основном тексте документа, за исключением разрыва строки, который можно вставить в любой объект в тексте. Возможные значения InsertLocation: Replace, Before и After.

#### Синтаксис
```js
rangeObject.insertBreak(breakType, insertLocation);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|breakType|BreakType|Обязательный параметр. Тип разрыва, добавляемого в диапазон.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Replace, Before и After.|

#### Возвращаемое значение
void

#### Дополнительные сведения
Невозможно вставить разрыв в верхние и нижние колонтитулы, сноску, концевую сноску, примечание и текстовое поле, за исключением разрыва строки. 

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert a page break after the selected text.
    range.insertBreak('page', 'After');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Inserted a page break after the selected text.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertContentControl()
Включает объект диапазона в элемент управления содержимым форматированного текста.

#### Синтаксис
```js
rangeObject.insertContentControl();
```

#### Параметры
Нет

#### Возвращаемое значение
[ContentControl](contentcontrol.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert a content control around the selected text,
    // and create a proxy content control object. We'll update the properties
    // on the content control.
    var myContentControl = range.insertContentControl();
    myContentControl.tag = "Customer-Address";
    myContentControl.title = "Enter Customer Address Here:";
    myContentControl.style = "Normal";
    myContentControl.insertText("One Microsoft Way, Redmond, WA 98052", 'replace');
    myContentControl.cannotEdit = true;
    myContentControl.appearance = "tags";
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Wrapped a content control around the selected text.');
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
Вставляет документ в диапазон в указанном расположении. Возможные значения InsertLocation: Replace, Start и End.

#### Синтаксис
```js
rangeObject.insertFileFromBase64(base64File, insertLocation);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|base64File|string|Обязательный параметр. Содержимое файла с кодировкой base64, которое необходимо вставить.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Replace, Start и End.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert base64 encoded .docx at the beginning of the range.
    // You'll need to implement getBase64() to make this work.
    range.insertFileFromBase64(getBase64(), Word.InsertLocation.start);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Added base64 encoded text to the beginning of the range.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertHtml(html: string, insertLocation: InsertLocation)
Вставляет HTML в диапазон в указанном расположении. Возможные значения InsertLocation: Replace, Start и End.

#### Синтаксис
```js
rangeObject.insertHtml(html, insertLocation);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|html|string|Обязательный параметр. HTML-код, который необходимо вставить в диапазон.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Replace, Start и End.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert HTML in to the beginning of the range.
    range.insertHtml('<strong>This is text inserted with range.insertHtml()</strong>', Word.InsertLocation.start);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('HTML added to the beginning of the range.');
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
Вставляет рисунок в содержимое в заданном расположении. Возможные значения insertLocation: Replace, Start, End, Before и After.

#### Синтаксис
rangeObject.insertInlinePictureFromBase64(image, insertLocation);

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Обязательный параметр. Вставляемое в диапазон изображение в кодировке base64.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Replace, Start, End, Before и After.|

#### Возвращаемое значение
[InlinePicture](inlinepicture.md)

### insertOoxml(ooxml: string, insertLocation: InsertLocation)
Вставляет OOXML или wordProcessingML в диапазон в указанном расположении. Возможные значения InsertLocation: Replace, Start и End.

#### Синтаксис
```js
rangeObject.insertOoxml(ooxml, insertLocation);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|"ooxml"|string|Обязательный параметр. OOXML или wordProcessingML, которые необходимо вставить в диапазон.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Replace, Start и End.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert OOXML in to the beginning of the range.
    range.insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", Word.InsertLocation.start);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('OOXML added to the beginning of the range.');
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
Рекомендации по работе с OOXML см. в статье [Создание улучшенных надстроек для Word с помощью Office Open XML](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx).

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Вставляет абзац в диапазон в указанном расположении. Возможные значения InsertLocation: Before и After.

#### Синтаксис
```js
rangeObject.insertParagraph(paragraphText, insertLocation);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|paragraphText|string|Обязательный параметр. Текст абзаца, который необходимо вставить.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Before и After.|

#### Возвращаемое значение
[Paragraph](paragraph.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert the paragraph after the range.
    range.insertParagraph('Content of a new paragraph', Word.InsertLocation.after);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Paragraph added to the end of the range.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertText(text: string, insertLocation: InsertLocation)
Вставляет текст в диапазон в указанном расположении. Возможные значения InsertLocation: Replace, Start и End.

#### Синтаксис
```js
rangeObject.insertText(text, insertLocation);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|text|string|Обязательный параметр. Текст, который необходимо вставить.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Replace, Start и End.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert the paragraph at the end of the range.
    range.insertText('New text inserted into the range.', Word.InsertLocation.end);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Text added to the end of the range.');
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
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|param|object|Необязательный параметр. Принимает имена параметров и связей в виде строки с разделителями или массива. Либо укажите объект [loadOption](loadoption.md).|

#### Возвращаемое значение
void

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to load font and style information for the range.
    context.load(range, 'font/size, font/name, font/color, style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Show the results of the load method. Here we show the
        // property values on the range object.
        var results = "  ---Font size: " + range.font.size +
                      "  ---Font name: " + range.font.name +
                      "  ---Font color: " + range.font.color +
                      "  ---Style: " + range.style;
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
Выполняет поиск с помощью указанного объекта searchOptions в области объекта диапазона. Результат поиска — это коллекция объектов диапазона.

#### Синтаксис
```js
rangeObject.search(searchText, searchOptions);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|searchText|string|Обязательный параметр. Текст для поиска.|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|Необязательный параметр. Параметры поиска.|

#### Возвращаемое значение
[SearchResultCollection](searchresultcollection.md)


### select(selectionMode: SelectionMode)
Выбор диапазона и переход к нему в пользовательском интерфейсе Word. Возможные значения selectionMode: Select, Start и End.

#### Синтаксис
```js
rangeObject.select(selectionMode);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|Необязательный параметр. Возможные режимы выбора: Select, Start и End. Значение по умолчанию — Select.|

#### Возвращаемое значение
void

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to get the current selection and then 
    // create a proxy range object with the results.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert HTML in to the beginning of the range.
    range.insertHtml('<strong>This is text inserted with range.insertHtml()</strong>', Word.InsertLocation.start);
    
    // Queue a command to select the HTML that was inserted.
    range.select();
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Selected the range.');
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

Используйте [набор требований](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx) в проверках среды выполнения, чтобы обеспечить поддержку ведущей версии Word для своего приложения. Дополнительные сведения о требованиях ведущих приложений и серверов Office см. в статье [Требования для запуска надстроек Office](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx). 
