# Объект Paragraph (API JavaScript для Word)

Представляет один абзац в выделении, диапазоне, элементе управления содержимым или тексте документа.

_Область применения: Word 2016, Word для iPad, Word для Mac_

## Свойства
| Свойство   | Тип|Описание
|:---------------|:--------|:----------|
|outlineLevel|int|Возвращает или задает уровень структуры абзаца.|
|style|string|Возвращает или задает стиль абзаца. Это имя предустановленного или пользовательского стиля. В примере [Word-Add-in-DocumentAssembly][paragraph.style] показано, как задать стиль абзаца.|
|text|string|Возвращает текст абзаца. Только для чтения.|

_Ознакомьтесь с [примерами](#property-access-examples) доступа к свойствам._

## Связи
| Связь | Тип|Описание|
|:---------------|:--------|:----------|
|alignment|**Alignment**|Возвращает или задает выравнивание для абзаца. Возможные значения: left, centered, right или justified.|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Возвращает коллекцию объектов элементов управления содержимым в абзаце. Только для чтения.|
|firstLineIndent|**float**|Возвращает или задает значение отступа первой строки или выступа в пунктах. Для установки отступа первой строки укажите положительное значение и используйте отрицательное значение, чтобы задать выступ.|
|font|[Font](font.md)|Возвращает формат текста абзаца. Используйте это свойство для получения и задания имени, размера, цвета и других свойств шрифта. Только для чтения.|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Возвращает коллекцию объектов inlinePicture в абзаце. Коллекция не содержит плавающие рисунки. Только для чтения.|
|leftIndent|**float**|Возвращает или задает значение отступа слева для абзаца (в пунктах).|
|lineSpacing|**float**|Возвращает или задает междустрочный интервал для указанного абзаца (в пунктах). В пользовательском интерфейсе Word это значение делится на 12.|
|lineUnitAfter|**float**|Возвращает или устанавливает междустрочный интервал после абзаца (в линиях сетки).|
|lineUnitBefore|**float**|Возвращает или устанавливает междустрочный интервал до абзаца (в линиях сетки).|
|parentContentControl|[ContentControl](contentcontrol.md)|Возвращает элемент управления содержимым, содержащий абзац. Возвращает значение null, если родительского элемента управления содержимым не существует. Только для чтения.|
|rightIndent|**float**|Возвращает или задает значение отступа справа для абзаца (в пунктах).|
|spaceAfter|**float**|Возвращает или задает междустрочный интервал после абзаца (в пунктах).|
|spaceBefore|**float**|Возвращает или задает междустрочный интервал до абзаца (в пунктах).|

## Методы

| Метод   | Возвращаемый тип|Описание|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Очищает содержимое объекта абзаца. Пользователь может отменить операцию для очищенного содержимого.|
|[delete()](#delete)|void|Удаляет абзац и его содержимое из документа.|
|[getHtml()](#gethtml)|string|Возвращает HTML-представление объекта абзаца.|
|[getOoxml()](#getooxml)|string|Возвращает OOXML-представление объекта абзаца.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Вставляет разрыв в указанном месте. Разрыв можно вставить только в абзацы, которые содержатся в основном тексте документа, за исключением разрыва строки, который можно вставить в любой объект текста. Возможные значения InsertLocation: After и Before.|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Включает объект абзаца в элемент управления содержимым форматированного текста.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Вставляет документ в текущий абзац в указанном расположении. Возможные значения InsertLocation: Start или End.|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Вставляет HTML в абзац в указанном расположении. Возможные значения InsertLocation: Replace, Start и End.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Вставляет рисунок в абзац в указанном расположении. Возможные значения InsertLocation: Before, After, Start и End.|
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Вставляет OOXML или wordProcessingML в абзац в указанном расположении. Возможные значения InsertLocation: Replace, Start и End.|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Вставляет абзац в указанном расположении. Возможные значения InsertLocation: Before и After.|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Вставляет текст в абзац в указанном расположении. Возможные значения InsertLocation: Replace, Start и End.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestringssearchoptions)|[SearchResultCollection](searchresultcollection.md)|Выполняет поиск с помощью указанного объекта searchOptions в области объекта абзаца. Результат поиска — это коллекция объектов диапазона.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Выбирает абзац и переходит к нему в пользовательском интерфейсе Word. Возможные режимы выбора: Select, Start и End. Значение по умолчанию — Select.|

## Сведения о методе

### clear()
Очищает содержимое объекта абзаца. Пользователь может отменить операцию для очищенного содержимого.

#### Синтаксис
```js
paragraphObject.clear();
```

#### Параметры
Нет

#### Возвращаемое значение
void

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to clear the contents of the first paragraph.
        paragraphs.items[0].clear();    
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Cleared the contents of the first paragraph.');
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

### delete()
Удаляет абзац и его содержимое из документа.

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
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the text property for all of the paragraphs.
    context.load(paragraphs, 'text');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to delete the first paragraph.
        paragraphs.items[0].delete();    
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Deleted the first paragraph.');
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

### getHtml()
Возвращает HTML-представление объекта абзаца.

#### Синтаксис
```js
paragraphObject.getHtml();
```

#### Параметры
Нет

#### Возвращаемое значение
string

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a a set of commands to get the HTML of the first paragraph.
        var html = paragraphs.items[0].getHtml();    
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Paragraph HTML: ' + html.value);
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

### getOoxml()
Возвращает OOXML-представление объекта абзаца.

#### Синтаксис
```js
paragraphObject.getOoxml();
```

#### Параметры
Нет

#### Возвращаемое значение
string

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a a set of commands to get the OOXML of the first paragraph.
        var ooxml = paragraphs.items[0].getOoxml();    
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Paragraph OOXML: ' + ooxml.value);
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

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)
Вставляет разрыв в указанном месте. Разрыв можно вставить только в абзацы, которые содержатся в основном тексте документа, за исключением разрыва строки, который можно вставить в любой объект текста. Возможные значения insertLocation: Before и After.

#### Синтаксис
```js
paragraphObject.insertBreak(breakType, insertLocation);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|breakType|BreakType|Обязательный параметр. Тип разрыва, который необходимо добавить в документ.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Before и After.|

#### Возвращаемое значение
void

#### Дополнительные сведения
Невозможно вставить разрыв в верхние и нижние колонтитулы, сноски, концевые сноски, примечания и текстовые поля. 

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to insert a page break after the first paragraph.
        paragraph.insertBreak('page', 'After');    
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted a page break after the paragraph.');
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

### insertContentControl()
Включает объект абзаца в элемент управления содержимым форматированного текста.

#### Синтаксис
```js
paragraphObject.insertContentControl();
```

#### Параметры
Нет

#### Возвращаемое значение
[ContentControl](contentcontrol.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to wrap the first paragraph in a rich text content control.
        paragraph.insertContentControl(); 
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Wrapped the first paragraph in a content control.');
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
В примере [Word-Add-in-DocumentAssembly][paragraph.insertContentControl] показано, как можно использовать метод insertContentControl.

### insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Вставляет документ в текущий абзац в указанном расположении. Возможные значения InsertLocation: Start или End.

#### Синтаксис
```js
paragraphObject.insertFileFromBase64(base64File, insertLocation);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|base64File|string|Обязательный параметр. Содержимое файла с кодировкой base64, которое необходимо вставить.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Start и End.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert base64 encoded .docx at the beginning of the first paragraph.
        // This won't work unless you have a definition for getBase64().
        paragraph.insertFileFromBase64(getBase64(), Word.InsertLocation.start);
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted base64 encoded content at the beginning of the first paragraph.');
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

### insertHtml(html: string, insertLocation: InsertLocation)
Вставляет HTML в абзац в указанном расположении. Возможные значения InsertLocation: Replace, Start и End.

#### Синтаксис
```js
paragraphObject.insertHtml(html, insertLocation);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|html|string|Обязательный параметр. HTML-код, который необходимо вставить в абзац.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Replace, Start и End.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to insert HTML content at the end of the first paragraph.
        paragraph.insertHtml('<strong>Inserted HTML.</strong>', Word.InsertLocation.end);
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted HTML content at the end of the first paragraph.');
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

### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Вставляет рисунок в абзац в указанном расположении. Возможные значения InsertLocation: Before, After, Start и End.

#### Синтаксис
```js
paragraphObject.insertInlinePictureFromBase64(base64EncodedImage, insertLocation);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Обязательный параметр. HTML-код, который необходимо вставить в абзац.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Before, After, Start и End.|

#### Возвращаемое значение
[InlinePicture](inlinepicture.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        var b64encodedImg = "iVBORw0KGgoAAAANSUhEUgAAAB4AAAANCAIAAAAxEEnAAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACFSURBVDhPtY1BEoQwDMP6/0+XgIMTBAeYoTqso9Rkx1zG+tNj1H94jgGzeNSjteO5vtQQuG2seO0av8LzGbe3anzRoJ4ybm/VeKEerAEbAUpW4aWQCmrGFWykRzGBCnYy2ha3oAIq2MloW9yCCqhgJ6NtcQsqoIKdjLbFLaiACnYyf2fODbrjZcXfr2F4AAAAAElFTkSuQmCC";

        // Queue a command to insert a base64 encoded image at the beginning of the first paragraph.
        paragraph.insertInlinePictureFromBase64(b64encodedImg, Word.InsertLocation.start);
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Added an image to the first paragraph.');
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
[Word-Add-in-DocumentAssembly][paragraph.insertpicture] — еще один пример вставки изображения в абзац.

### insertOoxml(ooxml: string, insertLocation: InsertLocation)
Вставляет OOXML или wordProcessingML в абзац в указанном расположении. Возможные значения InsertLocation: Replace, Start и End.

#### Синтаксис
```js
paragraphObject.insertOoxml(ooxml, insertLocation);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|"ooxml"|string|Обязательный параметр. OOXML или wordProcessingML, которые необходимо вставить в абзац.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Replace, Start и End.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to insert Ooxml content into the first paragraph.
        var ooxmlContent = "<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>";
        paragraph.insertOoxml(ooxmlContent, Word.InsertLocation.end);
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted OOXML at the end of the first paragraph.');
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
Рекомендации по работе с OOXML см. в статье [Создание улучшенных надстроек для Word с помощью Office Open XML](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx).

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Вставляет абзац в указанном расположении. Возможные значения InsertLocation: Before и After.

#### Синтаксис
```js
paragraphObject.insertParagraph(paragraphText, insertLocation);
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
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to insert the paragraph after the current paragraph.
        paragraph.insertParagraph('Content of a new paragraph', Word.InsertLocation.after);
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted a new paragraph at the end of the first paragraph.');
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

### insertText(text: string, insertLocation: InsertLocation)
Вставляет текст в абзац в указанном расположении. Возможные значения InsertLocation: Replace, Start и End.

#### Синтаксис
```js
paragraphObject.insertText(text, insertLocation);
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
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to insert text into the end of the paragraph.
        paragraph.insertText('New text inserted into the paragraph.', Word.InsertLocation.end);
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted text at the end of the first paragraph.');
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
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to load font information for the paragraph.
        context.load(paragraph, 'font/size, font/name, font/color');
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            // Show the results of the load method. Here we show the
            // property values on the paragraph object. Note that we 
            // requested the style property in the first load command.
            var results = "<strong>Paragraph</strong>--" +
                          "--Font size: " + paragraph.font.size +
                          "--Font name: " + paragraph.font.name +
                          "--Font color: " + paragraph.font.color +
                          "--Style: " + paragraph.style;

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

### search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)
Выполняет поиск с помощью указанного объекта searchOptions в области объекта абзаца. Результат поиска — это коллекция объектов диапазона.

#### Синтаксис
```js
paragraphObject.search(searchText, searchOptions);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|searchText|string|Обязательный параметр. Текст для поиска.|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|Необязательный параметр. Параметры поиска.|

#### Возвращаемое значение
[SearchResultCollection](searchresultcollection.md)

### select(selectionMode: SelectionMode)
Выбирает абзац и переходит к нему в пользовательском интерфейсе Word.

#### Синтаксис
```js
paragraphObject.select(selectionMode);
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
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the last paragraph a create a 
        // proxy paragraph object.
        var paragraph = paragraphs.items[paragraphs.items.length - 1]; 
        
        // Queue a command to select the paragraph. The Word UI will 
        // move to the selected paragraph.
        paragraph.select();
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Selected the last paragraph.');
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

## Сведения о поддержке

Используйте [набор требований](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx) в проверках среды выполнения, чтобы обеспечить поддержку ведущей версии Word для своего приложения. Дополнительные сведения о требованиях ведущих приложений и серверов Office см. в статье [Требования для запуска надстроек Office](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx). 


[paragraph.insertContentControl]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L161 "insert content control"[paragraph.style]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L172 "set style" [paragraph.insertpicture]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L236 "insert picture"
