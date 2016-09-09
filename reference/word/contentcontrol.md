# Объект ContentControl (API JavaScript для Word)

Представляет элемент управления содержимым. Элементы управления содержимым — это связанные и, возможно, помеченные фрагменты документа, выполняющие роль контейнеров для определенных типов содержимого. Отдельные элементы управления содержимым могут содержать изображения, таблицы или абзацы форматированного текста. На данный момент поддерживаются только элементы управления содержимым "форматированный текст".

_Область применения: Word 2016, Word для iPad, Word для Mac._

## Свойства
| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|cannotDelete|bool|Возвращает или задает значение, указывающее, может ли пользователь удалить элемент управления содержимым. Является взаимоисключающим со свойством removeWhenEdited.|
|cannotEdit|bool|Возвращает или задает значение, указывающее, может ли пользователь изменять содержимое элемента управления содержимым.|
|color|string|Возвращает или задает цвет элемента управления содержимым. Цвет задается в формате #RRGGBB или с помощью имени цвета.|
|placeholderText|string|Возвращает или задает замещающий текст элемента управления содержимым. Если элемент управления содержимым пуст, отображается затемненный текст.|
|removeWhenEdited|bool|Возвращает или задает значение, указывающее, удаляется ли элемент управления содержимым после изменения. Является взаимоисключающим со свойством cannotDelete.|
|style|string|Возвращает или задает стиль для элемента управления содержимым. Это имя предустановленного или пользовательского стиля.|
|tag|string|Возвращает или задает тег для определения элемента управления содержимым. В примере надстройки [Silly stories](https://aka.ms/sillystorywordaddin) показано, как можно использовать свойство **tag**.|
|text|string|Возвращает текст элемента управления содержимым. Только для чтения.|
|title|string|Возвращает или задает заголовок для элемента управления содержимым.|

_Ознакомьтесь с [примерами](#примерами) доступа к свойствам._

## Связи
| Связь | Тип   |Описание|
|:---------------|:--------|:----------|
|appearance|**ContentControlAppearance**|Возвращает или задает внешний вид элемента управления содержимым. Допустимые значения: boundingBox, tags или hidden.|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Возвращает коллекцию объектов ContentControl в элементе управления содержимым. Только для чтения.|
|font|[Font](font.md)|Возвращает текстовый формат элемента управления содержимым. Используйте это свойство для получения и установки имени, размера, цвета и других свойств шрифта. Только для чтения.|
|id|**uint**|Возвращает целое число, представляющее собой идентификатор элемента управления содержимым. Только для чтения.|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Возвращает коллекцию объектов inlinePicture в элементе управления содержимым. Коллекция не содержит плавающие рисунки. Только для чтения.|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Возвращает коллекцию объектов paragraph в элементе управления содержимым. Только для чтения.|
|parentContentControl|[ContentControl](contentcontrol.md)|Возвращает элемент управления содержимым, содержащий элемент управления содержимым. Возвращает значение null, если родительского элемента управления содержимым не существует. Только для чтения.|
|type|**ContentControlType**|Возвращает тип элемента управления содержимым. На данный момент поддерживаются только элементы управления содержимым в формате RTF. Только для чтения.|

## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Очищает содержимое элемента управления содержимым. Пользователь может отменить очистку содержимого.|
|[delete(keepContent: bool)](#deletekeepcontent-bool)|void|Удаляет элемент управления содержимым и его содержимое. Если параметру keepContent присвоено значение true, содержимое не удаляется.|
|[getHtml()](#gethtml)|string|Возвращает HTML-представление для объекта элемента управления содержимым.|
|[getOoxml()](#getooxml)|string|Возвращает OOXML-представление объекта элемента управления содержимым.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Вставляет разрыв в указанном расположении. Разрыв можно вставить только в объекты, содержащиеся в основном тексте документа, за исключением разрыва строки, который можно вставить в любой объект в тексте. Возможные значения insertLocation: Before, After, Start и End.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Вставляет документ в текущий элемент управления содержимым в указанном расположении. Возможные значения insertLocation: Replace, Start и End.|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Вставляет HTML-код в элемент управления содержимым в указанном расположении. Возможные значения insertLocation: Replace, Start и End.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Вставляет встроенный рисунок в элемент управления содержимым в указанном расположении. Возможные значения insertLocation: Replace, Start и End. |
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Вставляет код OOXML или wordProcessingML в элемент управления содержимым в указанном расположении.  Возможные значения insertLocation: Replace, Start и End.|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Вставляет абзац в указанном расположении. Возможные значения insertLocation: Before, After, Start и End.|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Вставляет текст в элемент управления содержимым в указанном расположении. Возможные значения insertLocation: Replace, Start и End.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[SearchResultCollection](searchresultcollection.md)|Выполняет поиск с помощью указанного объекта searchOptions в области объекта элемента управления содержимым. Результат поиска — это коллекция объектов диапазона.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Выбирает элемент управления содержимым. При этом Word переходит к выделенному фрагменту. Возможные режимы выбора: Select, Start и End.|

## Сведения о методе

### clear()
Очищает содержимое элемента управления содержимым. Пользователь может отменить операцию для очищенного содержимого.

#### Синтаксис
```js
contentControlObject.clear();
```

#### Параметры
Нет

#### Возвращаемое значение
void

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the content controls collection.
    contentControls.load('text');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        if (contentControls.items.length === 0) {
            console.log("There isn't a content control in this document.");
        } else {
            
            // Queue a command to clear the contents of the first content control.
            contentControls.items[0].clear();
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Content control cleared of contents.');
            });      
        }
            
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

### delete(keepContent: bool)
Удаляет элемент управления содержимым и его содержимое. Если свойство keepContent имеет значение true, содержимое не будет удалено.

#### Синтаксис
```js
contentControlObject.delete(keepContent);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|keepContent|bool|Обязательный параметр. Указывает, следует ли удалить содержимое вместе с элементом управления содержимым. Если свойству keepContent задано значение true, содержимое не удаляется.|

#### Возвращаемое значение
void

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the content controls collection.
    contentControls.load('text');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        if (contentControls.items.length === 0) {
            console.log("There isn't a content control in this document.");
        } else {
            
            // Queue a command to delete the first content control. The
            // contents will remain in the document.
            contentControls.items[0].delete(true);
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Content control cleared of contents.');
            });      
        }
            
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
Возвращает HTML-представление объекта элемента управления содержимым.

#### Синтаксис
```js
contentControlObject.getHtml();
```

#### Параметры
Нет

#### Возвращаемое значение
string

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection that contains a specific tag.
    var contentControlsWithTag = context.document.contentControls.getByTag('Customer-Address');
    
    // Queue a command to load the tag property for all of content controls. 
    context.load(contentControlsWithTag, 'tag');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTag.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to get the HTML contents of the first content control.
            var html = contentControlsWithTag.items[0].getHtml();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Content control HTML: ' + html.value);
            });
        }
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
Возвращает OOXML-представление объекта элемента управления содержимым.

#### Синтаксис
```js
contentControlObject.getOoxml();
```

#### Параметры
Нет

#### Возвращаемое значение
string

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to get the OOXML contents of the first content control.
            var ooxml = contentControls.items[0].getOoxml();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Content control OOXML: ' + ooxml.value);
            });
        }
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
Вставляет разрыв в указанном расположении. Разрыв можно вставить только в объекты, содержащиеся в основном тексте документа, за исключением разрыва строки, который можно вставить в любой объект в тексте. Возможные значения InsertLocation: Before, After, Start и End.

#### Синтаксис
```js
contentControlObject.insertBreak(breakType, insertLocation);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|breakType|BreakType|Обязательный параметр. Тип разрыва (breakType.md)|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Before, After, Start и End.|

#### Возвращаемое значение
void

#### Дополнительные сведения
За исключением разрывов строк, вы не можете вставлять разрывы в объекты, содержащиеся в верхних и нижних колонтитулах, сносках, концевых сносках, примечаниях и текстовых полях.  

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a commmand to load the id property for all of content controls. 
    context.load(contentControls, 'id');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion. We now will have 
    // access to the content control collection.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to insert a page break after the first content control. 
            contentControls.items[0].insertBreak('page', "After");
            
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion. 
            return context.sync()
                .then(function () {
                    console.log('Inserted a page break after the first content control.');    
            });
        }
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
Вставляет документ в текущий элемент управления содержимым в указанном расположении. Возможные значения insertLocation: Replace, Start и End.

#### Синтаксис
```js
contentControlObject.insertFileFromBase64(base64File, insertLocation);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|base64File|string|Обязательный параметр. Содержимое файла с кодировкой Base64, которое необходимо вставить.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Replace, Start и End.|

#### Возвращаемое значение
[Range](range.md)

### insertHtml(html: string, insertLocation: InsertLocation)
Вставляет HTML-код в элемент управления содержимым в указанном расположении. Возможные значения insertLocation: Replace, Start и End.

#### Синтаксис
```js
contentControlObject.insertHtml(html, insertLocation);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|html|string|Обязательный параметр. HTML-код, который необходимо вставить в элемент управления содержимым.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Replace, Start и End.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to put HTML into the contents of the first content control.
            contentControls.items[0].insertHtml('<strong>HTML content inserted into the content control.</strong>', 'Start');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted HTML in the first content control.');
            });
        }
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
Вставляет встроенный рисунок в элемент управления содержимым в указанном расположении. Возможные значения insertLocation: Replace, Start и End.

#### Синтаксис
contentControlObject.insertInlinePictureFromBase64(image, insertLocation);

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Обязательный параметр. Вставляемое в элемент управления содержимым изображение в кодировке base64.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Replace, Start и End.|

#### Возвращаемое значение
[InlinePicture](inlinepicture.md)



### insertOoxml(ooxml: string, insertLocation: InsertLocation)
Вставляет OOXML или wordProcessingML в элемент управления содержимым в указанном расположении. Возможные значения insertLocation: Replace, Start и End.

#### Синтаксис
```js
contentControlObject.insertOoxml(ooxml, insertLocation);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|"ooxml"|string|Обязательный параметр. OOXML или wordProcessingML, который необходимо вставить в элемент управления содержимым.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Replace, Start и End.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to put OOXML into the contents of the first content control.
            contentControls.items[0].insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", "End");
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted OOXML in the first content control.');
            });
        }
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
Рекомендации по работе с OOXML см. в статье [Создание улучшенных надстроек для Word с использованием Office Open XML](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx).

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Вставляет абзац в указанном расположении. Возможные значения insertLocation: Before, After, Start и End.

#### Синтаксис
```js
contentControlObject.insertParagraph(paragraphText, insertLocation);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|paragraphText|string|Обязательный параметр. Текст абзаца, который необходимо вставить.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Before, After, Start и End.|

#### Возвращаемое значение
[Paragraph](paragraph.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to insert a paragraph after the first content control. 
            contentControls.items[0].insertParagraph('Text of the inserted paragraph.', 'After');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted a paragraph after the first content control.');
            });
        }
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
Вставляет текст в элемент управления содержимым в указанном расположении. Возможные значения insertLocation: Replace, Start и End.

#### Синтаксис
```js
contentControlObject.insertText(text, insertLocation);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|text|string|Обязательный параметр. Текст, который необходимо вставить в элемент управления содержимым.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Replace, Start и End.|

#### Возвращаемое значение
[Range](range.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to replace text in the first content control. 
            contentControls.items[0].insertText('Replaced text in the first content control.', 'Replace');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Replaced text in the first content control.');
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

В примере надстройки [Silly stories](https://aka.ms/sillystorywordaddin) показано, как использовать метод **insertText**.

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
    
    // Create a proxy range object for the current selection.
    var range = context.document.getSelection();
    
    // Queue a commmand to create the content control.
    var myContentControl = range.insertContentControl();
    myContentControl.tag = 'Customer-Address';
    myContentControl.title = ' has t';
    myContentControl.style = 'Heading 2';
    myContentControl.insertText('One Microsoft Way, Redmond, WA 98052', 'replace');
    myContentControl.cannotEdit = true;
    
    // Queue a command to load the id property for the content control you created.
    context.load(myContentControl, 'id');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Created content control with id: ' + myContentControl.id);
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
Выполняет поиск с помощью указанного объекта searchOptions в области объекта элемента управления содержимым. Результат поиска — это коллекция объектов диапазона.

#### Синтаксис
```js
contentControlObject.search(searchText, searchOptions);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|searchText|string|Обязательный параметр. Текст для поиска.|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|Необязательный параметр. Параметры поиска.|

#### Возвращаемое значение
[SearchResultCollection](searchresultcollection.md)

### select(selectionMode: SelectionMode)
Выбирает элемент управления контентом. При этом Word переходит к выделенному фрагменту. Возможные режимы выбора: Select, Start и End.

#### Синтаксис
```js
contentControlObject.select(selectionMode);
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
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to select the first content control.
            contentControls.items[0].select();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Selected the first content control.');
            });
        }
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

### Загрузка всех свойств элемента управления содержимым
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to load the properties on the first content control. 
            contentControls.items[0].load(  'appearance,' +
                                            'cannotDelete,' +
                                            'cannotEdit,' +
                                            'id,' +
                                            'placeHolderText,' +
                                            'removeWhenEdited,' +
                                            'title,' +
                                            'text,' +
                                            'type,' +
                                            'style,' +
                                            'tag,' +
                                            'font/size,' +
                                            'font/name,' +
                                            'font/color');             
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Property values of the first content control:' + 
                        '   ----- appearance: ' + contentControls.items[0].appearance + 
                        '   ----- cannotDelete: ' + contentControls.items[0].cannotDelete +
                        '   ----- cannotEdit: ' + contentControls.items[0].cannotEdit +
                        '   ----- color: ' + contentControls.items[0].color +
                        '   ----- id: ' + contentControls.items[0].id +
                        '   ----- placeHolderText: ' + contentControls.items[0].placeholderText +
                        '   ----- removeWhenEdited: ' + contentControls.items[0].removeWhenEdited +
                        '   ----- title: ' + contentControls.items[0].title +
                        '   ----- text: ' + contentControls.items[0].text +
                        '   ----- type: ' + contentControls.items[0].type +
                        '   ----- style: ' + contentControls.items[0].style +
                        '   ----- tag: ' + contentControls.items[0].tag +
                        '   ----- font size: ' + contentControls.items[0].font.size +
                        '   ----- font name: ' + contentControls.items[0].font.name +
                        '   ----- font color: ' + contentControls.items[0].font.color);
            });
        }
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