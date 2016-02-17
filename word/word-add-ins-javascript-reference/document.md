# Объект Document (API JavaScript для Word)

Объект Document — это объект верхнего уровня. Объект Document содержит один или несколько разделов, элементы управления контентом и основной текст с содержанием документа.

_Область применения: Word 2016, Word для iPad, Word для Mac_

## Свойства
| Свойство   | Тип|Описание
|:---------------|:--------|:----------|
|saved|bool|Указывает, сохранены ли изменения, внесенные в документ. Значение true указывает на то, что с момента последнего сохранения в документ не вносились изменения. Только для чтения.|

## Связи
| Связь | Тип|Описание|
|:---------------|:--------|:----------|
|body|[Body](body.md)|Возвращает основной текст документа. Основной текст — это текст, не содержащий верхних и нижних колонтитулов, сносок, текстовых полей и т. д. Только для чтения.|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Возвращает коллекцию объектов элементов управления содержимым в текущем документе. Сюда входят элементы управления содержимым в основном тексте документа, верхних и нижних колонтитулах, текстовых полях и т. д. Только для чтения.|
|sections|[SectionCollection](sectioncollection.md)|Возвращает коллекцию объектов разделов в документе. Только для чтения.|

## Методы

| Метод   | Возвращаемый тип|Описание|
|:---------------|:--------|:----------|
|[getSelection()](#getselection)|[Range](range.md)|Возвращает текущий выбранный фрагмент документа. Получение нескольких выбранных фрагментов не поддерживается.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|
|[save()](#save)|void|Сохраняет документ. При этом используется соглашение об именовании файлов Word по умолчанию, если документ ранее не сохранялся.|

## Сведения о методе

### getSelection()
Возвращает текущий выбранный фрагмент документа. Получение нескольких выбранных фрагментов не поддерживается.

#### Синтаксис
```js
documentObject.getSelection();
```

#### Параметры
Нет

#### Возвращаемое значение
[Range](range.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    var textSample = 'This is an example of the insert text method. This is a method ' + 
        'which allows users to insert text into a selection. It can insert text into a ' +
        'relative location or it can overwrite the current selection. Since the ' +
        'getSelection method returns a range object, look up the range object documentation ' +
        'for everything you can do with a selection.';
    
    // Create a range proxy object for the current selection.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert text at the end of the selection.
    range.insertText(textSample, Word.InsertLocation.end);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Inserted the text at the end of the selection.');
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
    
    // Create a proxy object for the document.
    var thisDocument = context.document;
    
    // Queue a command to load content control properties.
    context.load(thisDocument, 'contentControls/id, contentControls/text, contentControls/tag');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (thisDocument.contentControls.items.length !== 0) {
            for (var i = 0; i < thisDocument.contentControls.items.length; i++) {
                console.log(thisDocument.contentControls.items[i].id);
                console.log(thisDocument.contentControls.items[i].text);
                console.log(thisDocument.contentControls.items[i].tag);
            }
        } else {
            console.log('No content controls in this document.');
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

### save()
Сохраняет документ. При этом используется соглашение об именовании файлов Word по умолчанию, если документ ранее не сохранялся.

#### Синтаксис
```js
documentObject.save();
```

#### Параметры
Нет

#### Возвращаемое значение
void

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document.
    var thisDocument = context.document;

    // Queue a commmand to load the document save state (on the saved property).
    context.load(thisDocument, 'saved');    
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        if (thisDocument.saved === false) {
            // Queue a command to save this document.
            thisDocument.save();
            
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Saved the document');
            });
        } else {
            console.log('The document has not changed since the last save.');
        }
    });  
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## Сведения о поддержке

Используйте [набор требований](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx) в проверках среды выполнения, чтобы обеспечить поддержку ведущей версии Word для своего приложения. Дополнительные сведения о требованиях ведущих приложений и серверов Office см. в статье [Требования для запуска надстроек Office](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx). 
