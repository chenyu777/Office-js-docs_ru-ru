# Объект ContentControlCollection (API JavaScript API для Word)

Содержит коллекцию объектов ContentControl. Элементы управления контентом — это связанные и, возможно, помеченные фрагменты документа, выполняющие роль контейнеров для определенных типов содержимого. Отдельные элементы управления контентом могут содержать изображения, таблицы или абзацы форматированного текста. На данный момент поддерживаются только элементы управления содержимым "форматированный текст".

_Область применения: Word 2016, Word для iPad, Word для Mac_

## Свойства
| Свойство   | Тип|Описание
|:---------------|:--------|:----------|
|items|[ContentControl[]](contentcontrol.md)|Коллекция объектов contentControl. Только для чтения.|

_Ознакомьтесь с [примерами](#property-access-examples) доступа к свойствам._

## Связи
Нет


## Методы

| Метод   | Возвращаемый тип|Описание|
|:---------------|:--------|:----------|
|[getById(id: number)](#getbyidid-number)|[ContentControl](contentcontrol.md)|Возвращает элемент управления содержимым по его идентификатору.|
|[getByTag(tag: string)](#getbytagtag-string)|[ContentControlCollection](contentcontrolcollection.md)|Возвращает элементы управления содержимым с указанным тегом.|
|[getByTitle(title: string)](#getbytitletitle-string)|[ContentControlCollection](contentcontrolcollection.md)|Возвращает элементы управления контентом с указанным заголовком.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе

### getById(id: number)
Возвращает элемент управления содержимым по его идентификатору.

#### Синтаксис
```js
contentControlCollectionObject.getById(id);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|id|number|Обязательный параметр. Идентификатор элемента управления контентом.|

#### Возвращаемое значение
[ContentControl](contentcontrol.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
	
	// Create a proxy object for the content control that contains a specific id.
	var contentControl = context.document.contentControls.getById(30086310);
		
	// Queue a command to load the text property for a content control. 
	context.load(contentControl, 'text');
	
	// Synchronize the document state by executing the queued commands, 
	// and return a promise to indicate task completion.
	return context.sync().then(function () {
		console.log('The content control with that Id has been found in this document.'); 
	});  
})
.catch(function (error) {
	console.log('Error: ' + JSON.stringify(error));
	if (error instanceof OfficeExtension.Error) {
		console.log('Debug info: ' + JSON.stringify(error.debugInfo));
	}
});
```

### getByTag(tag: string)
Возвращает элементы управления содержимым с указанным тегом.

#### Синтаксис
```js
contentControlCollectionObject.getByTag(tag);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|tag|string|Обязательный параметр. Тег, установленный на элемент управления контентом.|

#### Возвращаемое значение
[ContentControlCollection](contentcontrolcollection.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection that contains a specific tag.
    var contentControlsWithTag = context.document.contentControls.getByTag('Customer-Address');
        
    // Queue a command to load the text property for all of content controls with a specific tag. 
    context.load(contentControlsWithTag, 'text');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTag.items.length === 0) {
            console.log("There isn't a content control with a tag of Customer-Address in this document.");
        } else {
            console.log('The first content control with the tag of Customer-Address has this text: ' + contentControlsWithTag.items[0].text);    
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
[Word-Add-in-DocumentAssembly][contentControls.getByTag] — еще один пример использования метода getByTag.


### getByTitle(title: string)
Возвращает элементы управления контентом с указанным заголовком.

#### Синтаксис
```js
contentControlCollectionObject.getByTitle(title);
```

#### Параметры
| Параметр   | Тип|Описание|
|:---------------|:--------|:----------|
|title|string|Обязательный параметр. Заголовок элемента управления контентом.|

#### Возвращаемое значение
[ContentControlCollection](contentcontrolcollection.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection that contains a specific title.
    var contentControlsWithTitle = context.document.contentControls.getByTitle('Enter Customer Address Here');
        
    // Queue a command to load the text property for all of content controls with a specific title. 
    context.load(contentControlsWithTitle, 'text');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTitle.items.length === 0) {
            console.log("There isn't a content control with a title of 'Enter Customer Address Here' in this document.");
        } else {
            console.log('The first content control with the title of 'Enter Customer Address Here' has this text: ' + contentControlsWithTitle.items[0].text);    
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
[Word-Add-in-DocumentAssembly][contentControls.getByTitle] — еще один пример использования метода getByTitle.

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
                                            'color,' +
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

В примере надстройки [Silly stories](https://aka.ms/sillystorywordaddin) показано, как метод **load** используется для загрузки коллекции элементов управления контентом со свойствами **tag** и **title**.

## Сведения о поддержке

Используйте [набор требований](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx) в проверках среды выполнения, чтобы обеспечить поддержку ведущей версии Word для своего приложения. Дополнительные сведения о требованиях ведущих приложений и серверов Office см. в статье [Требования для запуска надстроек Office](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx). 


[contentControls.getByTag]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L300 "get by tag" [contentControls.getByTitle]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L331 "get by title"


