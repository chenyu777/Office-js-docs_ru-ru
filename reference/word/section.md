# Объект Section (API JavaScript для Word)

Представляет раздел в документе Word.

_Область применения: Word 2016, Word для iPad, Word для Mac_

## Свойства
Нет

## Связи
| Связь | Тип   |Описание|
|:---------------|:--------|:----------|
|Основной текст|[Body](body.md)|Возвращает текст раздела. Сюда не относятся колонтитулы и другие метаданные раздела. Только для чтения.|

## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[getFooter(type: HeaderFooterType)](#getfootertype-headerfootertype)|[Body](body.md)|Возвращает один из нижних колонтитулов раздела.|
|[getHeader(type: HeaderFooterType)](#getheadertype-headerfootertype)|[Body](body.md)|Возвращает один из верхних колонтитулов раздела.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе

### getFooter(type: HeaderFooterType)
Возвращает один из нижних колонтитулов раздела.

#### Синтаксис
```js
sectionObject.getFooter(type);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|type|HeaderFooterType|Обязательный параметр. Тип нижнего колонтитула, который необходимо возвратить. Возможные значения: primary, firstPage или evenPages.|

#### Возвращаемое значение
[Body](body.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy sectionsCollection object.
    var mySections = context.document.sections;
    
    // Queue a commmand to load the sections.
    context.load(mySections, 'body/style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Create a proxy object the primary footer of the first section. 
        // Note that the footer is a body object.
        var myFooter = mySections.items[0].getFooter("primary");
        
        // Queue a command to insert text at the end of the footer.
        myFooter.insertText("This is a footer.", Word.InsertLocation.end);
        
        // Queue a command to wrap the header in a content control.
        myFooter.insertContentControl();
                              
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Added a footer to the first section.");
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
### getHeader(type: HeaderFooterType)
Возвращает один из верхних колонтитулов раздела.

#### Синтаксис
```js
sectionObject.getHeader(type);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|type|HeaderFooterType|Обязательный параметр. Тип колонтитула, который необходимо возвратить. Возможные значения: primary, firstPage или evenPages.|

#### Возвращаемое значение
[Body](body.md)

#### Примеры
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy sectionsCollection object.
    var mySections = context.document.sections;
    
    // Queue a commmand to load the sections.
    context.load(mySections, 'body/style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Create a proxy object the primary header of the first section. 
        // Note that the header is a body object.
        var myHeader = mySections.items[0].getHeader("primary");
        
        // Queue a command to insert text at the end of the header.
        myHeader.insertText("This is a header.", Word.InsertLocation.end);
        
        // Queue a command to wrap the header in a content control.
        myHeader.insertContentControl();
                              
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Added a header to the first section.");
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
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|param|object|Необязательный параметр. Принимает имена параметров и связей в виде строки с разделителями или массива. Либо укажите объект [loadOption](loadoption.md).|

#### Возвращаемое значение
void

## Сведения о поддержке
Используйте [набор требований](../office-add-in-requirement-sets.md) в проверках среды выполнения, чтобы обеспечить поддержку ведущей версии Word для своего приложения. Дополнительные сведения о требованиях ведущих приложений и серверов Office см. в статье [Требования для запуска надстроек Office](../../docs/overview/requirements-for-running-office-add-ins.md).