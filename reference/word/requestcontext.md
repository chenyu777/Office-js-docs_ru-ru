# Объект RequestContext (API JavaScript для Word)

Объект RequestContext упрощает отправку запроса приложению Word из надстройки Word, так как оба приложения выполняются в различных процессах.

_Область применения: Word 2016, Word для iPad, Word для Mac_

## Свойства
Нет

## Методы

| Метод         | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[load(object: object, option: object)](#loadobject-object-option-object)  |void     |Заполняет объект прокси, созданный на уровне JavaScrypt, свойством и параметрами, которые указаны в параметре.|
|[sync()](#sync)  |Объект Promise |Отправляет очередь запросов в Word и возвращает объект Promise, который может использоваться для построения цепочки дальнейших действий.|

## Сведения о методе

### load(object: object, option: object)
Заполняет объект прокси, созданный на уровне JavaScrypt, свойством и параметрами, которые указаны в параметре.

#### Синтаксис
```js
requestContextObject.load(object, loadOption);
```

#### Параметры
| Параметр       | Тип    |Описание|
|:----------------|:--------|:----------|
|object|object|Необязательный параметр. Укажите имя объекта, который необходимо загрузить.|
|option|[loadOption](loadoption.md)|Необязательный параметр, но рекомендуется его использовать. Указывает параметры загрузки, например "выбрать", "развернуть", "пропустить" и "сверху". |

#### Возвращаемое значение
void

##### Примеры

Ниже приводится пример использования контекста запроса для загрузки свойства текста в коллекцию абзаца.

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

#### Дополнительные сведения

После добавления отслеживаемых объектов необходимо вызвать метод load().

### sync()
Отправляет очередь запросов в Word и возвращает объект Promise, который может использоваться для построения цепочки дальнейших действий.

#### Синтаксис
```js
requestContextObject.sync();
```

#### Параметры
Нет

#### Возвращаемое значение
Объект Promise.

#### Примеры

Ниже приводится пример двойного использования метода синхронизации: 1) для загрузки коллекции элементов управления содержимым со свойством текста для каждого элемента управления и 2) для очистки содержимого первого элемента управления содержимым в коллекции.

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

## Сведения о поддержке
Используйте [набор требований](../office-add-in-requirement-sets.md) в проверках среды выполнения, чтобы обеспечить поддержку ведущей версии Word для своего приложения. Дополнительные сведения о требованиях ведущих приложений и серверов Office см. в статье [Требования для запуска надстроек Office](../../docs/overview/requirements-for-running-office-add-ins.md).