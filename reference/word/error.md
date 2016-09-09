# Объект OfficeExtension.Error (API JavaScript для Word)

Представляет ошибки, которые возникают при использовании API JavaScript для Word.

_Область применения: Word 2016, Word для iPad, Word для Mac_

## Свойства
| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|code|string|Возвращает тип ошибки. Возможные значения: AccessDenied, GeneralException, ActivityLimitReached, InvalidArgument, ItemNotFound и NotImplemented. <!-- Values come from OfficeExtension.Error and Word.ErrorCodes. -->|
|debugInfo|string|Возвращает значение, которое указывает, что произошло при возникновении ошибки. Это значение предназначено для использования только во время разработки и отладки.  |
|сообщение |string| Возвращает локализованную понятную для пользователя строку, которая соответствует коду ошибки.|
|name |string| Возвращает значение OfficeExtension.Error. |
|traceMessages |string[]| Возвращает массив значений, которые соответствуют сообщениям инструментирования, заданным с помощью синтаксиса context.trace(); |

_Ознакомьтесь с [примерами](#примерами) доступа к свойствам._

## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[toString()](#tostring)|строка|Возвращает код ошибки и сообщение в следующем формате: "{0}: {1}", код, сообщение.|

## Сведения о методе

### toString()
Возвращает код ошибки и сообщение в следующем формате: "{0}: {1}", код, сообщение.

#### Синтаксис
```js
error.toString()
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

    // Queue a commmand to insert text in to the beginning of the body.
    // This will cause an OfficeExtension.Error.
    body.insertText(0);

    // Synchronize the document state by executing the queued-up commands,
    // and return a promise to indicate task completion.
    return context.sync();
})
.catch(function (error) {
    if (error instanceof OfficeExtension.Error) {
        console.log('Error code and message: ' + error.toString());
    }
});

```

## Примеры доступа к свойствам

### Инструментирование сообщений трассировки

В следующем примере показано, как можно инструментировать пакет команд, чтобы определить, где произошла ошибка. Первый пакет успешно вставляет первые два абзаца в документ. Второй пакет успешно вставляет третий и четвертый абзацы, но дает сбой при вызове для вставки пятого абзаца. Все остальные команды в пакете, в том числе команда, которая добавляет пятое сообщение трассировки, не выполняются. В этом случае ошибка произошла после вставки четвертого абзаца и перед добавлением пятого сообщения трассировки.

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert the paragraph at the end of the document body.
    // Start a batch of commands.
    body.insertParagraph('1st paragraph', Word.InsertLocation.end);
    // Queue a command for instrumenting this part of the batch.
    context.trace('1st paragraph successful');

    body.insertParagraph('2nd paragraph', Word.InsertLocation.end);
    context.trace('2nd paragraph successful');

    // Synchronize the document state by executing the queued-up commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Queue a commmand to insert the paragraph at the end of the document body.
        // Start a new batch of commands.
        body.insertParagraph('3rd paragraph', Word.InsertLocation.end);
        context.trace('3rd paragraph successful');

        body.insertParagraph('4th paragraph', Word.InsertLocation.end);
        context.trace('4th paragraph successful');

        // This command will cause an error. The trace messages in the queue up to
        // this point will be available via Error.traceMessages.
        body.insertParagraph(0, '5th paragraph', Word.InsertLocation.end);
        // Queue a command for instrumenting this part of the batch.
        // This trace message will not be set on Error.traceMessages.
        context.trace('5th paragraph successful');
    }).then(context.sync);
})
.catch(function (error) {
    if (error instanceof OfficeExtension.Error) {
        console.log('Trace messages: ' + error.traceMessages);
    }
});

// Output: "Trace messages: 3rd paragraph successful,4th paragraph successful"

```
