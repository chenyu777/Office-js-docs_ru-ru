# Справочник по API JavaScript для Word

Word предоставляет большой набор API. Вы можете использовать эти API для создания надстроек, взаимодействующих с контентом и метаданными документов. С помощью этих API вы сможете создавать привлекательные приложения, интегрируемые с Word и расширяющие возможности этой программы. Вы можете импортировать и экспортировать контент, собирать новые документы на основе различных источников данных, выполнять интеграцию с рабочими процессами документов и создавать пользовательские решения для работы с документами.

Для взаимодействия с объектами и метаданными в документе Word вы можете использовать два указанных ниже API JavaScript.

- API JavaScript для Word: впервые появился в Office 2016.
- [API JavaScript для Office](../javascript-api-for-office.md) (Office.js): впервые появился в Office 2013.

## API JavaScript для Word

API JavaScript для Word загружается с помощью файла Office.js. Этот API изменяет способ взаимодействия с объектами, например с документами и абзацами. Вместо набора отдельных асинхронных API для получения и обновления каждого из этих объектов новый API JavaScript для Word предоставляет прокси-объекты JavaScript, которые соответствуют реальным объектам, выполняемым в Word. Вы можете напрямую взаимодействовать с этими прокси-объектами, синхронно считывая и записывая их свойства, а также вызывая синхронные методы для операций над ними. Эти взаимодействия с прокси-объектами не сразу реализуются в выполняющихся сценариях. Метод **context.sync** синхронизирует состояние запущенного JavaScript и реальных объектов в Office, выполняя поставленные в очередь инструкции и получая свойства загруженных объектов Word для их использования в сценарии.

## API JavaScript для Office

Файл Office.js можно получить из следующих расположений:

* https://appsforoffice.microsoft.com/lib/1/hosted/office.js. Используйте этот ресурс для надстроек, выполняемых в рабочей среде.
* https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Используйте этот ресурс при испытаниях предварительных версий функций.

Если вы используете [Visual Studio](https://www.visualstudio.com/products/free-developer-offers-vs), то, чтобы получить шаблоны проектов, включающие файл Office.js, вы можете скачать [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx).  Кроме того, [чтобы получить файл Office.js, вы можете воспользоваться NuGet](https://www.nuget.org/packages/Microsoft.Office.js/).

Если вы используете TypeScript и у вас есть npm, то вы можете получить определения TypeScript, выполнив в интерфейсе командной строки следующую команду: ```typings install office-js --ambient```.

## Запуск надстроек Word

Чтобы запустить надстройку, воспользуйтесь обработчиком событий Office.initialize. Дополнительные сведения об инициализации надстроек см. в статье [Общие сведения об API](../../docs/develop/understanding-the-javascript-api-for-office.md).

Надстройки, предназначенные для Word 2016, выполняются путем передачи функции в метод **Word.run()**. У функции, передаваемой в метод **run**, обязательно должен быть аргумент контекста. Этот [объект контекста](../../reference/word/requestcontext.md) отличается от объекта контекста, который вы получаете из объекта Office, но также используется для взаимодействия со средой выполнения Word. Объект контекста предоставляет доступ к объектной модели API JavaScript для Word. В примере ниже показано, как инициализировать и выполнить надстройку Word с помощью метода **Word.run()**.

```js
    (function () {
        "use strict";

        // The initialize event handler must be run on each page to initialize Office JS.
        // You can add optional custom initialization code that will run after OfficeJS
        // has initialized.
        Office.initialize = function (reason) {
            // The reason object tells how the add-in was initialized. The values can be:
            // inserted - the add-in was inserted to an open document.
            // documentOpened - the add-in was already inserted in to the document and the document was opened.

            // Checks for the DOM to load using the jQuery ready function.
            $(document).ready(function () {
                // Set your optional initialization code.
                // You can also load saved settings from the Office object.
            });
        };

        // Run a batch operation against the Word JavaScript API object model.
        // Use the context argument to get access to the Word document.
        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;
            // ...
        })
    })();
```

### Синхронизация документов Word с помощью прокси-объектов API JavaScript для Word

Объектная модель API JavaScript для Word нестрого связана с объектами в Word. Объекты API JavaScript для Word представляют собой прокси-объекты для объектов в документе Word. Действия, выполняемые над прокси-объектами, не будут реализованы в Word, пока не будет синхронизировано состояние документа. И наоборот, состояние документа Word не будет реализовано в прокси-объектах, пока оно не будет синхронизировано. Чтобы синхронизировать состояние документа, выполните метод **context.sync()**. В примере ниже показано, как создать прокси-объект основного текста и помещенную в очередь команду для загрузки свойства текста в прокси-объекте основного текста и как использовать метод **context.sync()** для синхронизации основного текста документа Word с прокси-объектом основного текста.

```js
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        // The body object hasn't been set with any property values.
        var body = context.document.body;

        // Queue a command to load the text property for the proxy document body object.
        context.load(body, 'text');

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });
    })
```

### Выполнение пакета команд

У прокси-объектов Word есть методы для доступа к объектной модели и ее обновления. Эти методы выполняются последовательно в том порядке, в котором они были поставлены в очередь в пакете. При вызове метода context.sync() выполняются все команды, помещенные в очередь в пакете.

В примере ниже показано, как работает очередь команд. При вызове метода **context.sync()** в Word выполняется [команда загрузки](../../reference/word/loadoption.md) основного текста. Затем выполняется команда вставки текста в основной текст в Word. Результаты выполнения команд возвращаются в прокси-объект основного текста. Значение свойства **body.text** в API JavaScript для Word представляет собой значение основного текста документа Word <u>перед тем, как</u> текст был вставлен в документ Word.


```js
    // Run a batch operation against the Word JavaScript API.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        var body = context.document.body;

        // Queue a command to load the text property of the proxy body object.
        context.load(body, 'text');

        // Queue a command to insert text into the end of the Word document body.
        body.insertText('This is text inserted after loading the body.text property',
                        Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });
    })
```

## Открытые спецификации API для Word

По мере проектирования и разработки новых API для надстроек Word мы публикуем их на странице [Открытые спецификации API](../../reference/openspec.md), чтобы вы могли написать свои отзывы о них. Узнайте, какие новые функции запланированы в API JavaScript для Word, и сообщите свое мнение о проектируемых спецификациях.

## Дополнительные ресурсы

* [Обзор надстроек Word](../../docs/word/word-add-ins-programming-overview.md )
* [Обзор платформы надстроек Office](../../docs/overview/office-add-ins.md)
* [Примеры надстроек Word на сайте GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&query=Word)
