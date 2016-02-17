# Общие сведения о программировании надстроек Word

_Область применения: Word 2016, Word для iPad, Word для Mac_

Word 2016 использует новую объектную модель для работы с объектами Word. Эта объектная модель — дополнение к уже существующей модели, предоставляемой файлом Office.js для создания надстроек Word. Доступ к этой объектной модели осуществляется через код JavaScript, размещенный в веб-приложении.

## Манифест

В новом API JavaScript надстроек Word используется тот же формат манифеста, что и для модели надстроек Office 2013. Манифест описывает, где размещена надстройка, как она отображается, разрешения и другие сведения. Узнайте больше о том, как можно настраивать [манифесты надстроек](https://msdn.microsoft.com/en-us/library/office/fp161044.aspx). 

Существует несколько способов для публикации манифестов надстроек Word. Узнайте, как можно [публиковать надстройки для Office](https://msdn.microsoft.com/EN-US/library/office/fp123515.aspx) в сетевой папке, каталоге приложений или Магазине Office.

## Общие сведения об API JavaScript для Word

API JavaScript для Word загружается с помощью Office.js. Этот сценарий предоставляет набор прокси-объектов JavaScript, используемых для постановки в очередь набора команд, которые взаимодействуют с содержанием документа Word. Эти команды запускаются как единый пакет. В результате в документе Word выполняются действия, например вставка содержимого и синхронизация объектов Word с прокси-объектами JavaScript. 

### Запуск надстройки

Рассмотрим, что вам потребуется при запуске надстройки. Во всех надстройках должен быть обработчик событий Office.initialize. Дополнительные сведения об инициализации надстроек см. в статье [общие сведения об API](https://msdn.microsoft.com/EN-US/library/fp160953.aspx).  

Надстройки Word выполняются передачей функции в метод Word.run(). Функции, передаваемой в метод выполнения, обязательно должен быть присвоен контекстный аргумент. Этот [контекстный объект](word-add-ins-javascript-reference/requestcontext.md) отличается от контекстного объекта, получаемого из объекта Office, несмотря на то что он используется для той же цели — взаимодействия со средой выполнения Word. Контекстный объект предоставляет доступ к объектной модели JavaScript для Word. Давайте рассмотрим комментарии, а также код простой надстройки Word:

**Пример 1. Инициализация и выполнение надстройки Word**

```javascript
    (function () {
        "use strict";

        // The initialize event handler is run each time the page is loaded.
        Office.initialize = function (reason) {
            
            // Checks for the DOM to load using the jQuery ready function.
            $(document).ready(function () {
                // Set your initialization code. You can use the reason 
                // argument to determine how the add-in was loaded.
                // You can also load saved settings from the Office object.
            });
        };

        // Run a batch operation against the Word object model.
        // Use the context argument to get access to the Word document.
        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;
        })
    })();
```

Пример 1. Показывает основной код, необходимый для создания надстройки Word. Он инициализирует сценарий Office.js и содержит метод выполнения для взаимодействия с документом Word.

### Прокси-объекты

Объектная модель JavaScript для Word свободно объединяется с объектами в Word. Для настоящих объектов в документе Word объекты JavaScript — это прокси-объекты. Все действия над прокси-объектами не реализуются в Word, а состояние документа Word — в прокси-объектах, пока оно не будет синхронизировано. Состояние документа синхронизируется при выполнении context.sync(). Метод sync() фактически выполняет набор команд в очереди для каждого прокси-объекта. Пример 2 показывает создание прокси-объекта основного текста и команду в очереди для загрузки свойства текста в такой объект, а затем синхронизацию основного текста в документе Word с прокси-объектом основного текста. 

**Пример 2. Синхронизация текста документа с прокси-объектом текста.**

```javascript
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

### Очередь команд

У прокси-объектов Word есть методы для доступа и обновления объектной модели. Эти методы выполняются последовательно в том порядке, в каком они были добавлены в пакет. Пакет команд формируется до вызова context.sync(). Выполняются все команды в очереди всех объектов, использующих контекст.  

В примере 3 мы демонстрируем, как работает очередь команд. При вызове context.sync() первым делом в Word выполняется [команда, загружающая](Word%20Add-ins%20JavaScript%20Reference/loadoption.md) основной текст. Затем — команда для вставки текста в тело документа Word. Результаты возвращаются в прокси-объект основного текста. Значение свойства body.text в JavaScript для Word равняется значению основного текста документа Word <u>перед</u> вставкой в него текста. 

**Пример 3. Выполнение пакета команд.**

```javascript
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        var body = context.document.body;

        // Queue a command to load the text in the proxy body object.
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

## Оставьте свой отзыв

Ваши отзывы важны для нас. 

* Ознакомьтесь с документами и сообщите нам о любых возникших вопросах и проблемах, с которыми вы столкнулись, [отправив сообщение](https://github.com/OfficeDev/office-js-docs/issues) в этом репозитории.
* Поделитесь своими впечатлениями о работе средств, расскажите, что вы бы хотели видеть в последующих версиях, какие примеры кода вас интересуют и т. д. Вы можете внести свои предложения и поделиться идеями на [этом сайте](http://officespdev.uservoice.com/).


## Дополнительные ресурсы

* [Надстройки Word](word-add-ins.md)
* [Справочник по JavaScript надстроек Word](word-add-ins-javascript-reference.md)
* [Надстройки Office](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Начало работы с надстройками Office](http://dev.office.com/getting-started/addins)
* &lt;a herf="https://github.com/OfficeDev?utf8=%E2%9C%93&amp;query=Word"&gt;Надстройки Word на веб-сайте GitHub&lt;/a&gt;
* [Обозреватель фрагментов кода для Word](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)

