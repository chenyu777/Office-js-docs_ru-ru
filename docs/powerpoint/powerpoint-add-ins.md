
# Создание контентных надстроек и надстроек области задач для PowerPoint

Примеры кода в этой статье иллюстрируют некоторые основные задачи по разработке контентных надстроек PowerPoint. Для отображения информации эти примеры используют функцию  `app.showNotification`, включенную в шаблоны проектов Надстройки Office в Visual Studio. Если вы не используете Visual Studio для разработки надстройки, функцию  `showNotification` потребуется заменить собственным кодом. Некоторые из этих примеров также зависят от объекта `globals`, который объявляется вне следующих функций:  `var globals = {activeViewHandler:0, firstSlideId:0};`

Эти примеры кода требуют, чтобы проект [ссылался на библиотеку Office.js 1.1 или более поздней версии](../../docs/develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).


## Определение активного представления презентации и обработка события ActiveViewChanged

Функция  `getFileView` вызывает метод [Document.getActiveViewAsync](../../reference/shared/document.getactiveviewasync.md), который возвращает текущее представление презентации: "edit" (представления, в которых можно редактировать слайды, например  **Обычный режим** или **Режим структуры**) или "read" (**Показ слайдов** или **Режим чтения**).


```js
function getFileView() {
    //Gets whether the current view is edit or read.
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification(asyncResult.value);
        }
    });
}
```

Функция `registerActiveViewChanged` вызывает метод [addHandlerAsync](../../reference/shared/document.addhandlerasync.md), чтобы зарегистрировать обработчик события [Document.ActiveViewChanged](../../reference/shared/document.activeviewchanged.md). Если изменить представление презентации после выполнения этой функции, появится уведомление  `app.showNotification` с активным режимом просмотра ("read" или "edit").




```js
function registerActiveViewChanged() {
    Globals.activeViewHandler = function (args) {
        app.showNotification(JSON.stringify(args));
    }

    Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, Globals.activeViewHandler, 
        function (asyncResult) {
            if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
            else {
            app.showNotification(asyncResult.status);
            }
        });
}
```


## Получение URL-адреса презентации

Функция `getFileUrl` вызывает метод [Document.getFileProperties](../../reference/shared/document.getfilepropertiesasync.md), чтобы получить URL-адрес файла презентации.


```js
function getFileUrl() {
    //Get the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;
        if (fileUrl == "") {
            app.showNotification("The file hasn't been saved yet. Save the file and try again");
        }
        else {
            app.showNotification(fileUrl);
        }
    });
}
```


## Переход к определенному слайду презентации

Функция  `getSelectedRange` вызывает метод [Document.getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md), чтобы получить объект JSON, возвращаемый свойством  `asyncResult.value` и который включает в себя массив с именем slides, содержащий идентификаторы, заголовки и индексы выбранного диапазона слайдов (или текущего слайда). Кроме того, он сохраняет идентификатор первого слайда в выбранном диапазоне в глобальной переменной.


```js
function getSelectedRange() {
    // Get the id, title, and index of the current slide (or selected slides) and store the first slide id */
    Globals.firstSlideId = 0;

    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            Globals.firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
}
```

Функция  `goToFirstSlide` вызывает метод [Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md) для перехода к идентификатору первого слайда, сохраненному описанной выше функцией `getSelectedRange`.




```js
function goToFirstSlide() {
    Office.context.document.goToByIdAsync(Globals.firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```


## Переход между слайдами презентации

Функция  `goToSlideByIndex` вызывает метод **Document.goToByIdAsync** для перехода к следующему слайду в презентации.


```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```




## Дополнительные ресурсы

- [Сохранение состояния надстройки и параметров документа для контентных надстроек и надстроек области задач](../../docs/develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)

- [Чтение и запись данных при активном выделении фрагмента в документе или электронной таблице](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
- [Получение всего документа из надстройки PowerPoint или Word](../../docs/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md)
    
- [Использование тем документов в надстройках PowerPoint](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md)
    
