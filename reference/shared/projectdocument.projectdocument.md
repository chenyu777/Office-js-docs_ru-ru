

# Объект ProjectDocument
Абстрактный класс, представляющий документ проекта (активный проект), с которым взаимодействует надстройка Office.

|||
|:-----|:-----|
|**Ведущие приложения:**|Project|
|**Добавлено в версии**|1.0|

```js
Office.context.document
```


## Элементы


**Методы**


|**Имя**|**Описание**|
|:-----|:-----|
|[Метод addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md)|Асинхронно добавляет обработчик события в объекте **ProjectDocument**.|
|[Метод getMaxResourceIndexAsync](../../reference/shared/projectdocument.getmaxresourceindexasync.md)|Асинхронно получает максимальный индекс коллекции ресурсов в текущем проекте.|
|[Метод getMaxTaskIndexAsync](../../reference/shared/projectdocument.getmaxtaskindexasync.md)|Асинхронно получает максимальный индекс коллекции задач в текущем проекте.|
|[Метод getProjectFieldAsync](../../reference/shared/projectdocument.getprojectfieldasync.md)|Асинхронно получает значение указанного поля в активном проекте.|
|[Метод getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md)|Асинхронно получает GUID ресурса с указанным индексом в коллекции ресурсов.|
|[Метод getResourceFieldAsync](../../reference/shared/projectdocument.getresourcefieldasync.md)|Асинхронно получает значение указанного поля для заданного ресурса.|
|[Метод getSelectedDataAsync](../../reference/shared/projectdocument.getselecteddataasync.md)|Асинхронно получает данные, содержащиеся в текущем выборе одной или нескольких секций на диаграмме Ганта.|
|[Метод getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedresourceasync.md)|Асинхронно получает GUID выбранного ресурса.|
|[Метод getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md)|Асинхронно получает GUID выбранной задачи.|
|[Метод getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)|Асинхронно получает тип и имя активного представления.|
|[Метод getTaskAsync](../../reference/shared/projectdocument.gettaskasync.md)|Асинхронно получает имя задачи, ресурсы, назначенные задаче, и идентификатор задачи в синхронизированном списке задач SharePoint.|
|[Метод getTaskByIndexAsync](../../reference/shared/projectdocument.gettaskbyindexasync.md)|Асинхронно получает GUID задачи с указанным индексом в коллекции задач.|
|[Метод getTaskFieldAsync](../../reference/shared/projectdocument.gettaskfieldasync.md)|Асинхронно получает значение указанного поля для указанной задачи.|
|[Метод getWSSUrlAsync](../../reference/shared/projectdocument.getwssurlasync.md)|Асинхронно получает URL-адрес синхронизированного списка задач SharePoint.|
|[Метод removeHandlerAsync](../../reference/shared/projectdocument.removehandlerasync.md)|Асинхронно удаляет обработчик события в объекте **ProjectDocument**.|
|[Метод setResourceFieldAsync](../../reference/shared/projectdocument.setresourcefieldasync.md)|Асинхронно задает значение указанного поля для заданного ресурса.|
|[Метод setTaskFieldAsync](../../reference/shared/projectdocument.settaskfieldasync.md)|Асинхронно задает значение указанного поля для указанной задачи.|

**События**


|**Имя**|**Описание**|
|:-----|:-----|
|[событие ResourceSelectionChanged](../../reference/shared/projectdocument.resourceselectionchanged.event.md)|Возникает при изменении выбора ресурсов в активном проекте.|
|[событие TaskSelectionChanged](../../reference/shared/projectdocument.taskselectionchanged.event.md)|Происходит при изменении выбора задачи в активном проекте.|
|[событие ViewSelectionChanged](../../reference/shared/projectdocument.viewselectionchanged.event.md)|Возникает при изменении активного представления в активном проекте.|

## Заметки

Не вызывайте объект **ProjectDocument** напрямую и не создавайте его экземпляр к своем сценарии.


## Пример

В следующем примере надстройка инициализируется, а затем считывает свойства объекта [Document](../../reference/shared/document.md), доступные в контексте документа Project. Документ Project — это открытый, активный проект. Для доступа к членам объекта **ProjectDocument** используйте объект **Office.context.document**, как показано в примерах кода для методов и событий **ProjectDocument**.

В примере предполагается, что в надстройке имеется ссылка на библиотеку jQuery и в разделителе контента страницы определен такой элемент управления страницей:




```HTML
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // Get information about the document.
            showDocumentProperties();
        });
    };

    // Get the document mode and the URL of the active project.
    function showDocumentProperties() {
        var output = String.format(
            'The document mode is {0}.<br/>The URL of the active project is {1}.',
            Office.context.document.mode,
            Office.context.document.url);
        $('#message').html(output);
    }
})();
```


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает, что этот объект поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот объект.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**Типы надстроек**|Область задач|
|**Библиотека**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки


|**Версия**|**Изменения**|
|:-----|:-----|
|1.0|Представлено|

## См. также



#### Другие ресурсы


[Надстройки области задач для Project](../../docs/project/project-add-ins.md)
[Объект Document](../../reference/shared/document.md)

