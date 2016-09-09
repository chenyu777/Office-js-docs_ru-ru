
# Метод ProjectDocument.addHandlerAsync
Асинхронно добавляет обработчик события изменения в объект [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md).

|||
|:-----|:-----|
|**Ведущие приложения:**|Project|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Выделение|
|**Добавлен в версии**|1.0|

```
Office.context.document.addHandlerAsync(eventType, handler[, options][, callback]);
```


## Параметры



|**Имя**|**Тип**|**Описание**|
|:-----|:-----|:-----|
| _eventType_|[EventType](../../reference/shared/eventtype-enumeration.md)|Тип добавляемого события в виде константы [EventType](../../reference/shared/eventtype-enumeration.md) или соответствующего текстового значения. Обязательный. В следующей таблице показаны допустимые аргументы _eventType_ для объекта [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md).<table><tr><td>**Перечисление**</td><td>**Текстовое значение**</td></tr><tr><td>[Office.EventType.ResourceSelectionChanged](../../reference/shared/projectdocument.resourceselectionchanged.event.md)</td><td>resourceSelectionChanged</td></tr><tr><td>[Office.EventType.TaskSelectionChanged](../../reference/shared/projectdocument.taskselectionchanged.event.md)</td><td>taskSelectionChanged</td></tr><tr><td>[Office.EventType.ViewSelectionChanged](../../reference/shared/projectdocument.viewselectionchanged.event.md)</td><td>viewSelectionChanged</td></tr></table>|
| _handler_|**функция**|Имя обработчика событий. Обязательный.|
| _options_|**object**|Задает следующие [необязательные параметры](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):|
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** или **undefined**|Определяемый пользователем элемент любого типа, который возвращается в объекте **AsyncResult** без изменения.|
| _callback_|**object**|Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.|

## Значение обратного вызова

При выполнении функция _callback_ получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью параметра функции обратного вызова.

В случае метода **addHandlerAsync** возвращенный объект [AsyncResult](../../reference/shared/asyncresult.md) содержит такие свойства:


****


|**Имя**|**Описание**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Данные, передаваемые в необязательном параметре _asyncContext_ (если он использовался).|
|[error](../../reference/shared/asyncresult.error.md)|Сведения об ошибке, если свойство **status** имеет значение **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|Состояние **succeeded** или **failed** асинхронного вызова.|
|[value](../../reference/shared/asyncresult.value.md)|**addHandlerAsync** всегда возвращает значение **undefined**.|

## Пример

В приведенном ниже примере кода используется метод **addHandlerAsync**, чтобы добавить обработчик события [ViewSelectionChanged](../../reference/shared/projectdocument.viewselectionchanged.event.md).

При изменении активного представления обработчик проверяет его тип. Если выбрано представление ресурсов, он активирует кнопку, а в противном случае — отключает. По нажатию кнопки считывается GUID выбранного ресурса, который затем отображается в надстройке.

В данном примере подразумевается, что в вашей надстройке есть ссылка на библиотеку jQuery и что указанные ниже элементы управления страницы определены в теге div контента в тексте страницы.




```HTML
<input id="get-info" type="button" value="Get info" disabled="disabled" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            // Add a ViewSelectionChanged event handler.
            Office.context.document.addHandlerAsync(
                Office.EventType.ViewSelectionChanged,
                getActiveView);
            $('#get-info').click(getResourceGuid);

            // This example calls the handler on page load to get the active view
            // of the default page.
            getActiveView();
        });
    };

    // Activate the button based on the active view type of the document.
    // This is the ViewSelectionChanged event handler.
    function getActiveView() {
        Office.context.document.getSelectedViewAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var viewType = result.value.viewType;
                    if (viewType == 6 ||   // ResourceForm
                        viewType == 7 ||   // ResourceSheet
                        viewType == 8 ||   // ResourceGraph
                        viewType == 15) {  // ResourceUsage
                        $('#get-info').removeAttr('disabled');
                    }
                    else {
                        $('#get-info').attr('disabled', 'disabled');
                    }
                    var output = String.format(
                        'View name: {0}<br/>View type: {1}',
                        result.value.viewName, viewType);
                    $('#message').html(output);
                }
            }
        );
    }

    // Get the GUID of the currently selected resource and display it in the add-in.
    function getResourceGuid() {
        Office.context.document.getSelectedResourceAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    $('#message').html('Resource GUID: ' + result.value);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```

Полный пример кода, где показано, как использовать обработчик события [TaskSelectionChanged](../../reference/shared/projectdocument.taskselectionchanged.event.md) в надстройке Project, см. в статье [Создание первой надстройки области задач для Project с помощью текстового редактора](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**Доступен в наборах требований**||
|**Минимальный уровень разрешений**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Область задач|
|**Библиотека**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



****


|**Версия**|**Изменения**|
|:-----|:-----|
|1.0|Представлено|

## См. также



#### Другие ресурсы


[событие TaskSelectionChanged](../../reference/shared/projectdocument.taskselectionchanged.event.md)

[Метод removeHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md)

[Объект ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
