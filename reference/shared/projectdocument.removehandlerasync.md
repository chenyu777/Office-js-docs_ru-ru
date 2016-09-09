

# Метод ProjectDocument.removeHandlerAsync
Асинхронно удаляет обработчик события изменения выбора задачи в объекте [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md).

|||
|:-----|:-----|
|**Ведущие приложения:**|Project|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Выделение|
|**Добавлен в версии**|1.0|

```js
Office.context.document.removeHandlerAsync(eventType[, options][, callback]);
```


## Параметры
|**Имя**|**Тип**|**Описание**|**Примечания по вопросам поддержки**|
|:-----|:-----|:-----|:-----|
|_eventType_|[EventType](../../reference/shared/eventtype-enumeration.md)|Тип удаляемого события в виде константы [EventType](../../reference/shared/eventtype-enumeration.md) или соответствующего текстового значения. Обязательный.<br/><br/>В следующей таблице показаны допустимые аргументы eventType для объекта [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md).<br/><br/><table><tr><th>Перечисление</th><th>Текстовое значение</th></tr><tr><td><a href="https://msdn.microsoft.com/en-us/library/office/fp179836.aspx">Office.EventType.ResourceSelectionChanged</a></td><td>resourceSelectionChanged</td></tr><tr><td><a href="https://msdn.microsoft.com/en-us/library/office/fp179816.aspx">Office.EventType.TaskSelectionChanged</a></td><td>taskSelectionChanged</td></tr><tr><td><a href="https://msdn.microsoft.com/en-us/library/office/fp179839.aspx">Office.EventType.ViewSelectionChanged</a></td><td>viewSelectionChanged</td></tr></table>||
|_options_|**object**|Задает следующие [необязательные параметры](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
|_asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** или **undefined**|Определяемый пользователем элемент любого типа, который возвращается в объекте **AsyncResult** без изменения.||
|_callback_|**object**|Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.||


## Значение обратного вызова

При выполнении функция _callback_ получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью параметра функции обратного вызова.

В случае метода **removeHandlerAsync** возвращенный объект [AsyncResult](../../reference/shared/asyncresult.md) содержит следующие свойства:


|**Название**|**Описание**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Данные, передаваемые в необязательном параметре _asyncContext_ (если он использовался).|
|[error](../../reference/shared/asyncresult.error.md)|Сведения об ошибке, если свойство **status** имеет значение **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|Состояние **succeeded** или **failed** асинхронного вызова.|
|[value](../../reference/shared/asyncresult.value.md)|**removeHandlerAsync** всегда возвращает значение **undefined**.|

## Пример

В приведенном ниже примере кода используются метод [addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md), чтобы добавить обработчик события [ResourceSelectionChanged](../../reference/shared/projectdocument.resourceselectionchanged.event.md), и метод **removeHandlerAsync**, чтобы удалить этот обработчик.

При выборе ресурса в представлении ресурсов его GUID отображается в обработчике. Но если обработчик удалить, то GUID отображаться не будет.

В примере предполагается, что в надстройке имеется ссылка на библиотеку jQuery и в разделителе контента страницы определен такой элемент управления страницей:




```HTML
<input id="remove-handler" type="button" value="Remove handler" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            Office.context.document.addHandlerAsync(
                Office.EventType.ResourceSelectionChanged,
                getResourceGuid);
            $('#remove-handler').click(removeEventHandler);
        });
    };

    // Remove the event handler.
    function removeEventHandler() {
        Office.context.document.removeHandlerAsync(
            Office.EventType.ResourceSelectionChanged,
            {handler:getResourceGuid,
            asyncContext:'The handler is removed.'},
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    $('#remove-handler').attr('disabled', 'disabled');
                    $('#message').html(result.asyncContext);
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


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**Доступен в наборах требований**|Выделение|
|**Минимальный уровень разрешений**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Область задач|
|**Библиотека**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки

|**Версия**|**Изменения**|
|:-----|:-----|
|1.0|Представлено|

## См. также



#### Другие ресурсы


[Метод addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md)
[Перечисление EventType](../../reference/shared/eventtype-enumeration.md)
[Объект ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)

