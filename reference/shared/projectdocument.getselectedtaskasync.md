
# Метод ProjectDocument.getSelectedTaskAsync
Асинхронно получает GUID выбранной задачи в представлении задач.

|||
|:-----|:-----|
|**Ведущие приложения:**|Project|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Выделение|
|**Добавлен в версии**|1.0|

```
Office.context.document.getSelectedTaskAsync([options,] [callback]);
```


## Параметры



|**Имя**|**Тип**|**Описание**|**Примечания по вопросам поддержки**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|Задает следующие [необязательные параметры](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** или **undefined**|Определяемый пользователем элемент любого типа, который возвращается в объекте **AsyncResult** без изменения.||
| _callback_|**object**|Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.||

## Значение обратного вызова

При выполнении функция _callback_ получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью параметра функции обратного вызова.

В случае метода **getSelectedTaskAsync** возвращенный объект [AsyncResult](../../reference/shared/asyncresult.md) содержит следующие свойства:


****


|**Название**|**Описание**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Данные, передаваемые в необязательном параметре _asyncContext_ (если он использовался).|
|[error](../../reference/shared/asyncresult.error.md)|Сведения об ошибке, если свойство **status** имеет значение **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|Состояние **succeeded** или **failed** асинхронного вызова.|
|[value](../../reference/shared/asyncresult.value.md)|GUID выбранной задачи в виде **string**.|

## Заметки

В надстройках Project удобнее использовать GUID задачи, а не номер ее идентификатора (например, идентификатор первой задачи в диаграмме Ганта — **1**). С помощью GUID задачи можно получить доступ к сведениям о задачах Project, например задач в проекте SharePoint, который синхронизирован с Project Server в режиме видимости. GUID задачи можно также сохранить в локальной переменной и использовать ее для методов [getTaskAsync](../../reference/shared/projectdocument.gettaskasync.md) и [getTaskFieldAsync](../../reference/shared/projectdocument.gettaskfieldasync.md).

Метод **getSelectedTaskAsync** возвращает ошибку 5001 (внутренняя ошибка), если активно не представление задач, а другое (например, представление диаграммы Ганта или использования задач), или если в представлении задач не выбрана задача. Пример использования события [ViewSelectionChanged](../../reference/shared/projectdocument.addhandlerasync.md) и метода [getSelectedViewAsync](../../reference/shared/projectdocument.viewselectionchanged.event.md) для активации кнопки с учетом типа активного представления см. в статье [Метод addHandlerAsync](../../reference/shared/projectdocument.getselectedviewasync.md).


## Пример

Ниже приведен пример кода, который вызывает метод **getSelectedTaskAsync** для получения GUID выбранной задачи в представлении задач. Затем он получает ее свойства с помощью вызова метода [getTaskAsync](../../reference/shared/projectdocument.gettaskasync.md).

В данном примере подразумевается, что в вашей надстройке есть ссылка на библиотеку jQuery и что указанные ниже элементы управления страницы определены в теге div контента в тексте страницы.




```HTML
<input id="get-info" type="button" value="Get info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            $('#get-info').click(getTaskInfo);
        });
    };

    // // Get the GUID of the task, and then get local task properties.
    function getTaskInfo() {
        getTaskGuid().then(
            function (data) {
                getTaskProperties(data);
            }
        );
    }

    // Get the GUID of the selected task.
    function getTaskGuid() {
        var defer = $.Deferred();
        Office.context.document.getSelectedTaskAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    defer.resolve(result.value);
                }
            }
        );
        return defer.promise();
    }

    // Get local properties for the selected task, and then display it in the add-in.
    function getTaskProperties(taskGuid) {
        Office.context.document.getTaskAsync(
            taskGuid,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var taskInfo = result.value;
                    var output = String.format(
                        'Name: {0}<br/>GUID: {1}<br/>SharePoint task ID: {2}<br/>Resource names: {3}',
                        taskInfo.taskName, taskGuid, taskInfo.wssTaskId, taskInfo.resourceNames);
                    $('#message').html(output);
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
|**Минимальный уровень разрешений**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
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


[Метод getTaskAsync](../../reference/shared/projectdocument.gettaskasync.md)

[Объект AsyncResult](../../reference/shared/asyncresult.md)

[Объект ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
