
# Метод ProjectDocument.getMaxTaskIndexAsync
Асинхронно получает максимальный индекс коллекции задач в текущем проекте.

 **Важно!** Этот API работает только в Project 2016 на настольных компьютерах с Windows.

|||
|:-----|:-----|
|**Ведущие приложения:**|Project|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Выделение|
|**Добавлен в версии**|1.1|

```js
Office.context.document.getMaxTaskIndexAsync([options][, callback]);
```


## Параметры

_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Следующий **[необязательный параметр](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):**<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Тип: **array**, **boolean**, **null**, **number**, **object**, **string** **undefined**<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Пользовательский элемент любого типа, который возвращается в объекте [AsyncResult](../../reference/shared/asyncresult.md) без изменений. Необязательный.<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Например, вы можете передать аргумент _asyncContext_, используя формат `{asyncContext: 'Some text'}` или `{asyncContext: <object>}`.

_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Тип: **function**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Функция, вызываемая после получения результатов вызова метода, единственный параметр которой имеет тип [AsyncResult](../../reference/shared/asyncresult.md). Необязательный.   

## Значение обратного вызова

При выполнении функция _callback_ получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью параметра функции обратного вызова.

Объект [AsyncResult](../../reference/shared/asyncresult.md), возвращаемый методом **getMaxTaskIndexAsync**, содержит указанные ниже свойства.


|**Название**|**Описание**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Данные, передаваемые в необязательном параметре _asyncContext_ (если он использовался).|
|[error](../../reference/shared/asyncresult.error.md)|Сведения об ошибке, если свойство **status** имеет значение **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|Состояние **succeeded** или **failed** асинхронного вызова.|
|[value](../../reference/shared/asyncresult.value.md)|Максимальный номер индекса в коллекции задач текущего проекта.|

## Примечания

Возвращенное значение можно использовать с методом [getTaskByIndexAsync](../../reference/shared/projectdocument.gettaskbyindexasync.md), чтобы получить идентификаторы GUID задач. Задача с индексом 0 представляет суммарную задачу проекта.


## Пример

В примере кода ниже показано, как вызвать метод **getMaxTaskIndexAsync**, чтобы получить максимальный индекс коллекции задач в текущем проекте. Затем он использует возвращенное значение и метод [getTaskByIndexAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) для получения GUID каждой задачи.

В данном примере подразумевается, что в вашей надстройке есть ссылка на библиотеку jQuery и что указанные ниже элементы управления страницы определены в теге div контента в тексте страницы.




```HTML
<input id="get-info" type="button" value="Get info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";
    var taskGuids = [];

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            app.initialize();
            $('#get-info').click(getTaskInfo);
        });
    };

    // Get the maximum task index, and then get the task GUIDs.
    function getTaskInfo() {
        getMaxTaskIndex().then(
            function (data) {
                getTaskGuids(data);
            }
        );
    }

    // Get the maximum index of the tasks for the current project.
    function getMaxTaskIndex() {
        var defer = $.Deferred();
        Office.context.document.getMaxTaskIndexAsync(
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

    // Get each task GUID, and then display the GUIDs in the add-in.
    function getTaskGuids(maxTaskIndex) {
        var defer = $.Deferred();
        for (var i = 0; i <= maxTaskIndex; i++) {
            getTaskGuid(i);
        }
        return defer.promise();
        function getTaskGuid(index) {
            Office.context.document.getTaskByIndexAsync(index,
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        taskGuids.push(result.value);
                        if (index == maxTaskIndex) {
                            defer.resolve();
                            $('#message').html(taskGuids.toString());
                        }
                    }
                    else {
                        onError(result.error);
                    }
                }
            );
        }
    }
    function onError(error) {
        app.showNotification(error.name + ' ' + error.code + ': ' + error.message);
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
|**Доступен в наборах требований**||
|**Минимальный уровень разрешений**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Область задач|
|**Библиотека**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки




|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Представлено|

## См. также



#### Другие ресурсы


[getTaskByIndexAsync](../../reference/shared/projectdocument.gettaskbyindexasync.md)

[Объект AsyncResult](../../reference/shared/asyncresult.md)

[Объект ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
