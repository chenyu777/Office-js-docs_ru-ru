
# Метод ProjectDocument.setTaskFieldAsync (API JavaScript для Office 1.1)
Асинхронно задает значение указанного поля для указанной задачи.
 **Важно!** Этот API работает только в Project 2016 на настольных компьютерах с Windows.

|||
|:-----|:-----|
|**Ведущие приложения:**|Project|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Выделение|
|**Добавлен в версии**|1.1|

```js
Office.context.document.setTaskFieldAsync(taskId, fieldId, fieldValue[, options][, callback]);
```


## Параметры


_taskId_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;GUID задачи. Обязательный.<br/><br/>
_fieldId_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Идентификатор целевого поля в виде константы [ProjectTaskFields](../../reference/shared/projecttaskfields-enumeration.md) или соответствующего целого числа. Обязательный.<br/><br/>
_fieldValue_<br/>
&nbsp;&nbsp;&nbsp;&nbsp; **string**, **number**, **boolean** **object**. Обязательный.<br/><br/>
_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;[необязательный параметр](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):<br/><br/>

&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Тип: **array, boolean, null, number, object, string** или **undefined**<br/></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Пользовательский элемент любого типа, который возвращается в объекте [AsyncResult](../../reference/shared/asyncresult.md) без изменений. Необязательный.</br></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Например, вы можете передать аргумент _asyncContext_, используя формат `{asyncContext: 'Some text'}` или `{asyncContext: <object>}`.<br/><br/>
_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Тип: **function**<br/><br/>
&nbsp;&nbsp;&nbsp;&nbsp;Функция, вызываемая после получения результатов вызова метода, единственный параметр которой имеет тип [AsyncResult](../../reference/shared/asyncresult.md). Необязательный.
    

## Значение обратного вызова

При выполнении функция _callback_ получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью параметра функции обратного вызова.

Объект [AsyncResult](../../reference/shared/asyncresult.md), возвращаемый методом **setTaskFieldAsync**, содержит указанные ниже свойства.



|**Название**|**Описание**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Данные, передаваемые в необязательном параметре _asyncContext_ (если он использовался).|
|[error](../../reference/shared/asyncresult.error.md)|Сведения об ошибке, если свойство **status** имеет значение **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|Состояние **succeeded** или **failed** асинхронного вызова.|
|[value](../../reference/shared/asyncresult.value.md)|Этот метод не возвращает значение.|

## Заметки

Прежде всего вызовите метод [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) или [getTaskByIndexAsync](../../reference/shared/projectdocument.settaskfieldasync.md), чтобы получить GUID задачи, а затем передайте этот GUID в качестве аргумента _taskId_ в **setTaskFieldAsync**. При каждом асинхронном вызове можно обновить только одно поле для одной задачи.


## Пример

Ниже приведен пример кода, который вызывает метод [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) для получения GUID текущей выбранной задачи в представлении задач. Затем он задает два значения поля задач с помощью рекурсивного вызова метода **setTaskFieldAsync**.

Для используемого в примере метода [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) необходимо, чтобы представление задач (например, "Использование задач") было активным и чтобы эта задача была выбрана. Пример активации кнопки на основе активного типа представления см. в описании метода [addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md).

В данном примере подразумевается, что в вашей надстройке есть ссылка на библиотеку jQuery и что указанные ниже элементы управления страницы определены в теге div контента в тексте страницы.




```HTML
<input id="set-info" type="button" value="Set info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            
            // After the DOM is loaded, add-in-specific code can run.
            app.initialize();
            $('#set-info').click(setTaskInfo);
        });
    };

    // Get the GUID of the task, and then get the task fields.
    function setTaskInfo() {
        getTaskGuid().then(
            function (data) {
                setTaskFields(data);
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

    // Set the specified fields for the selected task.
    function setTaskFields(taskGuid) {
        var targetFields = [Office.ProjectTaskFields.Active, Office.ProjectTaskFields.Notes];
        var fieldValues = [true, 'Notes for the task.'];

        // Set the field value. If the call is successful, set the next field.
        for (var i = 0; i < targetFields.length; i++) {
            Office.context.document.setTaskFieldAsync(
                taskGuid,
                targetFields[i],
                fieldValues[i],
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        i++;
                    }
                    else {
                        onError(result.error);
                    }
                }
            );
        }
        $('#message').html('Field values set');
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
|**Минимальный уровень разрешений**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Область задач|
|**Библиотека**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Представлено|

## См. также



#### Другие ресурсы


[Метод getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedresourceasync.md)
[getTaskByIndexAsync](../../reference/shared/projectdocument.settaskfieldasync.md)
[Объект AsyncResult](../../reference/shared/asyncresult.md)
[Перечисление ProjectTaskFields](../../reference/shared/projecttaskfields-enumeration.md)
[Объект ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
