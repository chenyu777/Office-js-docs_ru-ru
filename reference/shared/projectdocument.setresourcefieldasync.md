

# Метод ProjectDocument.setResourceFieldAsync
Асинхронно задает значение указанного поля для заданного ресурса.
 **Важно!** Этот API работает только в Project 2016 на настольных компьютерах с Windows.

|||
|:-----|:-----|
|**Ведущие приложения:**|Project|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Выделение|
|**Добавлен в версии**|1.1|

```js
Office.context.document.setResourceFieldAsync(resourceId, fieldId, fieldValue[, options][, callback]);
```


## Параметры

_resourceId_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;GUID ресурса. Обязательный.
    
_fieldId_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Идентификатор целевого поля в виде константы [ProjectResourceFields](../../reference/shared/projectresourcefields-enumeration.md) или соответствующего целого числа. Обязательный.
    
_fieldValue_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;в **string**, **number**, **boolean** или **object**. Обязательный.
    
_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;[необязательный параметр](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):

&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Тип: **array, boolean, null, number, object, string** или **undefined**<br/></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Пользовательский элемент любого типа, который возвращается в объекте [AsyncResult](../../reference/shared/asyncresult.md) без изменений. Необязательный.</br></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Например, вы можете передать аргумент _asyncContext_, используя формат `{asyncContext: 'Some text'}` или `{asyncContext: <object>}`.


_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Тип: **function**

&nbsp;&nbsp;&nbsp;&nbsp;Функция, вызываемая после получения результатов вызова метода, единственный параметр которой имеет тип [AsyncResult](../../reference/shared/asyncresult.md). Необязательный.

    

## Значение обратного вызова

При выполнении функция _callback_ получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью параметра функции обратного вызова.

Объект [AsyncResult](../../reference/shared/asyncresult.md), возвращаемый методом **setResourceFieldAsync**, содержит указанные ниже свойства.


|**Название**|**Описание**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Данные, передаваемые в необязательном параметре _asyncContext_ (если он использовался).|
|[error](../../reference/shared/asyncresult.error.md)|Сведения об ошибке, если свойство **status** имеет значение **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|Состояние **succeeded** или **failed** асинхронного вызова.|
|[value](../../reference/shared/asyncresult.value.md)|Этот метод не возвращает значение.|

## Заметки

Прежде всего вызовите метод [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) или [getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md), чтобы получить GUID ресурса, а затем передайте этот GUID в качестве аргумента _resourceId_ в метод **setResourceFieldAsync**. При каждом асинхронном вызове можно обновить только одно поле для одного ресурса.


## Пример

В примере кода ниже показано, как вызвать метод [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) для получения GUID текущего выбранного ресурса в представлении ресурсов. Затем код задает два значения поля ресурсов с помощью рекурсивного вызова метода **setResourceFieldAsync**.

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
            $('#set-info').click(setResourceInfo);
        });
    };

    // Get the GUID of the resource, and then get the resource fields.
    function setResourceInfo() {
        getResourceGuid().then(
            function (data) {
                setResourceFields(data);
            }
        );
    }

    // Get the GUID of the selected resource.
    function getResourceGuid() {
        var defer = $.Deferred();
        Office.context.document.getSelectedResourceAsync(
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

    // Set the specified fields for the selected resource.
    function setResourceFields(resourceGuid) {
        var targetFields = [Office.ProjectResourceFields.StandardRate, Office.ProjectResourceFields.Notes];
        var fieldValues = [.28, 'Notes for the resource.'];

        // Set the field value. If the call is successful, set the next field.
        for (var i = 0; i < targetFields.length; i++) {
            Office.context.document.setResourceFieldAsync(
                resourceGuid,
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


[getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md)
[getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md)
[Объект AsyncResult](../../reference/shared/asyncresult.md)
[Перечисление ProjectResourceFields](../../reference/shared/projectresourcefields-enumeration.md)
[Объект ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)

