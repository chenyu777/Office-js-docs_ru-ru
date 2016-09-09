
# Метод ProjectDocument.getResourceFieldAsync
Асинхронно получает значение указанного поля для определенного ресурса в представлении ресурсов.

|||
|:-----|:-----|
|**Ведущие приложения:**|Project|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Выделение|
|**Добавлен в версии**|1.0|

```
Office.context.document.getResourceFieldAsync(resourceId, fieldId[, options][, callback]);
```


## Параметры



|**Имя**|**Тип**|**Описание**|**Примечания по вопросам поддержки**|
|:-----|:-----|:-----|:-----|
| _resourceId_|**строка**|GUID ресурса. Обязательный.||
| _fieldId_|[ProjectResourceFields](../../reference/shared/projectresourcefields-enumeration.md)|Идентификатор целевого поля. Обязательный параметр.||
| _options_|**object**|Задает следующие [необязательные параметры](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** или **undefined**|Определяемый пользователем элемент любого типа, который возвращается в объекте **AsyncResult** без изменения.||
| _callback_|**object**|Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.||

## Значение обратного вызова

При выполнении функция _callback_ получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью параметра функции обратного вызова.

В случае метода **getResourceFieldAsync** возвращенный объект [AsyncResult](../../reference/shared/asyncresult.md) содержит такие свойства:


****


|**Имя**|**Описание**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Данные, передаваемые в необязательном параметре _asyncContext_ (если он использовался).|
|[error](../../reference/shared/asyncresult.error.md)|Сведения об ошибке, если свойство **status** имеет значение **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|Состояние **succeeded** или **failed** асинхронного вызова.|
|[value](../../reference/shared/asyncresult.value.md)|Содержит свойство **fieldValue**, которое представляет значение указанного поля.|

## Заметки

Сначала вызывает метод [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md), чтобы получить GUID ресурса, а затем передает его в качестве аргумента _resourceId_ методу **getResourceFieldAsync**. Если активно не представление ресурсов (например, представление использования ресурсов или листа ресурсов) или если в представлении ресурсов не выбран ресурс, метод [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) возвращает ошибку 5001 (внутренняя ошибка). Пример использования события [ViewSelectionChanged](../../reference/shared/projectdocument.addhandlerasync.md) и метода [getSelectedViewAsync](../../reference/shared/projectdocument.viewselectionchanged.event.md) для активации кнопки с учетом типа активного представления см. в статье [Метод addHandlerAsync](../../reference/shared/projectdocument.getselectedviewasync.md).


## Пример

Ниже приведен пример кода, который вызывает метод [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) для получения GUID текущего выбранного ресурса в представлении ресурсов. Затем он получает три значения поля ресурсов с помощью рекурсивного вызова метода **getResourceFieldAsync**.

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
            $('#get-info').click(getResourceInfo);
        });
    };

    // Get the GUID of the resource and then get the resource fields.
    function getResourceInfo() {
        getResourceGuid().then(
            function (data) {
                getResourceFields(data);
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

    // Get the specified fields for the selected resource.
    function getResourceFields(resourceGuid) {
        var targetFields =
            [Office.ProjectResourceFields.Name, Office.ProjectResourceFields.Units, Office.ProjectResourceFields.BaseCalendar];
        var fieldValues = ['Name: ', 'Units: ', 'Base calendar: '];
        var index = 0; 
        getField();

        // Get each field, and then display the field values in the add-in.
        function getField() {
            if (index == targetFields.length) {
                var output = '';
                for (var i = 0; i < fieldValues.length; i++) {
                    output += fieldValues[i] + '<br />';
                }
                $('#message').html(output);
            }

            // If the call is successful, get the field value and then get the next field.
            else {
                Office.context.document.getResourceFieldAsync(
                    resourceGuid,
                    targetFields[index],
                    function (result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            fieldValues[index] += result.value.fieldValue;
                            getField(index++);
                        }
                        else {
                            onError(result.error);
                        }
                    }
                );
            }
        }
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


[Метод getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedresourceasync.md)

[Перечисление ProjectResourceFields](../../reference/shared/projectresourcefields-enumeration.md)

[Объект AsyncResult](../../reference/shared/asyncresult.md)

[Объект ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
