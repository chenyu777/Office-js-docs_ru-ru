
# Метод ProjectDocument.getProjectFieldAsync
Асинхронно получает значение указанного поля в активном проекте.

|||
|:-----|:-----|
|**Ведущие приложения:**|Project|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Выделение|
|**Добавлен в версии**|1.0|

```
Office.context.document.getProjectFieldAsync(fieldId[, options][, callback]);
```


## Параметры



|**Имя**|**Тип**|**Описание**|
|:-----|:-----|:-----|:-----|
| _fieldId_|[ProjectProjectFields](../../reference/shared/projectprojectfields-enumeration.md)|Идентификатор целевого поля. Обязательный параметр.|
| _options_|**object**|Задает следующие [необязательные параметры](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):|
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** или **undefined**|Определяемый пользователем элемент любого типа, который возвращается в объекте **AsyncResult** без изменения.|
| _callback_|**object**|Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.|

## Значение обратного вызова

При выполнении функция _callback_ получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью параметра функции обратного вызова.

В случае метода **getProjectFieldAsync** возвращенный объект [AsyncResult](../../reference/shared/asyncresult.md) содержит такие свойства:


****


|**Имя**|**Описание**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Данные, передаваемые в необязательном параметре _asyncContext_ (если он использовался).|
|[error](../../reference/shared/asyncresult.error.md)|Сведения об ошибке, если свойство **status** имеет значение **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|Состояние **succeeded** или **failed** асинхронного вызова.|
|[value](../../reference/shared/asyncresult.value.md)|Содержит свойство **fieldValue**, которое представляет значение указанного поля.|

## Пример

Ниже приведен пример кода, который получает значения трех указанных полей активного проекта и отображает их в надстройке.

В примере осуществляется рекурсивный вызов метода **getProjectFieldAsync** после успешного возвращения предыдущего вызова. Кроме того, отслеживаются вызовы, чтобы определить момент выполнения всех вызовов.

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

            // Get information for the active project.
            getProjectInformation();
        });
    };

    // Get the specified fields for the active project.
    function getProjectInformation() {
        var fields =
            [Office.ProjectProjectFields.Start, Office.ProjectProjectFields.Finish, Office.ProjectProjectFields.GUID];
        var fieldValues = ['Start: ', 'Finish: ', 'GUID: '];
        var index = 0; 
        getField();

        // Get each field, and then display the field values in the add-in.
        function getField() {
            if (index == fields.length) {
                var output = '';
                for (var i = 0; i < fieldValues.length; i++) {
                    output += fieldValues[i] + '<br />';
                }
                $('#message').html(output);
            }
            else {
                Office.context.document.getProjectFieldAsync(
                    fields[index],
                    function (result) {

                        // If the call is successful, get the field value and then get the next field.
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
|**Доступен в наборах требований**||
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


[Перечисление ProjectProjectFields](../../reference/shared/projectprojectfields-enumeration.md)

[Объект AsyncResult](../../reference/shared/asyncresult.md)

[Объект ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
