

# Метод ProjectDocument.getWSSUrlAsync
Асинхронно получает URL-адрес синхронизированного списка задач SharePoint.

|||
|:-----|:-----|
|**Ведущие приложения:**|Project|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Выделение|
|**Добавлен в версии**|1.0|

```js
Office.context.document.getWSSUrlAsync([options,] [callback]);
```


## Параметры



|**Имя**|**Тип**|**Описание**|**Примечания по вопросам поддержки**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|Задает следующие [необязательные параметры](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** или **undefined**|Определяемый пользователем элемент любого типа, который возвращается в объекте **AsyncResult** без изменения.||
| _callback_|**object**|Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.||

## Значение обратного вызова

При выполнении функция _callback_ получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью параметра функции обратного вызова.

В случае метода **getWSSUrlAsync** возвращенный объект [AsyncResult](../../reference/shared/asyncresult.md) содержит такие свойства:


|**Название**|**Описание**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Данные, передаваемые в необязательном параметре _asyncContext_ (если он использовался).|
|[error](../../reference/shared/asyncresult.error.md)|Сведения об ошибке, если свойство **status** имеет значение **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|Состояние **succeeded** или **failed** асинхронного вызова.|
|[value](../../reference/shared/asyncresult.value.md)|Содержит следующие свойства:<br/><br/><ul><li>Свойство <b>listName</b> — название синхронизированного списка задач SharePoint.</li><li>Свойство <b>serverUrl</b> — URL-адрес синхронизированного списка задач SharePoint.</li></ul>|

## Замечания

Если активный проект не синхронизирован со списком задач SharePoint, значения **serverUrl** и **listName** будут пустыми.


## Пример

Ниже приведен пример кода, который вызывает метод **getWSSUrlAsync**, чтобы получить название и URL-адрес синхронизированного списка задач SharePoint.

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

            // After the DOM is loaded, add-in-specific code can run.
            getSharePointTaskListUrl();
        });
    };

    // Get the URL of the the synchronized SharePoint task list.
    function getSharePointTaskListUrl() {
        Office.context.document.getWSSUrlAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    var output = String.format(
                        'List name: {0}<br />List URL: {1}',
                        result.value.listName, result.value.serverUrl);
                    $('#message').html(output);
                }
                else {
                    onError(result.error);
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
|**Доступен в наборах требований**||
|**Минимальный уровень разрешений**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Область задач|
|**Библиотека**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки


|**Версия**|**Изменения**|
|:-----|:-----|
|1.0|Представлено|

## См. также



#### Другие ресурсы


[Объект AsyncResult](../../reference/shared/asyncresult.md)
[Объект ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
