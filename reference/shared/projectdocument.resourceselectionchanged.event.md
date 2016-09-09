

# Событие ProjectDocument.ResourceSelectionChanged
Возникает при изменении выбора ресурсов в активном проекте.

|||
|:-----|:-----|
|**Ведущие приложения:**|Project|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Выделение|
|**Добавлен в версии**|1.0|

```js
Office.EventType.ResourceSelectionChanged
```


## Замечания

 Событие **ResourceSelectionChanged** — константа перечисления [EventType](../../reference/shared/eventtype-enumeration.md), которую можно использовать в методах [ProjectDocument.addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) и [ProjectDocument.removeHandlerAsync](../../reference/shared/projectdocument.removehandlerasync.md) для добавления или удаления обработчика событий.


## Пример

Ниже показан пример кода, который добавляет обработчик события **ResourceSelectionChanged**. Если в документе выбран другой ресурс, он получает его GUID.

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
            Office.context.document.addHandlerAsync(
                Office.EventType.ResourceSelectionChanged,
                getResourceGuid);
        });
    };

    // Get the GUID of the selected resource and display it in the add-in.
    function getResourceGuid() {
        Office.context.document.getSelectedResourceAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    $('#message').html(result.value);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```

Полный пример кода, где показано, как использовать обработчик события **ResourceSelectionChanged** в надстройке Project, см. в статье [Создание первой надстройки области задач для Project 2013 с помощью текстового редактора](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает, что данное событие поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает это событие.

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

|**Версия**|**Изменения**|
|:-----|:-----|
|1.0|Представлено|

## См. также



#### Другие ресурсы


[Создание первой надстройки области задач для Project 2013 с помощью текстового редактора](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
[Перечисление EventType](../../reference/shared/eventtype-enumeration.md)
[Метод ProjectDocument.addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md)
[Метод ProjectDocument.removeHandlerAsync](../../reference/shared/projectdocument.removehandlerasync.md)
[Объект ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
