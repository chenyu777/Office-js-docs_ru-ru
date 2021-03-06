
# Событие ProjectDocument.TaskSelectionChanged
Происходит при изменении выбора задачи в активном проекте.

|||
|:-----|:-----|
|**Ведущие приложения:**|Project|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Выделение|
|**Добавлен в версии**|1.0|

```js
Office.EventType.TaskSelectionChanged
```


## Замечания

 Событие **TaskSelectionChanged** — это константа перечисления [EventType](../../reference/shared/eventtype-enumeration.md), которую можно использовать в методах [ProjectDocument.addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) и [ProjectDocument.removeHandlerAsync](../../reference/shared/projectdocument.removehandlerasync.md) для добавления или удаления обработчика событий.


## Пример

Ниже приведен пример кода для добавления обработчика события **TaskSelectionChanged**. Если в документе выбрана другая задача, он получает ее GUID.

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
                Office.EventType.TaskSelectionChanged,
                getTaskGuid);
            getTaskGuid();
        });
    };

    // Get the GUID of the selected task and display it in the add-in.
    function getTaskGuid() {
        Office.context.document.getSelectedTaskAsync(
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

Пример, где показано, как использовать обработчик события **TaskSelectionChanged** в надстройке Project, см. в статье [Создание первой надстройки области задач для Project 2013 с помощью текстового редактора](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает, что данное событие поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает это событие.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||||
|:-----|:-----|:-----|
||Office для рабочего стола Windows|Office Online (в браузере)|
|**Project**|Y||

|||
|:-----|:-----|
|**Доступен в наборах требований**|Выделение|
|**Минимальный уровень разрешений**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Область задач|
|**Библиотека**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



|**Версия**|**Изменения**|
|:-----|:-----|
|1.0|<ul><li>Представлено</li></ul>|

## См. также



#### Другие ресурсы


[Создание первой надстройки области задач для Project 2013 с помощью текстового редактора](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
[Перечисление EventType](../../reference/shared/eventtype-enumeration.md)
[Метод ProjectDocument.addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md)
[Метод ProjectDocument.removeHandlerAsync](../../reference/shared/projectdocument.removehandlerasync.md)
[Объект ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
