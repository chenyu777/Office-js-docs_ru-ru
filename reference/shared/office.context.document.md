
# Свойство Context.document
Получает объект, представляющий документ, с которым взаимодействует надстройка.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, PowerPoint, Project, Word|
|**Последнее изменение в **|1.1|

```js
var _document = Office.context.document;
```


## Возвращаемое значение

Объект [Document](../../reference/shared/document.md).


## Заметки

Надстройка может использовать свойство **document**, чтобы получить доступ к API для взаимодействия с содержимым документов, книг, презентаций, проектов и баз данных (в веб-приложениях Access).


## Пример




```js
// Extension initialization code.
var _document;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, code specific to the add-in can run.
    // Initialize instance variables to access API objects.
    _document = Office.context.document;
    });
}

```


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что это свойство поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает это свойство.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Да|Y|
|**PowerPoint**|Y|Да|Y|
|**Project**|Y|||
|**Word**|Y|Да|Y|

|||
|:-----|:-----|
|**Минимальный уровень разрешений**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки




|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel, PowerPoint и Word в Office для iPad.|
|1.1|Для свойства **Office.context.document** добавлена поддержка доступа к базе данных в контентных надстройках для Access.|
|1.0|Представлено|
