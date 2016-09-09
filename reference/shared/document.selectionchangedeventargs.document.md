
# Свойство DocumentSelectionChangedEventArgs.document
Получает объект **Document**, представляющий документ, который вызвал событие **SelectionChanged**.

|||
|:-----|:-----|
|**Ведущие приложения:**|Excel, Word|
|**Добавлено в версии**|1.1|




```js
var myDoc = eventArgsObj.document;
```


## Возвращаемое значение

Объект [Document](../../reference/shared/document.md), представляющий документ, вызвавший событие [SelectionChanged](../../reference/shared/document.selectionchanged.event.md).


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Поддерживаемые ведущие приложения по платформе**


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Да|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Минимальный уровень разрешений**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



****


|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel, PowerPoint и Word в Office для iPad.|
|1.0|Представлено|
