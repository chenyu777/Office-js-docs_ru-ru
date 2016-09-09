
# Событие Document.SelectionChanged
Происходит при изменении выбора в документе.

|||
|:-----|:-----|
|**Ведущие приложения:**|Excel, PowerPoint, Word|
|**Представлены в**|1.1|

```
Office.EventType.DocumentSelectionChanged
```

## Замечания

Чтобы добавить обработчик события документа **SelectionChanged**, используйте метод [addHandlerAsync](../../reference/shared/document.addhandlerasync.md) объекта **Document**.


## Пример




```
function addEventHandlerToDocument() {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, MyHandler);
}

function MyHandler(eventArgs) {
    doSomethingWithDocument(eventArgs.document);
}

```




## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Поддерживаемые ведущие приложения по платформе**


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Да|Y|
|**PowerPoint**|Y|Да|Y|
|**Word**|Y|Да|Y|

|||
|:-----|:-----|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



****


|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel, PowerPoint и Word в Office для iPad.|
|1.0|Представлено|
