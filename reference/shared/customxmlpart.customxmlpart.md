
# Объект CustomXmlPart
Представляет один объект **CustomXMLPart** в коллекции [CustomXMLParts](../../reference/shared/customxmlparts.customxmlparts.md).

|||
|:-----|:-----|
|**Ведущие приложения:**|Word|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Последнее изменение в **|1.1|

```
Office.context.document.customXmlParts.getByIdAsync(id);
```


## Элементы


**Свойства**


|**Имя**|**Описание**|
|:-----|:-----|
|[builtIn](../../reference/shared/customxmlpart.builtin.md)|Получает значение, указывающее, является ли объект CustomXMLPart встроенным.|
|[id](../../reference/shared/customxmlpart.id.md)|Получает GUID объекта CustomXMLPart|
|[namespaceManager](../../reference/shared/customxmlpart.namespacemanager.md)|Получает набор сопоставлений префиксов пространства имен (CustomXMLPrefixMappings), используемых для текущего объекта CustomXMLPart.|

**Методы**


|**Имя**|**Описание**|
|:-----|:-----|
|[addHandlerAsync](../../reference/shared/customxmlpart.addhandlerasync.md)|Асинхронно добавляет обработчик событий для события объекта **CustomXmlPart**.|
|[deleteAsync](../../reference/shared/customxmlpart.deleteasync.md)|Асинхронно удаляет настраиваемую XML-часть из коллекции.|
|[getNodesAsync](../../reference/shared/customxmlpart.getnodesasync.md)|Асинхронно получает все объекты CustomXmlNode в настраиваемой XML-части, соответствующие указанному параметру XPath.|
|[getXmlAsync](../../reference/shared/customxmlpart.getxmlasync.md)|Асинхронно получает XML внутри настраиваемой XML-части.|
|[removeHandlerAsync](../../reference/shared/customxmlpart.removehandlerasync.md)|Удаляет обработчик события объекта **Document**.|

**События**


|**Имя**|**Описание**|
|:-----|:-----|
|[dataNodeDeleted](../../reference/shared/customxmlpart.datanodedeleted.event.md)|Происходит при удалении узла.|
|[dataNodeInserted](../../reference/shared/customxmlpart.datanodeinserted.event.md)|Происходит при вставке узла.|
|[dataNodeReplaced](../../reference/shared/customxmlpart.datanodereplaced.event.md)|Происходит при замене узла.|

## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|CustomXmlParts|
|**Минимальный уровень разрешений**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Область задач|
|**Библиотека**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



****


|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Word в Office для iPad.|
|1.0|Представлено|
