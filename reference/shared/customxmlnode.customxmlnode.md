
# Объект CustomXmlNode
Представляет XML-узел в дереве документа.

|||
|:-----|:-----|
|**Ведущие приложения:**|Word|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Последнее изменение в **|1.1|

```js
CustomXmlNode
```


## Элементы


**Свойства**


|**Имя**|**Описание**|
|:-----|:-----|
|[baseName](../../reference/shared/customxmlnode.basename.md)|Получает базовое имя узла без префикса пространства имен (если существует).|
|[nodeType](../../reference/shared/customxmlnode.nodetype.md)|Получает тип **CustomXMLNode**.|
|[namespaceUri](../../reference/shared/customxmlnode.namespaceuri.md)|Получает GUID строки **CustomXMLPart**.|

**Методы**


|**Имя**|**Описание**|
|:-----|:-----|
|[getNodesAsync](../../reference/shared/customxmlnode.getnodesasync.md)|Асинхронно получает узлы в виде массива объектов **CustomXMLNode**, соответствующих относительному выражению XPath.|
|[getNodeValueAsync](../../reference/shared/customxmlnode.getnodevalueasync.md)|Асинхронно получает значение узла.|
|[getTextAsync](customxmlnode.gettextasync.md)|Асинхронно получает текст узла XML в настраиваемой XML-части.|
|[getXmlAsync](../../reference/shared/customxmlnode.getxmlasync.md)|Асинхронно получает XML узла.|
|[setNodeValueAsync](../../reference/shared/customxmlnode.setnodevalueasync.md)|Асинхронно задает значение узла.|
|[setTextAsync](customxmlnode.settextasync.md)|Асинхронно задает текст узла XML в настраиваемой XML-части.|
|[setXmlAsync](../../reference/shared/customxmlnode.setxmlasync.md)|Асинхронно задает XML узла.|

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
