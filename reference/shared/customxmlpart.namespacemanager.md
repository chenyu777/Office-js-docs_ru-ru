
# Свойство CustomXmlPart.namespaceManager
Получает набор сопоставлений префиксов пространства имен (CustomXMLPrefixMappings), используемых для текущего объекта CustomXMLPart.

|||
|:-----|:-----|
|**Ведущие приложения:**|Word|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Последнее изменение в **|1.1|

```js
var nsManager = customXmlPart.namespaceManager);
```


## Возвращаемое значение

Объект [CustomXmlPrefixMappings](../../reference/shared/customxmlprefixmappings.customxmlprefixmappings.md).


## Пример




```js
function setXMLPartNamespaceManagerNamespace() {
    Office.context.document.customXmlParts.getByIdAsync("{3BC85265-09D6-4205-B665-8EB239A8B9A1}", function (result) {
        var xmlPart = result.value;
        xmlPart.namespaceManager.addNamespaceAsync("myPrefix", "myNamespace");
    });
}

```




## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что это свойство поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает это свойство.

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
