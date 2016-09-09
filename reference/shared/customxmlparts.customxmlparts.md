
# Объект CustomXmlParts
Представляет коллекцию объектов [CustomXMLPart](../../reference/shared/customxmlpart.customxmlpart.md).

|||
|:-----|:-----|
|**Ведущие приложения:**|Word|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Последнее изменение в **|1.1|

```
Office.context.document.customXmlParts
```


## Элементы


**Методы**


|**Имя**|**Описание**|
|:-----|:-----|
|[addAsync](../../reference/shared/customxmlparts.addasync.md)|Асинхронно добавляет новую настраиваемую XML-часть в файл.|
|[getByIdAsync](../../reference/shared/customxmlparts.getbyidasync.md)|Асинхронно получает настраиваемую XML-часть по идентификатору.|
|[getByNamespaceAsync](../../reference/shared/customxmlparts.getbynamespaceasync.md)|Асинхронно получает массив настраиваемых XML-частей, соответствующих указанному пространству имен.|

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
