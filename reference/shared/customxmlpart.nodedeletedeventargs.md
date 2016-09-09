
# Объект NodeDeletedEventArgs
Предоставляет сведения об удаленном узле, который вызвал событие [dataNodeDeleted](../../reference/shared/customxmlpart.datanodedeleted.event.md).

|||
|:-----|:-----|
|**Ведущие приложения:**|Word|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Добавлено в версии**|1.1|

```
NodeDeletedEventArgs
```


## Элементы


**Свойства**


|**Имя**|**Описание**|
|:-----|:-----|
|[isUndoRedo](../../reference/shared/customxmlpart.isundoredo.md)|Получает значение, указывающее, был ли узел удален при отмене или повторе действия пользователя.|
|[oldNextSibling](../../reference/shared/customxmlpart.oldnextsibling.md)|Получает следующий одноуровневый узел для узла, который был только что удален из объекта **CustomXMLPart**.|
|[oldNode](../../reference/shared/customxmlpart.oldnode.md)|Получает узел, который был только что удален из объекта **CustomXmlPart**.|

## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает, что этот объект поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот объект.

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




|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Word в Office для iPad.|
|1.0|Представлено|
