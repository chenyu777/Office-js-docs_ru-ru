
# Перечисление EventType
Указывает тип вызванного события. Возвращается свойством **type** объекта _EventName_**EventArgs**.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, PowerPoint, Project, Word|
|**Последнее изменение в Selection**|1.1|

```js
Office.EventType
```


## Элементы


**Значения**


|Перечисление|Значение|Описание|
|:-----|:-----|:-----|
|Office.EventType.ActiveViewChanged|"documentActiveViewChanged"|Было вызвано событие [Document.ActiveViewChanged](../../reference/shared/document.activeviewchanged.md).|
|Office.EventType.DocumentSelectionChanged|"documentSelectionChanged"|Было вызвано событие [Document.SelectionChanged](../../reference/shared/document.selectionchanged.event.md).|
|Office.EventType.BindingSelectionChanged|"bindingSelectionChanged"|Было вызвано событие [Binding.BindingSelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md).|
|Office.EventType.BindingDataChanged|"bindingDataChanged"|Было вызвано событие [Binding.BindingDataChanged](../../reference/shared/binding.bindingdatachangedevent.md).|
|Office.EventType.DataNodeDeleted|"nodeDeleted"|Было вызвано событие [CustomXmlPart.dataNodeDeleted](../../reference/shared/customxmlpart.datanodedeleted.event.md).|
|Office.EventType.DataNodeInserted|"nodeInserted"|Было вызвано событие [CustomXmlPart.dataNodeInserted](../../reference/shared/customxmlpart.datanodeinserted.event.md).|
|Office.EventType.DataNodeReplaced|"nodeReplaced"|Было вызвано событие [CustomXmlPart.dataNodeReplaced](../../reference/shared/customxmlpart.datanodereplaced.event.md).|
|Office.EventType.SettingsChanged|"settingsChanged"|Было вызвано событие [Settings.settingsChanged](../../reference/shared/settings.settingschangedevent.md).|

## Замечания


 >**Примечание**.  Надстройки для Project поддерживают типы событий **Office.EventType.ResourceSelectionChanged**, **Office.EventType.TaskSelectionChanged** и **Office.EventType.ViewSelectionChanged**.


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что данное перечисление поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает это перечисление.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Поддерживаемые ведущие приложения по платформе**


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Да|Y|
|**PowerPoint**|Y|Y||
|**Project**|Y|||
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



|**Версия**|**Изменения**|
|:-----|:-----|
|1.1| Добавлено перечисление Office.EventType.ActiveViewChanged для нового события **Document.ActiveViewChanged**.|
|1.0|Представлено|
