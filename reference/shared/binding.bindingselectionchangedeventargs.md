
# Объект BindingSelectionChangedEventArgs
Предоставляет сведения о привязке, которая вызвала событие [SelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md).

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Word|
|**Последнее изменение в TableBinding**|1.1|

```
Office.EventType.BindingSelectionChanged
```


## Элементы


**Свойства**


|**Имя**|**Описание**|
|:-----|:-----|
|[binding](../../reference/shared/binding.bindingselectionchangedevent.binding.md)|Получает объект [Binding](../../reference/shared/binding.md), представляющий привязку, вызвавшую событие **SelectionChanged**.|
|[columnCount](../../reference/shared/binding.bindingselectionchangedevent.columncount.md)|Получает количество выбранных столбцов.|
|[rowCount](../../reference/shared/binding.bindingselectionchangedevent.rowcount.md)|Получает количество выбранных строк.|
|[startRow](../../reference/shared/binding.bindingselectionchangedevent.startrow.md)|Получает индекс первой строки выборки (с отсчетом от нуля).|
|[startColumn](../../reference/shared/binding.bindingselectionchangedevent.startcolumn.md)|Получает индекс первого столбца текущего выбора (с отсчетом от нуля).|
|[type](../../reference/shared/binding.bindingselectionchangedevent.type.md)|Получает значение перечисления [EventType](../../reference/shared/eventtype-enumeration.md), указывающее вид вызванного события.|

## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Поддерживаемые ведущие приложения по платформе**


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Да|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



****


|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel и Word в Office для iPad.|
|1.1|Добавлена поддержка привязки таблиц в надстройках для Access.|
|1.0|Представлено|
