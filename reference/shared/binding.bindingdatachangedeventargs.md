
# Объект BindingDataChangedEventArgs
Предоставляет сведения о привязке, которая вызвала событие [DataChanged](../../reference/shared/binding.bindingdatachangedevent.md).

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Word|
|**Последнее изменение в BindingEvents**|1.1|

```js
Office.EventType.BindingDataChanged
```


## Элементы


**Свойства**


|**Имя**|**Описание**|
|:-----|:-----|
|[binding](../../reference/shared/binding.bindingdatachangedeventargs.binding.md)|Получает объект [Binding](../../reference/shared/binding.md), представляющий привязку, которая вызвала событие **DataChanged**.|
|[type](../../reference/shared/binding.bindingdatachangedeventargs.type.md)|Получает значение перечисления [EventType](../../reference/shared/eventtype-enumeration.md), указывающее вид вызванного события.|

## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает, что этот объект поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот объект.

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




|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel и Word в Office для iPad.|
|1.1|Добавлена поддержка этого события в надстройках для Access.|
|1.0|Представлено|
