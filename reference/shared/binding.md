
# Объект Binding
Абстрактный класс, представляющий привязку к разделу документа.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Word|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBinding, TableBinding, TextBinding|
|**Последнее изменение в TableBinding**|1.1|

```js
Office.context.document.bindings.getByIdAsync(id);
```

## Элементы


**Объекты**


|**Имя**|**Описание**|
|:-----|:-----|
|[MatrixBinding](../../reference/shared/binding.matrixbinding.md)|Представляет привязку в двух измерениях строк и столбцов.|
|[TableBinding](../../reference/shared/binding.tablebinding.md)|Представляет привязку в двух измерениях строк и столбцов, при желании с заголовками.|
|[TextBinding](../../reference/shared/binding.textbinding.md)|Представляет выбор привязанного текста в документе.|

**Свойства**


|**Имя**|**Описание**|
|:-----|:-----|
|[document](../../reference/shared/binding.document.md)|Получает объект **Document**, связанный с привязкой.|
|[id](../../reference/shared/binding.id.md)|Получает идентификатор объекта.|
|[type](../../reference/shared/binding.type.md)|Получает тип привязки.|

**Методы**


|**Имя**|**Описание**|
|:-----|:-----|
|[addHandlerAsync](../../reference/shared/binding.addhandlerasync.md)|Добавляет обработчик к привязке для указанного типа события.|
|[getDataAsync](../../reference/shared/binding.getdataasync.md)|Возвращает данные, содержащиеся в привязке.|
|[removeHandlerAsync](../../reference/shared/binding.removehandlerasync.md)|Удаляет указанный обработчик из привязки для указанного типа события.|
|[setDataAsync](../../reference/shared/binding.setdataasync.md)|Записывает данные в привязанный раздел документа, представленный указанным объектом привязки.|
|[TableBinding.setFormatsAsync](../../reference/shared/binding.tablebinding.setformatsasync.md)|Задает или обновляет форматирование определенных элементов и данных в связанной таблице.|

**События**


|**Имя**|**Описание**|
|:-----|:-----|
|[bindingDataChanged](../../reference/shared/binding.bindingdatachangedevent.md)|Происходит при изменении данных в привязке.|
|[bindingSelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md)|Происходит при изменении выбора в привязке.|

## Заметки

Объект **Binding** предоставляет функциональные возможности, которыми обладают все привязки независимо от типа.

Объект **Binding** никогда не вызывается непосредственно. Это абстрактный родительский класс объектов, представляющих типы привязки [MatrixBinding](../../reference/shared/binding.matrixbinding.md), [TableBinding](../../reference/shared/binding.tablebinding.md) и [TextBinding](../../reference/shared/binding.textbinding.md). Все три объекта наследуют методы **getDataAsync** и **setDataAsync** от объекта **Binding**, который позволяет взаимодействовать с данными в привязке. Они также наследуют свойства **id** и **type** и позволяют запрашивать значения этих свойств. Кроме того, объекты **MatrixBinding** и **TableBinding** предоставляют дополнительные методы для работы с матрицами и таблицами, например, для подсчета количества строк и столбцов.


## Сведения о поддержке


Поддержка каждого элемента API объекта **Binding** зависит от ведущего приложения Office. Информацию о поддержке элемента в том или ином приложении см. в соответствующем разделе "Сведения о поддержке".

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


|||
|:-----|:-----|
|**Доступен в наборах требований**|MatrixBinding, TableBinding, TextBinding|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|
