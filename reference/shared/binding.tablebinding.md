
# Объект TableBinding
Представляет привязку в двух измерениях строк и столбцов, куда при желании можно добавить заголовки.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, PowerPoint, Project, Word|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Последнее изменение в Selection**|1.1|

```
TableBinding
```


## Элементы


**Свойства**


|**Имя**|**Описание**|**Обновления для Office.js версии 1.1**|
|:-----|:-----|:-----|
|[columnCount](../../reference/shared/binding.tablebinding.columncount.md)|Получает количество столбцов в указанном объекте **TableBinding**.|Добавлена поддержка табличной привязки в контентных надстройках для Access.|
|[hasHeaders](../../reference/shared/binding.tablebinding.hasheaders.md)|Если в указанном объекте **TableBinding** есть заголовки, возвращается значение true, а в противном случае — false.|Добавлена поддержка табличной привязки в контентных надстройках для Access.|
|[rowCount](../../reference/shared/binding.tablebinding.rowcount.md)|Количество строк в указанном объекте **TableBinding**.|Для повышения производительности в контентных надстройках Access всегда возвращается значение –1.|

**Методы**


|**Имя**|**Описание**|**Обновления для Office.js версии 1.1**|
|:-----|:-----|:-----|
|[addColumnsAsync](../../reference/shared/binding.tablebinding.addcolumnsasync.md)|Добавляет столбцы и значения в таблицу.||
|[addRowsAsync](../../reference/shared/binding.tablebinding.addrowsasync.md)|Добавляет строки и значения в таблицу.|Добавлена поддержка табличной привязки в контентных надстройках для Access.|
|[clearFormatsAsync](../../reference/shared/binding.tablebinding.clearformatsasync.md)|Очищает форматирование в привязанной таблице.|Новые возможности Office.js версии 1.1 для надстроек Excel.|
|[deleteAllDataValuesAsync](../../reference/shared/binding.tablebinding.deletealldatavaluesasync.md)|Удаляет из таблицы все строки и их значения, кроме строк заголовка. Сдвиг зависит от ведущего приложения.|Добавлена поддержка табличной привязки в контентных надстройках для Access.|
|[setDataAsync](../../reference/shared/binding.setdataasync.md)|Записывает данные в привязанный раздел документа, представленный указанным объектом привязки.|<ul><li>Добавлена поддержка табличной привязки в контентных надстройках для Access.</li><li>Добавлена поддержка настройки форматирования при записи данных в связанные таблицы в надстройках Excel.</li></ul>|
|[setFormatsAsync](../../reference/shared/binding.tablebinding.setformatsasync.md)|Задает формат ячеек и таблиц для указанных элементов и данных в связанной таблице.|Позволяет задавать формат таблиц в надстройках для Excel.|
|[setTableOptionsAsync](../../reference/shared/binding.tablebinding.settableoptionsasync.md)|Обновляет параметры форматирования связанной таблицы.|Позволяет задавать формат таблиц в надстройках для Excel.|

## Заметки

Объект **TableBinding** наследует свойство [id](../../reference/shared/binding.id.md), свойство [type](../../reference/shared/binding.type.md), метод [getDataAsync](../../reference/shared/binding.getdataasync.md) и метод [setDataAsync](../../reference/shared/binding.setdataasync.md) от абстрактного объекта [Binding](../../reference/shared/binding.md).

После создания табличной привязки в Excel каждая новая строка, добавляемая пользователем в таблицу, автоматически включается в привязку (значение **rowCount** будет увеличиваться).


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает, что этот объект поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот объект.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Поддерживаемые ведущие приложения по платформе**


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Да|Y|
|**Word**|Y|Да|Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|TableBindings|
|**Минимальный уровень разрешений**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки




|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel и Word в Office для iPad.|
|1.1|Добавлена поддержка [задания формата при вставке таблиц](../../docs/excel/format-tables-in-add-ins-for-excel.md) в Excel.|
|1.1|Добавлена поддержка надстроек для Access.|
|1.0|Представлено|
