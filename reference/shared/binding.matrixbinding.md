
# Объект MatrixBinding
Представляет привязку в двух измерениях строк и столбцов. 

|||
|:-----|:-----|
|**Ведущие приложения:**|Excel, Word|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings|
|**Последнее изменение в Selection**|1.1|

```
MatrixBinding
```


**Свойства**


|**Имя**|**Описание**|
|:-----|:-----|
|[columnCount](../../reference/shared/binding.matrixbinding.columncount.md)|Получает количество столбцов в матричной структуре данных. Целочисленное значение.|
|[rowCount](../../reference/shared/binding.matrixbinding.rowcount.md)|Получает количество строк в матричной структуре данных. Целочисленное значение.|

## Заметки

Объект **MatrixBinding** наследует свойство [id](../../reference/shared/binding.id.md), свойство [type](../../reference/shared/binding.type.md), метод [getDataAsync](../../reference/shared/binding.getdataasync.md) и метод [setDataAsync](../../reference/shared/binding.setdataasync.md) от объекта [Binding](../../reference/shared/binding.md).


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Поддерживаемые ведущие приложения по платформе**


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Да|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|MatrixBindings|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



****


|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel и Word в Office для iPad.|
|1.0|Представлено|
