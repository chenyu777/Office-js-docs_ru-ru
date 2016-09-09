
# Объект TextBinding
Представляет выбор привязанного текста в документе.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, PowerPoint, Project, Word|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TextBindings|
|**Добавлен в версии**|1.0|

```
TextBinding
```


## Заметки

Объект **TextBinding** наследует свойство [id](../../reference/shared/binding.id.md), свойство [type](../../reference/shared/binding.type.md), метод [getDataAsync](../../reference/shared/binding.getdataasync.md) и метод [setDataAsync](../../reference/shared/binding.setdataasync.md) от объекта [Binding](../../reference/shared/binding.md). Он не реализует дополнительные свойства или методы.


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает, что этот объект поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот объект.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Поддерживаемые ведущие приложения по платформе**


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Да|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|TextBindings|
|**Минимальный уровень разрешений**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



****


|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel и Word в Office для iPad.|
|1.0|Представлено|
