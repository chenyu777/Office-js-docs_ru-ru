
# Объект Settings
Представляет настраиваемые параметры надстройки области задач или контентной надстройки, которые хранятся в документе узла как пары "имя-значение".

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, PowerPoint, Word|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Параметры|
|**Последнее изменение в **|1.1|

```
Office.context.document.settings
```


## Элементы


**Методы**

|||
|:-----|:-----|
|Имя|Описание|
|[addHandlerAsync](../../reference/shared/settings.addhandlerasync.md)|Добавляет обработчик события **settingsChanged**.|
|[get](../../reference/shared/settings.get.md)|Извлекает указанный параметр.|
|[refreshAsync](../../reference/shared/settings.refreshasync.md)|Считывает все параметры, сохраненные в документе, и обновляет копию этих параметров в памяти для надстройки.|
|[удалить](../../reference/shared/settings.remove.md)|Удаляет указанный параметр.|
|[removeHandlerAsync](../../reference/shared/settings.removehandlerasync.md)|Удаляет обработчик события **settingsChanged**.|
|[saveAsync](../../reference/shared/settings.saveasync.md)|Сохраняет параметры.|
|[set](../../reference/shared/settings.set.md)|Устанавливает или создает указанный параметр.|

**События**


|**Имя**|**Описание**|
|:-----|:-----|
|[settingsChanged](../../reference/shared/settings.settingschangedevent.md)|Происходит при изменении параметра.|

## Заметки

Параметры, созданные с помощью методов объекта **Settings**, сохраняются на уровне приложения и на уровне документа. Таким образом, они доступны только для создавшего их приложения и только из того документа, в котором они сохранены.

Имя параметра имеет тип **string**, а типом значения может быть **string**, **number**, **boolean**, **null**, **object** или **array**.

Объект **Settings** автоматически загружается как часть объекта [Document](../../reference/shared/document.md). К нему можно обратиться через свойство [settings](../../reference/shared/document.settings.md) этого объекта после активации надстройки. Разработчик должен предусмотреть вызов метода [saveAsync](../../reference/shared/settings.saveasync.md) после добавления или удаления параметров, чтобы сохранить параметры в документе.


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает, что этот объект поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот объект.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Да|Y|
|**PowerPoint**|Y|Да|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|Параметры|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки

|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel, PowerPoint и Word в Office для iPad.|
|1.1|С помощью методов **addHandlerAsync** и **removeHandlerAsync** добавлять и удалять обработчики событий в контентных надстройках для Access. Добавлена поддержка пользовательских параметров для методов **get**, **refreshAsync**, **remove**, **saveAsync** и **set**в контентных надстройках для Access.|
|1.0|Представлено|