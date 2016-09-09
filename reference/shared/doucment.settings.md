
# Объект Settings
Представляет настраиваемые параметры надстройки области задач или контентной надстройки, которые хранятся в документе узла как пары "имя-значение".

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, PowerPoint, Word|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Параметры|
|**Последнее изменение в**|1.1|

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
|[remove](../../reference/shared/settings.remove.md)|Удаляет указанный параметр.|
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


|**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Access**|Y|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|Параметры|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Библиотека**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки




|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel, PowerPoint и Word в Office для iPad.|
|1.1|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Для методов <a href="7c4780cf-a779-4ac9-a362-c0bacae64a96.htm">addHandlerAsync</a> and <a href="735a255b-2a86-4b43-b1fa-e2a305815615.htm">removeHandlerAsync</a> появилась возможность добавлять и удалять обработчики событий для события <span class="keyword">SettingsChanged</span> в контентных надстройках для Access. </p></li><li><p>Для методов <a href="aeac06dd-994e-4235-b208-1bd117395296.htm">get</a>, <a href="53a52c47-24b4-4d2d-b840-fe1b242cd795.htm">refreshAsync</a>, <a href="a92446bf-de65-45bd-8412-36ea8e77c5a2.htm">remove</a>, <a href="7147c221-937c-477c-98a6-f59d6200c27b.htm">saveAsync</a> и <a href="4e2c9758-953e-41e8-aca6-d8daf764a584.htm">set</a> добавлена поддержка настраиваемых параметров в контентных надстройках для Access.</p></li></ul>|
|1.0|Представлено|

