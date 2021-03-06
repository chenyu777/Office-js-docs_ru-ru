

# Событие Settings.settingsChanged
Возникает, когда копия контейнера свойств в памяти сохраняется в документ с помощью метода [Settings.saveAsync](../../reference/shared/settings.saveasync.md).

|||
|:-----|:-----|
|**Ведущие приложения:**|Excel |
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Параметры|
|**Последнее изменение в **|1.0|

```js
Office.EventType.SettingsChanged
```


## Заметки

Чтобы добавить обработчик для события **settingsChanged**, используйте метод [addHandlerAsync](../../reference/shared/settings.addhandlerasync.md) объекта **Settings**.

Событие **settingsChanged** возникает, только если скрипт надстройки вызывает метод **Settings.saveAsync** для сохранения копии параметров в памяти в файл документа. Событие **settingsChanged** не вызывается при вызове методов [Settings.set](../../reference/shared/settings.set.md) и [Settings.remove](../../reference/shared/settings.remove.md).

Событие **settingsChanged** разработано для разрешения возможных конфликтов, когда несколько пользователей пытаются одновременно сохранить параметры, а надстройка используется в общем (совместно редактируемом) документе.


 >**Важно!** Обработчик события **settingsChanged** можно зарегистрировать с помощью кода вашей надстройки, когда эта надстройка работает с клиентом Excel. Но событие будет возникать, только если электронная таблица загружаемой надстройки открывается в Excel Online _и_ с ней работают несколько пользователей (совместное редактирование). Поэтому фактически событие **settingsChanged** поддерживается только в Excel Online со сценарием совместного редактирования.


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает, что данное событие поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает это событие.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).



||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||Y||

|||
|:-----|:-----|
|**Доступен в наборах требований**|Параметры|
|**Минимальный уровень разрешений**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки




|**Версия**|**Изменения**|
|:-----|:-----|
|1.0|Представлено|
