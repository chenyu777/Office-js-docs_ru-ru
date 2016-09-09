

# Метод Settings.refreshAsync
Считывает все параметры, сохраненные в документе, и обновляет копию этих параметров в памяти для контентной надстройки или надстройки области задач.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, PowerPoint, Word|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Параметры|
|**Последнее изменение в **|1.1|

```js
Office.context.document.settings.refreshAsync(callback);
```


## Параметры

_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Тип: **object**

&nbsp;&nbsp;&nbsp;&nbsp;Функция, вызываемая после получения результатов обратного вызова, единственный параметр которой имеет тип **AsyncResult**.

    



## Значение обратного вызова

Когда выполняется функция, переданная в параметр _callback_, она получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью единственного параметра функции обратного вызова.

Если функция обратного вызова передана методу **refreshAsync**, можно использовать свойства объекта **AsyncResult** для возврата следующей информации.



|**Свойство**|**Применение**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Получает доступ к объекту [Settings](../../reference/shared/settings.md) с обновленными значениями.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Определяет, удалось ли выполнить операцию.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Получает доступ к объекту [Error](../../reference/shared/error.md), который содержит сведения об ошибке, если операция завершилась неудачно.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Получает доступ к определенному пользователем объекту **object** или значению, если они передаются как параметр _asyncContext_.|

## Заметки

Этот метод полезен в сценариях совместного создания документов в Word и PowerPoint, когда несколько экземпляров надстройки работают с одним документом. Поскольку каждый экземпляр надстройки работает с копией параметров, загруженной в память из документа во время его открытия, значения параметров, используемые пользователями, могут рассинхронизироваться. Это может произойти, когда экземпляр надстройки вызывает метод [Settings.saveAsync](../../reference/shared/settings.saveasync.md) для сохранения всех параметров пользователя в документе. Вызов метода **refreshAsync** из обработчика события [settingsChanged](../../reference/shared/settings.settingschangedevent.md) экземпляра надстройки обновляет значения параметров для всех пользователей.

Метод **refreshAsync** вызывается из надстроек, созданных для Excel, однако это не имеет смысла, так как он не поддерживает совместную работу с документами.


## Пример




```js
function refreshSettings() {
    Office.context.document.settings.refreshAsync(function (asyncResult) {
        write('Settings refreshed with status: ' + asyncResult.status);
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).



||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Да|Y|
|**PowerPoint**|Y|Да|Y|
|**Word**|Y|Да|Y|

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
|1.1|Добавлена поддержка PowerPoint Online.|
|1.1|Добавлена поддержка Excel, PowerPoint и Word в Office для iPad.|
|1.1|Добавлена поддержка настраиваемых параметров в контекстных надстройках для Access.|
|1.0|Представлено|
