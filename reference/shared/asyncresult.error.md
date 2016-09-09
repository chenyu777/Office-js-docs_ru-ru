
# Свойство AsyncResult.error
Получает объект **Error** с описанием ошибки, если таковая возникает.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Последнее изменение в **|1.1|

```js
var errorObj = asyncResult.error;
```


## Возвращаемое значение

Объект **[Error](../../reference/shared/error.md)**.


## Пример




```js
function getData() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Table, function(asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write(asyncResult.error.message);
        }
        else {
            write(asyncResult.value);
        }
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


| |**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|**OWA для устройств**|**Outlook для Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Да|Y|||
|**Outlook**|Y|Да||Да|Y|
|**PowerPoint**|Y|Да|Y|||
|**Project**|Y|||||
|**Word**|Y|Да|Y|||

|||
|:-----|:-----|
|**Минимальный уровень разрешений**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Надстройки области задач, надстройки Outlook, контентные надстройки|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка PowerPoint Online.|
|1.1|Добавлена поддержка Excel, PowerPoint и Word в Office для iPad.|
|1.1|Добавлена поддержка надстроек для Access.|
|1.0|Представлено|
