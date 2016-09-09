
# Свойство AsyncResult.asyncContext
Получает определяемый пользователем элемент, передаваемый необязательному параметру _asyncContext_ вызванного метода в том состоянии, в каком был передан.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Последнее изменение в **|1.1|

```
var myContext = asynchResult.asyncContext;
```


## Возвращаемое значение

Возвращает пользовательский элемент (который может быть любым типом JavaScript: **String**, **Number**, **Boolean**, **Object**, **Array**, **Null** или **Undefined**), переданный в необязательном параметре _asyncContext_ вызванного метода. Возвращает значение **Undefined**, если параметр _asyncContext_ не был указан.


## Пример




```js
function getDataWithContext() {
    var format = "Your data: ";
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, { asyncContext: format }, showDataWithContext);
}

 function showDataWithContext(asyncResult) {
    write(asyncResult.asyncContext + asyncResult.value);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```




## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Поддерживаемые ведущие приложения по платформе**


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|**OWA для устройств**|**Outlook для Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**|Y|||||
|**Excel**|Y|Да|Y|||
|**Outlook**|Y|Да||Да|Y|
|**PowerPoint**|Y|Да|Y|||
|**Project**||||||
|**Word**|Y|Да|Y|||

|||
|:-----|:-----|
|**Минимальный уровень разрешений**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Надстройки области задач, надстройки Outlook, контентные надстройки|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



****


|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка PowerPoint Online.|
|1.1|Добавлена поддержка Excel, PowerPoint и Word в Office для iPad.|
|1.1|Добавлена поддержка надстроек для Access.|
|1.0|Представлено|
