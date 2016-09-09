
# Свойство AsyncResult.value
Получает полезные данные или содержимое асинхронной операции (если имеется).

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Последнее изменение в **|1.1|

```js
var dataValue = asyncResult.value;
```


## Возвращаемое значение

Возвращает значение запроса во время выполнения асинхронного вызова. 


 >**Примечание.** Значение, возвращаемое свойством **value** для конкретного метода "Async", зависит от назначения и контекста этого метода. Сведения о значениях, возвращаемых свойством **value** для метода "Async", см. в разделе "Значение обратного вызова" описания соответствующего метода. Полный список методов "Async" см. в разделе "Заметки" описания объекта [AsyncResult](../../reference/shared/asyncresult.md).


## Замечания

Получает доступ к объекту **AsyncResult** предоставляет функция, переданная в качестве аргумента для параметра _callback_ метода "Async", такого как методы [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) и [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) объекта **Document**.


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

||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|**OWA для устройств**|**Office для Mac**|
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
