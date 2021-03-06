
# Свойство TableBinding.hasHeaders
Получает значение, указывающее, есть ли в таблице заголовки.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, PowerPoint, Project, Word|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Последнее изменение в Selection**|1.1|

```
var colCount = bindingObj.hasHeaders;
```


## Возвращаемое значение

Если в указанной таблице [TableBinding](../../reference/shared/binding.tablebinding.md) есть заголовки, возвращает **true**; в противном случае возвращает **false**.


## Пример




```js
function showBindingHasHeaders() {
    Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
        write("Binding has headers: " + asyncResult.value.hasHeaders);
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что это свойство поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает это свойство.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Поддерживаемые ведущие приложения по платформе**


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Да|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|TableBindings|
|**Минимальный уровень разрешений**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки





****


|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel и Word в Office для iPad.|
|1.1|Добавлена поддержка этого события в надстройках для Access.|
|1.0|Представлено|
