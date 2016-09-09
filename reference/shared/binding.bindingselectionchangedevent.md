
# Событие Binding.bindingSelectionChanged
Происходит при изменении выбора в привязке.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Word|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|BindingEvents|
|**Последнее изменение в Selection**|1.1|

```
Office.EventType.BindingSelectionChanged
```

## Замечания

Для добавления обработчика события привязки **BindingSelectionChanged** используйте метод [addHandlerAsync](../../reference/shared/binding.addhandlerasync.md) объекта **Binding**. Обработчик событий получает аргумент типа [BindingSelectionChangedEventArgs](../../reference/shared/binding.bindingselectionchangedeventargs.md).


## Пример




```
function addEventHandlerToBinding() {
 Office.select("bindings#MyBinding").addHandlerAsync(Office.EventType.BindingSelectionChanged, onBindingSelectionChanged);
}

function onBindingSelectionChanged(eventArgs) {
    write(eventArgs.binding.id + " has been selected.");
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает, что данное событие поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает это событие.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Поддерживаемые ведущие приложения по платформе**


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Да|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|BindingEvents|
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
