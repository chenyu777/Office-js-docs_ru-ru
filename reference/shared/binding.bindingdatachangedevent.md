
# Событие Binding.bindingDataChanged
Происходит при изменении данных в привязке.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Word|
|**Последнее изменение в BindingEvents**|1.1|

```js
Office.EventType.BindingDataChanged
```


## Заметки

Для добавления обработчика события привязки **BindingDataChanged** используйте метод [addHandlerAsync](../../reference/shared/binding.addhandlerasync.md) объекта **Binding**. Обработчик события получает аргумент типа [BindingDataChangedEventArgs](../../reference/shared/binding.bindingdatachangedeventargs.md).


## Пример




```js
function addEventHandlerToBinding() {
    Office.select("bindings#MyBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
}

function onBindingDataChanged(eventArgs) {
    write("Data has changed in binding: " + eventArgs.binding.id);
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

|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel и Word в Office для iPad.|
|1.1|Добавлена поддержка этого события в надстройках для Access.|
|1.0|Представлено|
