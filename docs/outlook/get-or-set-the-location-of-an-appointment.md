
# Считывание и запись расположения при создании встречи в Outlook

API JavaScript для Office предоставляет асинхронные методы ([getAsync](../../reference/outlook/Location.md) и [setAsync](../../reference/outlook/Location.md)), чтобы получать и задавать место проведения встречи, создаваемой пользователем. Эти методы доступны только для надстроек создания. Чтобы использовать их, необходимо настроить манифест для активации надстройки в формах создания Outlook, как описано в статье [Создание надстроек Outlook для форм создания](../outlook/compose-scenario.md).

Свойство [location](../../reference/outlook/Office.context.mailbox.item.md) доступно для чтения в формах просмотра и создания встреч. В форме просмотра к этому свойству можно обращаться непосредственно из родительского объекта, например:




```js
item.location
```

Но в формах создания, в которых пользователь и надстройка могут одновременно вставлять или изменять сведения о расположении, для получения этих сведений следует использовать асинхронный метод  **getAsync**, как показано ниже:




```js
item.location.getAsync
```

Свойство  **location** доступно для записи только в формах создания встреч.

Как и большинство асинхронных методов в API JavaScript для Office, методы **getAsync** и **setAsync** принимают необязательные входные параметры. Дополнительные сведения об указании этих параметров см. в статье [Асинхронное программирование в надстройках для Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md).


## Получение расположения


В этом разделе приводится пример кода, который получает и отображает сведения о месте создаваемой пользователем встречи. В этом примере предполагается, что в манифесте надстройки указано правило, согласно которому надстройка активируется в форме создания встречи, как показано ниже.


```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>

```

Чтобы использовать метод  **item.location.getAsync**, создайте метод обратного вызова, который проверяет состояние и результат асинхронного вызова. Все необходимые аргументы метода обратного вызова можно передать через необязательный параметр  _asyncContext_. Состояние, результат и возможные ошибки можно считывать с помощью выходного параметра  _asyncResult_ метода обратного вызова. если асинхронный вызов выполнен успешно, строковое значение расположения можно получить с помощью свойства [AsyncResult.value](../../reference/outlook/simple-types.md).




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        getLocation();
    });
}

// Get the location of the item that the user is composing.
function getLocation() {
    item.location.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the location, display it.
                write ('The location is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Установка расположения


В этом разделе приводится пример кода, который задает место создаваемой пользователем встречи. Как и в предыдущем примере, здесь предполагается, что в манифесте надстройки указано правило, согласно которому она активируется в форме создания встречи.

Чтобы использовать метод  **item.location.setAsync**, укажите строку длиной до 255 символов в параметре данных. При необходимости можно указать метод обратного вызова и его аргументы в качестве параметра  _asyncContext_. Состояние, результат и возможное сообщение об ошибке можно получить из выходного параметра  _asyncResult_ метода обратного вызова. Если асинхронный вызов выполнен успешно, метод **setAsync** вставляет указанную строку расположения в виде обычного текста, заменяя предыдущее расположение.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the location of the item being composed.
        setLocation();
    });
}

// Set the location of the item that the user is composing.
function setLocation() {
    item.location.setAsync(
        'Conference room A',
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the location.
                // Do whatever appropriate for your scenario
                // using the arguments var1 and var2 as applicable.
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Дополнительные ресурсы



- [Считывание и запись данных элемента в форме создания элементов Outlook](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [Считывание и запись данных элемента Outlook в формах просмотра и создания](../outlook/item-data.md)
    
- [Создание надстроек Outlook для форм создания](../outlook/compose-scenario.md)
    
- [Асинхронное программирование в случае надстроек Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    
- [Чтение, запись и добавление получателей при создании встречи или сообщения в Outlook](../outlook/get-set-or-add-recipients.md)
    
- [Считывание и запись темы при создании встречи или сообщения в Outlook](../outlook/get-or-set-the-subject.md)
    
- [Вставка данных в основной текст при создании встречи или сообщения в Outlook](../outlook/insert-data-in-the-body.md)
    
- [Считывание и запись времени при создании встречи в Outlook](../outlook/get-or-set-the-time-of-an-appointment.md)
    
