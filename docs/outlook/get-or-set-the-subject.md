
# Считывание и запись темы при создании встречи или сообщения в Outlook

API JavaScript для Office предоставляет асинхронные методы ([subject.getAsync](../../reference/outlook/Subject.md) и [subject.setAsync](../../reference/outlook/Subject.md)), чтобы получать и задавать тему встречи или сообщения, создаваемого пользователем. Эти методы доступны только для надстроек создания. Чтобы использовать их, необходимо настроить манифест для активации надстройки в формах создания Outlook.

Свойство  **subject** доступно для чтения в формах создания и формах чтения встреч и сообщений. В форме чтения доступ к свойству можно получить напрямую из родительского объекта, например:




```js
item.subject
```

Но так как в форме создания и пользователь, и ваша надстройка могут вставлять или изменять тему одновременно, для получения темы необходимо использовать асинхронный метод  **getAsync**, как показано ниже:




```js
item.subject.getAsync
```

Свойство  **subject** доступно для записи только в формах создания, но не в формах чтения.

Как и большинство асинхронных методов в API JavaScript для Office, методы **getAsync** и **setAsync** принимают необязательные входные параметры. Дополнительные сведения об указании этих параметров см. в разделе "Передача дополнительных параметров в асинхронные методы" статьи [Асинхронное программирование в надстройках для Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md).


## Получение темы


В этом разделе показан пример кода, получающий и отображающий тему создаваемой встречи или сообщения. В примере предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи или сообщения, как показано ниже.


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```

Чтобы использовать метод  **item.subject.getAsync**, предоставьте метод обратного вызова, который проверяет состояние и результат асинхронного вызова. Вы можете указать любые необходимые аргументы метода обратного вызова с помощью дополнительного параметра  _asyncContext_. Состояние, результаты и сообщения об ошибках можно получить с помощью выходного параметра  _asyncResult_ метода обратного вызова. Если асинхронный вызов выполнен успешно, вы можете получить тему как текстовую строку, используя свойство [AsyncResult.value](../../reference/outlook/simple-types.md).




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the subject of the item being composed.
        getSubject();
    });
}

// Get the subject of the item that the user is composing.
function getSubject() {
    item.subject.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the subject, display it.
                write ('The subject is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Установка темы


В этом разделе показан пример кода, задающий тему создаваемой встречи или сообщения. Как и в предыдущем примере, предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи или сообщения.

Чтобы использовать метод  **item.subject.setAsync**, укажите строку длиной до 255 символов в параметре data. При необходимости можно предоставить метод обратного вызова и все его аргументы в параметре  _asyncContext_. Следует проверить состояние, результат и наличие ошибок в выходном параметре  _asyncResult_ метода обратного вызова. Если асинхронный вызов выполнен успешно, **setAsync** вставляет указанную строку темы как обычный текст, перезаписывая существующую тему этого элемента.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the subject of the item being composed.
        setSubject();
    });
}

// Set the subject of the item that the user is composing.
function setSubject() {
    var today = new Date();
    var subject;

    // Customize the subject with today's date.
    subject = 'Summary for ' + today.toLocaleDateString();

    item.subject.setAsync(
        subject,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the subject.
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
    
- [Вставка данных в основной текст при создании встречи или сообщения в Outlook](../outlook/insert-data-in-the-body.md)
    
- [Считывание и запись расположения при создании встречи в Outlook](../outlook/get-or-set-the-location-of-an-appointment.md)
    
- [Считывание и запись времени при создании встречи в Outlook](../outlook/get-or-set-the-time-of-an-appointment.md)
    
