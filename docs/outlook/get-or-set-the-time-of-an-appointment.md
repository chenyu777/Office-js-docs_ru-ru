
# Считывание и запись времени при создании встречи в Outlook

API JavaScript для Office предоставляет асинхронные методы ([Time.getAsync](../../reference/outlook/Time.md) и [Time.setAsync](../../reference/outlook/Time.md)), чтобы получать и задавать время начала или окончания встречи, создаваемой пользователем. Эти методы доступны только для надстроек создания. Чтобы использовать их, необходимо настроить манифест для активации надстройки в формах создания Outlook, как описано в статье [Создание надстроек Outlook для форм создания](../outlook/compose-scenario.md).

Свойства [start](../../reference/outlook/Office.context.mailbox.item.md) и [end](../../reference/outlook/Office.context.mailbox.item.md) доступны для встреч в формах создания и чтения. в форме чтения доступ к свойствам можно получить напрямую из родительского объекта, как в следующем примере:




```
item.start
```

И в этом примере:




```
item.end
```

Но так как в форме создания и пользователь, и ваша надстройка могут вставлять или изменять сведения о времени одновременно, для получения времени начала и окончания необходимо использовать асинхронный метод  **getAsync**, как показано ниже:




```
item.start.getAsync
```

И в следующем примере:




```
item.end.getAsync
```

Как и большинство асинхронных методов в API JavaScript для Office, методы  **getAsync** и **setAsync** принимают необязательные входящие параметры. Дополнительные сведения о считывании этих параметров см. в разделе [Передача необязательных параметров в асинхронные методы](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) статьи [Асинхронное программирование в надстройках для Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md).


## Получение времени начала или окончания


В этом разделе показан пример кода, который получает время начала встречи, создаваемой пользователем, и отображает его. Вы можете использовать тот же код, заменив свойство  **start** на **end**, чтобы получить время окончания. В этом примере предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи, как показано ниже.


```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>

```

Чтобы использовать методы  **item.start.getAsync** и **item.end.getAsync**, предоставьте метод обратного вызова, который проверяет состояние и результат асинхронного вызова. Вы можете указать любые необходимые аргументы метода обратного вызова с помощью дополнительного параметра  _asyncContext_. Состояние, результаты и сообщения об ошибках можно получить с помощью выходного параметра  _asyncResult_ метода обратного вызова. Если асинхронный вызов выполнен успешно, вы можете получить начальное время как объект **Date** в формате UTC, используя свойство [AsyncResult.value](../../reference/outlook/simple-types.md).




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the start time of the item being composed.
        getStartTime();
    });
}

// Get the start time of the item that the user is composing.
function getStartTime() {
    item.start.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the start time, display it, first in UTC and 
                // then convert the Date object to local time and display that.
                write ('The start time in UTC is: ' + asyncResult.value.toString());
                write ('The start time in local time is: ' + asyncResult.value.toLocaleString());
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Установка времени начала или окончания


В этом разделе показан пример кода, получающий время начало встречи, создаваемой пользователем. Можно использовать тот же код, заменив свойство  **start** на **end**, чтобы получить время начала. Обратите внимание, что если у формы создания уже есть время начала, последующая установка времени начала приведет к изменению времени окончания, чтобы сохранить предыдущую длительность встречи. Если у формы создания уже есть время окончания, последующая установка времени окончания приведет к изменению длительности и времени окончания. Если встреча создана как событие на весь день, установки времени начала приведет к смещению времени окончания на 24 часа и отмены выбора параметра события на весь день в форме создания.

Как и в предыдущем примере, здесь предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи.

Чтобы использовать методы  **item.start.setAsync** и **item.end.setAsync**, укажите значение  **Date** в формате UTC в параметре _dateTime_. Если вы получаете дату на основе данных, введенных пользователем в клиенте, с помощью [mailbox.convertToUtcClientTime](../../reference/outlook/Office.context.mailbox.md) можно преобразовать полученное значение в объект **Date** в формате UTC. Можно предоставить необязательный метод обратного вызова и все его аргументы в параметре _asyncContext_. Следует проверить состояние, результат и наличие ошибок в выходном параметре  _asyncResult_ метода обратного вызова. Если асинхронный вызов выполнен успешно, **setAsync** вставляет указанное строку времени начала или окончания как обычный текст, перезаписывая существующее время начала или окончания для этого элемента.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the start time of the item being composed.
        setStartTime();
    });
}

// Set the start time of the item that the user is composing.
function setStartTime() {
    var startDate = new Date("September 27, 2012 12:30:00");
    
    item.start.setAsync(
        startDate,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the start time.
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
    
- [Считывание и запись расположения при создании встречи в Outlook](../outlook/get-or-set-the-location-of-an-appointment.md)
    
