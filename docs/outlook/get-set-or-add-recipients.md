
# Чтение, запись и добавление получателей при создании встречи или сообщения в Outlook


API JavaScript для Office предоставляет асинхронные методы ([Recipients.getAsync](../../reference/outlook/Recipients.md), [Recipients.setAsync](../../reference/outlook/Recipients.md) или [Recipients.addAysnc](../../reference/outlook/Recipients.md)), чтобы соответственно получать, задавать и добавлять получателей в форме создания встречи или сообщения. Эти асинхронные методы доступны только для надстроек создания. Чтобы использовать их, необходимо настроить манифест для активации надстройки в формах создания Outlook, как описано в статье [Создание надстроек Outlook для форм создания](../outlook/compose-scenario.md).

Некоторые свойства, представляющие получателей в сообщении или встрече, доступны для чтения в формах создания и чтения. Это свойства [optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md) и [requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md) для встреч, а также свойства [cc](../../reference/outlook/Office.context.mailbox.item.md) и [to](../../reference/outlook/Office.context.mailbox.item.md) для сообщений. В форме чтения доступ к свойству можно получить напрямую из родительского объекта, например:




```js
item.cc
```

Но так как в форме создания и пользователь, и ваша надстройка могут вставлять или изменять получателя одновременно, для получения этих свойств необходимо использовать асинхронный метод  **getAsync**, как в следующем примере:




```js
item.cc.getAsync
```

Эти свойства доступны для записи только в формах создания, но не формах чтения.

Как и большинство асинхронных методов в JavaScript API для Office,  **getAsync**,  **setAsync** и **addAsync** принимают дополнительные входные параметры. Дополнительные сведения о передаче этих параметров см. в разделе [Передача необязательных параметров в асинхронные методы](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) статьи [Асинхронное программирование в надстройках для Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md).


## Извлечение получателей


В этом разделе показан пример кода, извлекающий получателей создаваемой встречи или сообщения, и показывающий адреса электронной почты получателей. В примере предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи или сообщения, как показано ниже. 


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

В API JavaScript для Office, так как свойства, представляющие получателей встречи ( **optionalAttendees** м **requiredAttendees**), отличаются от соответствующих свойств сообщения ([bcc](../../reference/outlook/Office.context.mailbox.item.md),  **cc** и **to**), следует использовать свойство [item.itemType](../../reference/outlook/Office.context.mailbox.item.md), чтобы определить, является ли создаваемый элемент встречей или сообщением. В режиме создания все эти свойства встреч и сообщений являются объектами [Recipients](../../reference/outlook/Recipients.md), поэтому вы можете применить асинхронный метод,  **Recipients.getAsync**, чтобы получить соответствующих получателей. 

Для использования  **getAsync** предоставьте метод обратного вызова, чтобы проверить состояние, результаты и наличие ошибок, возвращенных асинхронным вызовом **getAsync**. Методу обратного вызова можно передать любые аргументы, используя дополнительный параметр  _asyncContext_. Метод обратного вызова возвращает выходной параметр  _asyncResult_. С помощью свойств  **status** и **error** объекта [AsyncResult](../../reference/outlook/simple-types.md) можно проверить состояние и наличие сообщений об ошибках для асинхронного вызова, с помощью свойства **value** можно извлечь фактических получателей. Они представляются как массив объектов [EmailAddressDetails](../../reference/outlook/simple-types.md).

Так как метод  **getAsync** является асинхронным, при наличии последующих действий, зависящих от успешного извлечения получателей, такие действия должны начинаться в коде только в соответствующем методе обратного вызова после успешного завершения асинхронного вызова.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get all the recipients of the composed item.
        getAllRecipients();
    });
}

// Get the email addresses of all the recipients of the composed item.
function getAllRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;
    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method getAsync to get each type of recipients
    // of the composed item. Each time, this example passes an anonymous 
    // callback function that doesn't take any parameters.
    toRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get to-recipients of the item completed.
            // Display the email addresses of the to-recipients. 
            write ('To-recipients of the item:');
            displayAddresses(asyncResult);
        }    
    }); // End getAsync for to-recipients.

    // Get any cc-recipients.
    ccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get cc-recipients of the item completed.
            // Display the email addresses of the cc-recipients.
            write ('Cc-recipients of the item:');
            displayAddresses(asyncResult);
        }
    }); // End getAsync for cc-recipients.

    // If the item has the bcc field, i.e., item is message,
    // get any bcc-recipients.
    if (bccRecipients) {
        bccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get bcc-recipients of the item completed.
            // Display the email addresses of the bcc-recipients.
            write ('Bcc-recipients of the item:');
            displayAddresses(asyncResult);
        }
                        
        }); // End getAsync for bcc-recipients.
     }
}

// Recipients are in an array of EmailAddressDetails
// objects passed in asyncResult.value.
function displayAddresses (asyncResult) {
    for (var i=0; i<asyncResult.value.length; i++)
        write (asyncResult.value[i].emailAddress);
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Установка получателей


В этом разделе показан пример кода, устанавливающий получателей встречи или сообщения, создаваемых пользователем. При установке получателей все существующие получателей перезаписываются. Как и предыдущий пример, извлекающий получателей в форме создания, этот пример предполагает, что надстройка активируется в формах создания встреч и сообщений. Этот пример сначала проверяет, является ли создаваемый элемент встречей или сообщением, чтобы применить асинхронный метод,  **Recipients.setAsync**, к соответствующим свойствам встречи или сообщения.

При вызове метода  **setAsync** предоставьте массив в качестве входного аргумента параметра _recipients_ в одном из следующих форматов:


- массив строк, являющихся SMTP-адресами;
    
- массив словарей, каждый из которых содержит отображаемое имя и адрес электронной почты, как показано в следующем примере кода;
    
- массив объектов  **EmailAddressDetails** (аналогично массиву, возвращаемому методом **getAsync**).
    
Вы можете предоставить метод обратного вызова как входной аргумент метода  **setAsync**, чтобы выполнять любой код, зависящий от успешной установки получателей, только после соответствующей операции. Кроме того, методу обратного вызова можно передать любые аргументы, используя дополнительный параметр  _asyncContext_. Если вы используете метод обратного вызова, можно обратиться к выходному параметру  _asyncResult_ и использовать свойства **status** и **error** объекта **AsyncResult**, чтобы проверить состояние и наличие ошибок асинхронного вызова.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set recipients of the composed item.
        setRecipients();
    });
}

// Set the display name and email addresses of the recipients of 
// the composed item.
function setRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;

    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method setAsync to set each type of recipients
    // of the composed item. Each time, this example passes a set of
    // names and email addresses to set, and an anonymous 
    // callback function that doesn't take any parameters. 
    toRecipients.setAsync(
        [{
            "displayName":"Graham Durkin", 
            "emailAddress":"graham@contoso.com"
         },
         {
            "displayName" : "Donnie Weinberg",
            "emailAddress" : "donnie@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set to-recipients of the item completed.

            }    
    }); // End to setAsync.


    // Set any cc-recipients.
    ccRecipients.setAsync(
        [{
             "displayName":"Perry Horning", 
             "emailAddress":"perry@contoso.com"
         },
         {
             "displayName" : "Guy Montenegro",
             "emailAddress" : "guy@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set cc-recipients of the item completed.
            }
    }); // End cc setAsync.


    // If the item has the bcc field, i.e., item is message,
    // set bcc-recipients.
    if (bccRecipients) {
        bccRecipients.setAsync(
            [{
                 "displayName":"Lewis Cate", 
                 "emailAddress":"lewis@contoso.com"
             },
             {
                 "displayName" : "Francisco Stitt",
                 "emailAddress" : "francisco@contoso.com"
             }],
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Async call to set bcc-recipients of the item completed.
                    // Do whatever appropriate for your scenario.
                }
        }); // End bcc setAsync.
    }
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```


## Добавление получателей


Если вы не хотите перезаписывать существующих получателей в сообщении или встрече, то вместо использования  **Recipients.setAsync** можно воспользоваться асинхронным методом **Recipients.addAsync**, который добавляет получателей.  **addAsync** работает аналогично **setAsync**, т. е. требует входной аргумент  _recipients_. При необходимости можно предоставить метод обратного вызова и аргументы для него в параметре asyncContext. Затем можно проверить состояние, результат и наличие ошибок асинхронного вызова  **addAsync**, используя выходной параметр  _asyncResult_ метода обратного вызова. Следующий пример проверяет, является ли создаваемый элемент встречей и добавляет двух необходимых участников в встречу.


```js
// Add specified recipients as required attendees of
// the composed appointment. 
function addAttendees() {
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        item.requiredAttendees.addAsync(
        [{
            "displayName":"Kristie Jensen", 
            "emailAddress":"kristie@contoso.com"
         },
         {
            "displayName" : "Pansy Valenzuela",
            "emailAddress" : "pansy@contoso.com"
          }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to add attendees completed.
                // Do whatever appropriate for your scenario.
            }
        }); // End addAsync.
    }
}
```


## Дополнительные ресурсы



- [Считывание и запись данных элемента в форме создания элементов Outlook](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [Считывание и запись данных элемента Outlook в формах просмотра и создания](../outlook/item-data.md)
    
- [Создание надстроек Outlook для форм создания](../outlook/compose-scenario.md)
    
- [Асинхронное программирование в случае надстроек Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    
- [Считывание и запись темы при создании встречи или сообщения в Outlook](../outlook/get-or-set-the-subject.md)
    
- [Вставка данных в основной текст при создании встречи или сообщения в Outlook](../outlook/insert-data-in-the-body.md)
    
- [Считывание и запись расположения при создании встречи в Outlook](../outlook/get-or-set-the-location-of-an-appointment.md)
    
- [Считывание и запись времени при создании встречи в Outlook](../outlook/get-or-set-the-time-of-an-appointment.md)
    
