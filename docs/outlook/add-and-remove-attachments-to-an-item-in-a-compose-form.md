
# Добавление и удаление вложений в форме создания элементов Outlook

Методы [addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) и [addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) можно использовать для прикрепления файла и элемента Outlook к элементу, который создает пользователь. Оба этих метода асинхронные, т. е. они могут выполняться без ожидания завершения действия add-attachment. В зависимости от исходного расположения и размера добавляемого вложения, для завершения асинхронного вызова add-attachment может потребоваться определенное время. Если какие-то задачи зависят от завершения действия, их следует выполнить в методе обратного вызова. Это необязательный метод, который вызывается после завершения отправки вложения. Он принимает объект [AsyncResult](http://dev.outlook.com/reference/add-ins/simple-types.md) как выходной параметр, в котором представлено состояние, ошибка и возвращаемое значение действия add-attachment. Если для обратного вызова требуются дополнительные параметры, их можно указать в необязательном параметре _options.aysncContext_. _options.asyncContext_ может быть любого типа, который ожидается методом обратного вызова.

Например, можно определить _options.asyncContext_ как объект JSON, содержащий одну или несколько пар "ключ-значение", где символ ":" разделяет ключ и значение, а символы "," разделяют пары между собой. Дополнительные примеры [передачи необязательных параметров в асинхронные методы](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) для платформы надстроек Office см. в статье [Асинхронное программирование в надстройках Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md). В следующем примере показано, как использовать параметр **asyncContext** для передачи 2 аргументов методу обратного вызова.




```js
{ asyncContext: { var1: 1, var2: 2} }
```

Успешность обратного вызова асинхронного метода можно проверить с помощью свойств **status** и **error** объекта **AsyncResult**. Если операция вложения завершается успешно, вы можете использовать свойство **AsyncResult.value**, чтобы получить идентификатор вложения. Это целое число, которое можно использовать в дальнейшем, чтобы удалить вложение.


 >**Примечание.** Рекомендуется использовать идентификатор вложения для его удаления, только если та же надстройка добавила вложение в том же сеансе. В Outlook Web App и OWA для устройств идентификатор вложения действителен только в одном сеансе. Сеанс завершается, когда пользователь закрывает надстройку или начинает создавать элемент во встроенной форме и затем продолжает работу в отдельном окне.


## Прикрепление файла

Вы можете прикрепить файл к сообщению или встрече в форме создания, используя метод **addFileAttachmentAsync** и указав URI файла. Если файл защищен, можно добавить соответствующее удостоверение или токен проверки подлинности как параметр строки запроса URI. Exchange вызовет URI, чтобы получить вложение, а веб-службе, которая защищает файл, потребуется использовать токен для проверки подлинности.

Следующий пример JavaScript — это надстройка создания, которая прикрепляет файл picture.png с веб-сервера к создаваемому сообщению или встрече. Метод обратного вызова принимает **asyncResult** как параметр, проверяет состояние вложения и получает его идентификатор, если операция прикрепления была выполнена успешно.




```js
var mailbox;
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

Office.initialize = function () {
    mailbox = Office.context.mailbox;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add the specified file attachment to the item
        // being composed.
        // When the attachment finishes uploading, the
        // callback method is invoked and gets the attachment ID. 
        // You can optionally pass any object that you would  
        // access in the callback method as an argument to  
        // the asyncContext parameter.
        mailbox.item.addFileAttachmentAsync(
            attachmentURI,
            'picture.png',
            { asyncContext: null },
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Get the ID of the attached file.
                    attachmentID = asyncResult.value;
                    write('ID of added attachment: ' + attachmentID);
                }
            });
    });
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Прикрепление элемента Outlook

Вы можете прикрепить элемент Outlook (например, сообщение электронной почты, элемент календаря или контакт) к сообщению или встрече в форме создания, указав идентификатор элемента в веб-службах Exchange (EWS) и вызвав метод **addItemAttachmentAsync**. Вы можете получить идентификатор EWS для элемента сообщения, календаря, контакта или задачи в почтовом ящике пользователя, вызвав метод [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) и используя операцию EWS [FindItem](http://msdn.microsoft.com/en-us/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx). Свойство [item.itemId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.md) также предоставляет идентификатор EWS существующего элемента в форме чтения.

Следующая функция JavaScript, `addItemAttachment`, расширяет приведенный выше пример и добавляет элемент как вложение создаваемого сообщения электронной почты или встречи. В качестве параметра функция принимает идентификатор EWS прикрепляемого элемента. Если операция прикрепления выполнена успешно, функция получает идентификатор вложения для дальнейшей обработки, в том числе удаления этого вложения в том же сеансе.




```js
// Adds the specified item as an attachment to the composed item.
// ID is the EWS ID of the item to be attached.
function addItemAttachment(ID) {
    // When the attachment finishes uploading, the
    // callback method is invoked. Here, the callback
    // method uses only asyncResult as a parameter,
    // and if the attaching succeeds, gets the attachment ID.
    // You can optionally pass any other object you wish to 
    // access in the callback method as an argument to 
    // the asyncContext parameter.
    mailbox.item.addItemAttachmentAsync(
        ID,
        'Welcome email',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```


 >**Примечание.** В Outlook Web App или OWA для устройств надстройку создания можно использовать для прикрепления экземпляра повторяющейся встречи. Однако в расширенном клиенте Outlook попытка прикрепить такой экземпляр приведет к прикреплению ряда повторений (основной встречи).


## Удаление вложения


Вы можете удалить вложение из элемента сообщения или встречи в форме создания, указав соответствующий идентификатор вложения и вызвав метод [removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md). Следует удалять только вложения, которые были добавлены той же надстройкой в том же сеансе. Необходимо убедиться, что идентификатор вложения соответствует действительному вложению, иначе метод вернет ошибку. Как и методы **addFileAttachmentAsync** и **addItemAttachmentAsync**, **removeAttachmentAsync** является асинхронным методом. Следует указать метод обратного вызова, чтобы проверить состояние и наличие ошибок, используя объект выходного параметра **AsyncResult**. Методу обратного вызова также можно передать дополнительные параметры, используя необязательный параметр **asyncContext**, который является объектом JSON, состоящим из пар "ключ-значение".

Следующая функция JavaScript, `removeAttachment`, расширяет предыдущие примеры и удаляет заданное вложение из создаваемого сообщения или встречи. В качестве аргумента функция принимает идентификатор вложения, которое требуется удалить. Идентификатор можно получить после успешного вызова метода **addFileAttachmentAsync** или **addItemAttachmentAsync** и сохранить его для последующего вызова метода **removeAttachmentAsync**.




```js
// Removes the specified attachment from the composed item.
// ID is the Exchange identifier of the attachment to be 
// removed. 
function removeAttachment(ID) {
    // When the attachment is removed, the
    // callback method is invoked. Here, the callback
    // method uses an asyncResult parameter and gets
    // the ID of the removed attachment if the removal
    // succeeds.
    // You can optionally pass any object you wish to 
    // access in the callback method as an argument to 
    // the asyncContext parameter.
    mailbox.item.removeAttachmentAsync(
        ID,
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                write('Removed attachment with the ID: ' + asyncResult.value);
            }
        });
}
```


## Советы по добавлению и удалению вложений


Если ваши надстройки создания добавляют или удаляют вложения, сформируйте код так, чтобы передавать действительный идентификатор вложения при вызове метода remove-attachment и обрабатывать сценарий, когда **AsyncResult.error** возвращает **InvalidAttachmentId**. В зависимости от расположения и размера вложения для завершения операции прикрепления может потребоваться определенное время. Следующий пример сдержит вызов методов **addFileAttachmentAsync**, `write` и **removeAttachmentAsync**. Можно считать, что они выполняются последовательно друг за другом.


```js
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

// Gets the current time in minutes, seconds and milliseconds.
function minutesSecondsMilliSeconds()
{
    var d = new Date();
    return d.getMinutes() + ":" + d.getSeconds() + ":" + d.getMilliseconds();
}

Office.context.mailbox.item.addFileAttachmentAsync(
        attachmentURI,
        'Welcome document',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(1): ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('(2): ' + minutesSecondsMilliSeconds() + ' ' + 
                    'ID of added attachment: ' + attachmentID);
            }
            write ('(3): ' + minutesSecondsMilliSeconds() + ' ' + 
                'Finishing addFileAttachmentAsync callback method.');
        });

write ('(4): ' + minutesSecondsMilliSeconds() + ' ' + 
    'attachmentID is: ' + attachmentID);

Office.context.mailbox.item.removeAttachmentAsync(
        attachmentID,      
        { asyncContext: null },
       function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(5): ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {           
                write('(6): ' + minutesSecondsMilliSeconds() + ' ' + 
                    ID of removed attachment: ' + asyncResult.value);
            }
        });


```

Так как метод **addFileAttachmentAsync** асинхронный, то несмотря на то, что **addFileAttachmentAsync** выполняется перед **removeAttachmentAsync**, вызовы `write` и **removeAttachmentAsync** могут быть выполнены до завершения **addFileAttachmentAsync**. Когда это происходит, `attachmentID` остается **undefined**, а для вызова **removeAttachmentAsync** возникает ошибка, как показано далее:




```
 (4): 46:18:245 attachmentID is: undefined
Error executing code: Sys.ArgumentException: Sys.ArgumentException: Value does not fall within the expected range. Parameter name: attachmentId
 (2): 46:18:255 ID of added attachment: 0
 (3): 46:18:262 Finishing addFileAttachmentAsync callback method.
```

Один из способов предотвратить это — убедиться, что `attachmentID` задан перед вызовом **removeAttachmentAsync**. Другой способ — вызвать **removeAttachmentAsync** в методе обратного вызова **addFileAttachmentAsync**, как в следующем примере:




```js
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

function minutesSecondsMilliSeconds()
{
    var d = new Date();
    return d.getMinutes() + ":" + d.getSeconds() + ":" + d.getMilliseconds();
}

Office.context.mailbox.item.addFileAttachmentAsync(
        attachmentURI,
        'Welcome document',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(1) ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('(2) ' + minutesSecondsMilliSeconds() + ' ' + 
                    'ID of added attachment: ' + attachmentID);

                // Move the write and removeAttachmentAsync calls here 
                // inside the addFileAttachmentAsync callback, after the 
                // attaching has succeeded.
                write ('(4): ' + minutesSecondsMilliSeconds() + ' ' + 
                    'attachmentID is: ' + attachmentID);

                Office.context.mailbox.item.removeAttachmentAsync(
                    attachmentID,
                    { asyncContext: null },
                    function (asyncResult) {
                        if (asyncResult.status == Office.AsyncResultStatus.Failed){
                            write('(5) ' + minutesSecondsMilliSeconds() + ' ' + 
                                asyncResult.error.message);
                        }
                        else {
                            write('(6) ' + minutesSecondsMilliSeconds() + ' ' + 
                                'ID of removed attachment: ' + attachmentID);
                        }
                    });
            }

            write('(3) ' + minutesSecondsMilliSeconds() + ' ' + 
                'Finishing addFileAttachmentAsync callback method.');
        });

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Далее показан пример выходных данных:




```
(2) 49:25:775 ID of added attachment: 1
(4) 49:25:782 attachmentID is: 1
(3) 49:25:783 Finishing addFileAttachmentAsync callback method.
(6) 49:25:789 ID of removed attachment: 1
```

Обратите внимание, что обратный вызов **removeAttachmentAsync** вложен в обратный вызов **addFileAttachmentAsync**. Так как методы **addFileAttachmentAsync** и **removeAttachmentAsync** асинхронные, последняя строка обратного вызова **addFileAttachmentAsync** может выполняться до завершения обратного вызова **removeAttachmentAsync**.


## Дополнительные ресурсы



- [Создание надстроек Outlook для форм создания](../outlook/compose-scenario.md)
    
- [Асинхронное программирование в случае надстроек Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    


