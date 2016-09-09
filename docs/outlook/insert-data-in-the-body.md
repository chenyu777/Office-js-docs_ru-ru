
# Вставка данных в основной текст при создании встречи или сообщения в Outlook

Вы можете использовать асинхронные методы ([Body.getAsync](../../reference/outlook/Body.md), [Body.getTypeAsync](../../reference/outlook/Body.md), [Body.prependAsync](../../reference/outlook/Body.md), [Body.setAsync](../../reference/outlook/Body.md) и [Body.setSelectedDataAsync](../../reference/outlook/Body.md)), чтобы получить тип основного текста и вставить данные в основной текст элемента встречи или сообщения, создаваемых пользователем. Эти асинхронные методы доступны только для надстроек создания. Чтобы использовать эти методы, необходимо настроить манифест для активации надстройки в Outlook, как описано в статье [Создание надстроек Outlook для форм создания](../outlook/compose-scenario.md).

В Outlook пользователь может создавать сообщения (текстовые, а также в формате HTML и RTF) и встречи (в формате HTML). Перед вставкой всегда необходимо сначала проверить поддерживаемый формат элемента, вызвав метод **getTypeAsync**, так как может понадобиться выполнить дополнительные действия. Значение, которое возвращает метод **getTypeAsync**, зависит от исходного формата элемента, а также от того, поддерживают ли операционная система устройства и узел редактирование в формате HTML (1). Затем соответствующим образом укажите параметр _coercionType_ метода **prependAsync** или **setSelectedDataAsync** (2) для вставки данных, как показано в таблице ниже. Если вы не укажете аргумент, методы **prependAsync** и **setSelectedDataAsync** поведут себя так, как будто данные вставляются в текстовом формате.



|**Данные для вставки**|**Формат элемента, возвращенный методом getTypeAsync**|**Необходимый параметр coercionType**|
|:-----|:-----|:-----|
|Текст|Текст (1)|Текст|
|HTML|Текст (1)|Текст (2)|
|Текст|HTML|Текст или HTML|
|HTML|HTML |HTML|

1.  На планшетах и смартфонах метод **getTypeAsync** возвращает **Office.MailboxEnums.BodyType.Text** в формате HTML, если операционная система или узел не поддерживает редактирование элемента, изначально созданного в этом формате.

2.  Если вставляются данные HTML, а метод **getTypeAsync** возвращает текстовый тип, преобразуйте данные в текст и вставьте их, используя в качестве **coercionType** _Office.MailboxEnums.BodyType.Text_. Если просто вставить данные HTML с помощью типа приведения text, узел отобразит HTML-теги в виде текста. Если вы попытаетесь вставить данные HTML, используя в качестве **coercionType** _Office.MailboxEnums.BodyType.Html_, возвратится ошибка.

Как и большинство асинхронных методов в API JavaScript для Office,  _coercionType_,  **getTypeAsync** и **prependAsync** принимают кроме параметра **setSelectedDataAsync** другие входные параметры (необязательные). Дополнительные сведения об указании последних см. в разделе [Передача дополнительных параметров в асинхронные методы](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) статьи [Асинхронное программирование в надстройках для Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md).


## Вставка данных в текущей позиции курсора


В этом раздел представлен пример кода, который использует  **getTypeAsync** для проверки типа текста создаваемого элемента, а затем вызывает метод **setSelectedDataAsync** для вставки данных в текущем положении курсора.

Вы можете передать метод обратного вызова и необязательные входные параметры в  **getTypeAsync**. Тогда состояние и результаты будут возвращены в параметре вывода  _asyncResult_. Если метод выполнен успешно, вы получите тип текста элемента в свойстве [AsyncResult.value](../../reference/shared/asyncresult.status.md), значение которого — "text" или "html".

Необходимо передать строку данных как входной параметр метода  **setSelectedDataAsync**. В зависимости от типа текста элемента можно указать эту строку в виде текста или HTML соответственно. Как было сказано ранее, при необходимости тип вставляемых данных можно указать в параметре  _coercionType_. Кроме того, вы можете предоставить метод обратного вызова и его параметры в качестве дополнительных входных параметров.

Если пользователь не разместил курсор в тексте элемента,  **setSelectedDataAsync** вставляет данные в начало текста. Если пользователь выбрал текст в элементе, **setSelectedDataAsync** заменяет выбранный текст указанными вами данными. Обратите внимание, что вызов **setSelectedDataAsync** может завершиться ошибкой, если пользователь одновременно меняет позицию курсора при создании элемента. Максимальное число символов, которые можно вставить за один раз — 1 000 000.

В этом примере предполагается, что в манифесте задано правило, которое активирует надстройку в форме создания встречи или сообщения, как показано ниже.




```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set data in the body of the composed item.
        setItemBody();
    });
}


// Get the body type of the composed item, and set data in 
// in the appropriate data type in the item body.
function setItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(result.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Set data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of setSelectedDataAsync.
                    item.body.setSelectedDataAsync(
                        '<b> Kindly note we now open 7 days a week.</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.setSelectedDataAsync(
                        ' Kindly note we now open 7 days a week.',
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Вставка данных в начале текста элемента


Кроме того, с помощью метода  **prependAsync** можно вставить данные в начале текста элемента независимо от положения курсора. Помимо точки вставки, методы **prependAsync** и **setSelectedDataAsync** работают одинаково:


- Если вы добавляете HTML-данные в начало текста сообщения, сначала следует проверить тип текста сообщения, чтобы предотвратить вставку HTML-данных в текстовое сообщение.
    
- Предоставьте следующие входные параметры для метода  **prependAsync**: строка данных в текстовом формате или формате HTML и, при необходимости, формат вставляемых данных, метод обратного вызова и его параметры.
    
- Максимальное число символов, которые можно вставить в начало за один раз — 1 000 000.
    
Следующий код JavaScript является частью примера надстройки, которая активируется в формах создания встреч и сообщений. Пример вызывает метод  **getTypeAsync** для проверки типа текста элемента, вставляет HTML-данные в начало элемента, если это встреча или HTML-сообщение, а в противном случае вставляет данные в текстовом формате.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Insert data in the top of the body of the composed 
        // item.
        prependItemBody();
    });
}

// Get the body type of the composed item, and prepend data  
// in the appropriate data type in the item body.
function prependItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Prepend data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of prependAsync.
                    item.body.prependAsync(
                        '<b>Greetings!</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.prependAsync(
                        'Greetings!',
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
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
    
- [Считывание и запись расположения при создании встречи в Outlook](../outlook/get-or-set-the-location-of-an-appointment.md)
    
- [Считывание и запись времени при создании встречи в Outlook](../outlook/get-or-set-the-time-of-an-appointment.md)
    
