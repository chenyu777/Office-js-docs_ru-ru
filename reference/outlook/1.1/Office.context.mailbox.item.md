

# item

## [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). item

Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](Office.context.mailbox.item.md#itemtype-officemailboxenumsitemtype).

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Применимый режим Outlook| Создание или чтение|

### Пример

В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.

```JavaScript
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### Элементы

#### attachments :Array.<[AttachmentDetails](simple-types.md#attachmentdetails)>

Получает массив вложений для элемента. Только в режиме чтения.

##### Тип:

*   Array.<[AttachmentDetails](simple-types.md#attachmentdetails)>

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Пример

С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.

```JavaScript
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  bcc :[Recipients](Recipients.md)

Получает или задает получателей скрытой копии сообщения. Только в режиме создания.

##### Тип:

*   [Recipients](Recipients.md)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание|

##### Пример

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  body :[Body](Body.md)

Получает объект, предоставляющий методы для работы с основным текстом элемента.

##### Тип:

*   [Body](Body.md)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|
####  cc :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Recipients](Recipients.md)

Получает или задает получателей копии сообщения.

##### Режим чтения

Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.

##### Режим создания

Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для работы с получателями, которые указаны в строке **Копия** сообщения.

##### Тип:

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  (nullable) conversationId :String

Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.

Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.

Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.

##### Тип:

*   String

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|
#### dateTimeCreated :Date

Получает дату и время создания элемента. Только в режиме чтения.

##### Тип:

*   Date

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Пример

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### dateTimeModified :Date

Получает дату и время последнего изменения элемента. Только в режиме чтения.

##### Тип:

*   Date

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Пример

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  end :Date|[Time](Time.md)

Получает или задает дату и время окончания встречи.

Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные дату и время клиента можно с помощью метода [`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).

##### Режим чтения

Свойство `end` возвращает объект `Date`.

##### Режим создания

Свойство `end` возвращает объект `Time`.

Если время окончания задается с помощью метода [`Time.setAsync`](Time.md#setasync), необходимо использовать метод [`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date), чтобы изменить местное время в клиенте на время в формате UTC для сервера.

##### Тип:

*   Date | [Time](Time.md)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

В примере ниже показано, как с помощью метода [`setAsync`](Time.md#setasync) объекта `Time` задать время окончания встречи в режиме создания.

```JavaScript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### from :[EmailAddressDetails](simple-types.md#emailaddressdetails)

Получает адрес электронной почты отправителя сообщения. Только в режиме чтения.

Свойства `from` и [`sender`](Office.context.mailbox.item.md#sender) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.

##### Тип:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|
#### internetMessageId :String

Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.

##### Тип:

*   String

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Пример

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### itemClass :String

Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.

Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.

| Тип | Описание | Класс элемента |
| --- | --- | --- |
| Элементы встречи | Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`. | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| Элементы сообщения | Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения. | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.

##### Тип:

*   String

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Пример

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### (nullable) itemId :String

Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.

Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange. Свойство `itemId` не совпадает с идентификатором записи Outlook.

Свойство `itemId` возвращает значение `null` в режиме создания для элементов, не сохраненных на сервере. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](Office.context.mailbox.item.md#saveAsync) можно сохранить элемент на сервере. При этом в параметре [`AsyncResult.value`](simple-types.md#asyncresult) в функции обратного вызова возвращается идентификатор элемента.

##### Тип:

*   String

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Пример

Следующий код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен на сервере, а из асинхронного результата будет получен идентификатор элемента.

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  itemType :[Office.MailboxEnums.ItemType](Office.MailboxEnums.md#itemtype-string)

Получает тип элемента, который представляет экземпляр.

Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.

##### Тип:

*   [Office.MailboxEnums.ItemType](Office.MailboxEnums.md#itemtype-string)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  location :String|[Location](Location.md)

Получает или задает место встречи.

##### Режим чтения

Свойство `location` возвращает строку, содержащую сведения о месте встречи.

##### Режим создания

Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.

##### Тип:

*   String | [Location](Location.md)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### normalizedSubject :String

Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.

Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с префиксами используйте свойство [`subject`](Office.context.mailbox.item.md#subject).

##### Тип:

*   String

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Пример

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  optionalAttendees :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Recipients](Recipients.md)

Получает или задает список адресов электронной почты необязательных участников.

##### Режим чтения

Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.

##### Режим создания

Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения и задания необязательных участников собрания.

##### Тип:

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### organizer :[EmailAddressDetails](simple-types.md#emailaddressdetails)

Получает адрес электронной почты организатора указанного собрания. Только в режиме чтения.

##### Тип:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Пример

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  requiredAttendees :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Recipients](Recipients.md)

Получает или задает список адресов электронной почты обязательных участников.

##### Режим чтения

Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.

##### Режим создания

Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить и задать сведения об обязательных участниках собрания.

##### Тип:

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### resources :[EmailAddressDetails](simple-types.md#emailaddressdetails)

Получает ресурсы, необходимые для встречи. Только в режиме чтения.

##### Тип:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|
#### sender :[EmailAddressDetails](simple-types.md#emailaddressdetails)

Получает адрес электронной почты отправителя сообщения. Только в режиме чтения.

Свойства [`from`](Office.context.mailbox.item.md#from) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.

##### Тип:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Пример

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  start :Date|[Time](Time.md)

Получает или задает дату и время начала встречи.

Свойство `start` представлено в виде значения даты и времени в формате UTC. Его можно преобразовать в местные дату и время клиента с помощью метода [`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).

##### Режим чтения

Свойство `start` возвращает объект `Date`.

##### Режим создания

Свойство `start` возвращает объект `Time`.

Если время начала задается с помощью метода [`Time.setAsync`](Time.md#setasync), необходимо использовать метод [`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date), чтобы изменить местное время в клиенте на время в формате UTC для сервера.

##### Тип:

*   Date | [Time](Time.md)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

В следующем примере с помощью метода [`setAsync`](Time.md#setasync) объекта `Time` задается время начала встречи в режиме создания.

```JavaScript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  subject :String|[Subject](Subject.md)

Получает или задает описание, которое отображается в поле темы элемента.

Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.

##### Режим чтения

Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](Office.context.mailbox.item.md#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.

```
var subject = Office.context.mailbox.item.subject;
```

##### Режим создания

Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### Тип:

*   String | [Subject](Subject.md)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|
####  to :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Recipients](Recipients.md)

Получает или задает получателей сообщения электронной почты.

##### Режим чтения

Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.

##### Режим создания

Свойство `to` возвращает объект `Recipients`, предоставляющий методы для работы с получателями в строке **Кому** сообщения.

##### Тип:

*   Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)> | [Recipients](Recipients.md)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### Методы

####  addFileAttachmentAsync(uri, attachmentName, [options], [callback])

Добавляет файл в сообщение или встречу в качестве вложения.

Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.

Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](Office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.

##### Parameters:removeattachmentasyncattachmentid-options-callback

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`uri`| String||Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.|
|`attachmentName`| String||Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция| &lt;необязательно&gt;|После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). <br/>После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.<br/>Если вложение добавить не удалось, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.<br/><table class="nested-table"><thead><tr><th>Код ошибки</th><th>Описание</th></tr></thead><tbody><tr><td><code>AttachmentSizeExceeded</code></td><td>Вложение превышает максимальный размер.</td></tr><tr><td><code>FileTypeNotSupported</code></td><td>Расширение вложения не поддерживается.</td></tr><tr><td><code>NumberOfAttachmentsExceeded</code></td><td>Сообщение или встреча содержат слишком много вложений.</td></tr></tbody></table>|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Применимый режим Outlook| Создание|

##### Пример

```JavaScript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

####  addItemAttachmentAsync(itemId, attachmentName, [options], [callback])

Добавляет к сообщению или встрече элемент Exchange, например сообщение, в виде вложения.

С помощью метода `addItemAttachmentAsync` в элемент в форме создания вкладывается элемент с указанным идентификатором Exchange. Если указан метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о методе обратного вызова.

Затем вы можете использовать идентификатор с методом [`removeAttachmentAsync`](Office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение во время того же сеанса.

Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`itemId`| String||Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.|
|`attachmentName`| String||Тема вкладываемого элемента. Максимальная длина — 255 символов.|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция| &lt;необязательно&gt;|После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). <br/>После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.<br/>Если добавить вложение не удалось, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.<br/><table class="nested-table"><thead><tr><th>Код ошибки</th><th>Описание</th></tr></thead><tbody><tr><td><code>NumberOfAttachmentsExceeded</code></td><td>Сообщение или встреча содержат слишком много вложений.</td></tr></tbody></table>|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Применимый режим Outlook| Создание|

##### Пример

В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.

```JavaScript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

#### displayReplyAllForm(formData)

Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.

В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.

Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.

> **ПРИМЕЧАНИЕ.** В наборе требований 1.1 не поддерживается возможность включения вложений в вызов `displayReplyAllForm`. Поддержка вложений была добавлена в `displayReplyAllForm` в наборах требований 1.2 и более поздних версий.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`formData`| String &#124; Object|Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.<br/>**ИЛИ**<br/>Объект, который содержит текст или данные вложения и функцию обратного вызова. Объект имеет следующие значения:<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>htmlBody</code></td><td>String</td><td>&lt;необязательно&gt;</td><td>Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Длина строки ограничена 32 символами.</td></tr><tr><td><code>callback</code></td><td>функция</td><td>&lt;необязательно&gt;</td><td>После выполнения метода функция, переданная в параметре <code>callback</code>, вызывается с помощью параметра <code>asyncResult</code>, который представляет собой объект <a href="simple-types.md#asyncresult"><code>AsyncResult</code></a>. Дополнительные сведения см. в статье <a href="tutorial-asynchronous.html">Использование асинхронных методов</a>.</td></tr></tbody></table>|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Примеры

Приведенный ниже код передает строку в функцию `displayReplyAllForm`.

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

Ответ с пустым текстом сообщения.

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

Ответ только с текстом сообщения.

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

Ответ с текстом сообщения и вложенным файлом.

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### displayReplyForm(formData)

Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.

В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.

Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.

> **ПРИМЕЧАНИЕ.** В наборе требований 1.1 не поддерживается возможность включения вложений в вызов `displayReplyForm`. Поддержка вложений была добавлена в `displayReplyForm` в наборах требований 1.2 и более поздних версий.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`formData`| String &#124; Object|Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.<br/>**ИЛИ**<br/>Объект, который содержит текст или данные вложения и функцию обратного вызова. Объект имеет следующие значения:<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>htmlBody</code></td><td>String</td><td>&lt;необязательно&gt;</td><td>Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Длина строки ограничена 32 символами.</td></tr><tr><td><code>callback</code></td><td>функция</td><td>&lt;необязательно&gt;</td><td>После выполнения метода функция, переданная в параметре <code>callback</code>, вызывается с помощью параметра <code>asyncResult</code>, который представляет собой объект <a href="simple-types.md#asyncresult"><code>AsyncResult</code></a>. Дополнительные сведения см. в статье <a href="tutorial-asynchronous.html">Использование асинхронных методов</a>.</td></tr></tbody></table>|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Примеры

Приведенный ниже код передает строку в функцию `displayReplyForm`.

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

Ответ с пустым текстом сообщения.

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

Ответ только с текстом сообщения.

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

Ответ с текстом сообщения и вложенным файлом.

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### getEntities() → {[Entities](simple-types.md#entities)}

Получает сущности, обнаруженные в выбранном элементе.

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Возвращаемое значение:

Тип: [Entities](simple-types.md#entities)

##### Пример

Ниже приведен пример получения доступа к сущностям контактов в текущем элементе.

```JavaScript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>}

Получает массив всех сущностей указанного типа, обнаруженных в выбранном элементе.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`entityType`| [Office.MailboxEnums.EntityType](Office.MailboxEnums.md#.entitytype-string)|Одно из значений перечисления EntityType.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Применимый режим Outlook| Чтение|

##### Возвращаемое значение:

Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL. Если в элементе отсутствуют сущности указанного типа, метод возвращает пустой массив. В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.

Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.

| Значение параметра `entityType` | Тип объектов в возвращаемом массиве | Необходимый уровень разрешений |
| --- | --- | --- |
| `Address` | String | **Restricted** |
| `Contact` | Contact | **ReadItem** |
| `EmailAddress` | String | **ReadItem** |
| `MeetingSuggestion` | MeetingSuggestion | **ReadItem** |
| `PhoneNumber` | PhoneNumber | **Restricted** |
| `TaskSuggestion` | TaskSuggestion | **ReadItem** |
| `URL` | String | **Restricted** |

Тип:  Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))></dd>


##### Пример

В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теме или основном тексте текущего элемента.

```JavaScript
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>}

Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.

Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](https://msdn.microsoft.com/en-us/library/office/fp161166.aspx) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`name`| String|Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Возвращаемое значение:

Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает значение `null`. Если параметр `name` не соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.


Тип: Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>


#### getRegExMatches() → {Object}

Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.

Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующая строка должна содержаться в свойстве элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.

Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения типа `.*` для получения всего текста элемента не всегда приносит ожидаемые результаты. Вместо него используйте метод [`Body.getAsync`](Body.md#getAsync).

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Возвращаемое значение:

Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.

<dl class="param-type">

<dt>Тип</dt>

<dd>Object</dd>

</dl>

##### Пример

В примере ниже показано, как получить доступ к массиву совпадений для элементов <rule> регулярного выражения `fruits` и `veggies`, которые указаны в манифесте.</rule>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### getRegExMatchesByName(name) → (nullable) {Array.<String>}

Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.

Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.

Если вы указываете правило `ItemHasRegularExpressionMatch` для текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения типа `.*` для получения всего текста элемента не всегда приносит ожидаемые результаты.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`name`| String|Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Возвращаемое значение:

Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.

<dl class="param-type">

<dt>Тип</dt>

<dd>Array.<String></dd>

</dl>

##### Пример

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  getSelectedDataAsync(coercionType, [options], callback) → {String}

Асинхронно возвращает данные, выбранные в теме или тексте сообщения.

Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`coercionType`| [Office.CoercionType](Office.md#coerciontype-string)||Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция||После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult).

Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`. Для доступа к исходному свойству, на основе которого созданы выбранные данные, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Применимый режим Outlook| Создание|

##### Возвращаемое значение:

Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.

<dl class="param-type">

<dt>Тип</dt>

<dd>String</dd>

</dl>

##### Пример

```JavaScript
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

####  loadCustomPropertiesAsync(callback, [userContext])

Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.

Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`callback`| function||После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult).

Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](CustomProperties.md) в свойстве `asyncResult.value`. Этот объект позволяет получить, задать и удалить настраиваемые свойства из элемента, а также сохранить изменения, внесенные в настраиваемое свойство, на сервере.| |`userContext`| Объект| &lt;необязательно&gt;|В функции обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ. Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

В приведенном ниже примере показано, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. В этом примере также показано, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода используются метод `CustomProperties.get` для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` для записи настраиваемого свойства `otherProp` и метод `saveAsync` для сохранения настраиваемых свойств.

```JavaScript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  removeAttachmentAsync(attachmentId, [options], [callback])

Удаляет вложение из сообщения или встречи.

Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`attachmentId`| String||Идентификатор удаляемого вложения. Максимальная длина строки — 100 символов.|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция| &lt;необязательно&gt;|После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). <br/>Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.<br/><table class="nested-table"><thead><tr><th>Код ошибки</th><th>Описание</th></tr></thead><tbody><tr><td><code>InvalidAttachmentId</code></td><td>Идентификатор вложения не существует.</td></tr></tbody></table>|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Применимый режим Outlook| Создание|

##### Пример

Указанный ниже код удаляет вложение с идентификатором "0".

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```
