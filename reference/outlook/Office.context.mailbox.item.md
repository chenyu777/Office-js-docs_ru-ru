

# item

## [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).item

Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](Office.context.mailbox.item.md#itemtype-officemailboxenumsitemtype).

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Применимый режим Outlook| Создание или чтение|

### Пример

В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.

```
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Пример

С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.

```
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание|

##### Пример

```
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|
#### dateTimeCreated :Date

Получает дату и время создания элемента. Только в режиме чтения.

##### Тип:

*   Date

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Пример

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### dateTimeModified :Date

Получает дату и время последнего изменения элемента. Только в режиме чтения.

##### Тип:

*   Date

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Пример

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  end :Date|[Time](Time.md)

Получает или задает дату и время окончания встречи.

Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные дату и время клиента можно с помощью метода [`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).

##### Режим чтения

Свойство `end` возвращает объект `Date`.

##### Режим создания

Свойство `end` возвращает объект `Time`.

Если вы задаете время окончания с помощью метода [`Time.setAsync`](Time.md#setasyncdatetime-options-callback), необходимо использовать метод [`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.

##### Тип:

*   Date | [Time](Time.md)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

В примере ниже показано, как с помощью метода [`setAsync`](Time.md#setasyncdatetime-options-callback) объекта `Time` задать время окончания встречи в режиме создания.

```
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

Свойства `from` и [`sender`](Office.context.mailbox.item.md#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.

##### Тип:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|
#### internetMessageId :String

Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.

##### Тип:

*   String

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Пример

```
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Пример

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### (nullable) itemId :String

Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.

Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange. Свойство `itemId` не совпадает с идентификатором записи Outlook.

Свойство `itemId` возвращает значение `null` в режиме создания для элементов, не сохраненных на сервере. Если вам нужен идентификатор элемента, сохраните его на сервере с помощью метода [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Идентификатор будет возвращен в параметре [`AsyncResult.value`](simple-types.md#asyncresult) в функции обратного вызова.

##### Тип:

*   String

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Пример

Следующий код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен на сервере, а из асинхронного результата будет получен идентификатор элемента.

```
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### normalizedSubject :String

Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.

Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Чтобы получить тему элемента с префиксами, используйте свойство [`subject`](Office.context.mailbox.item.md#subject-stringsubject).

##### Тип:

*   String

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Пример

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  notificationMessages :[NotificationMessages](NotificationMessages.md)

Получает сообщения уведомления для элемента.

##### Тип:

*   [NotificationMessages](NotificationMessages.md)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.3|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|
|[Recipients](Recipients.md)|
####  optionalAttendees :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Пример

```
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|
#### sender :[EmailAddressDetails](simple-types.md#emailaddressdetails)

Получает адрес электронной почты отправителя сообщения. Только в режиме чтения.

Свойства [`from`](Office.context.mailbox.item.md#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.

##### Тип:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Пример

```
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

Если вы задаете время начала с помощью метода [`Time.setAsync`](Time.md#setasyncdatetime-options-callback), необходимо использовать метод [`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.

##### Тип:

*   Date | [Time](Time.md)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

В примере ниже с помощью метода [`setAsync`](Time.md#setasyncdatetime-options-callback) объекта `Time` задается время начала встречи в режиме создания.

```
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

```
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```
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

Затем вы можете использовать идентификатор с методом [`removeAttachmentAsync`](Office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение во время того же сеанса.

##### Параметры: removeattachmentasyncattachmentid-options-callback
|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`uri`| String||Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.|
|`attachmentName`| String||Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция| &lt;необязательно&gt;|После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). <br/>После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.<br/>Если вложение добавить не удалось, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.<br/><table class="nested-table"><thead><tr><th>Код ошибки</th><th>Описание</th></tr></thead><tbody><tr><td><code>AttachmentSizeExceeded</code></td><td>Вложение превышает максимальный размер.</td></tr><tr><td><code>FileTypeNotSupported</code></td><td>Расширение вложения не поддерживается.</td></tr><tr><td><code>NumberOfAttachmentsExceeded</code></td><td>Сообщение или встреча содержат слишком много вложений.</td></tr></tbody></table>|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Применимый режим Outlook| Создание|

##### Пример

```
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Применимый режим Outlook| Создание|

##### Пример

В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.

```
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

####  close()

Закрывает текущий создаваемый элемент.

Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.

Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.3|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Применимый режим Outlook| Создание|
#### displayReplyAllForm(formData)

Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.

В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.

Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.

Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`formData`| String &#124; Object|Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.<br/>**ИЛИ**<br/>Объект, который содержит текст или данные вложения и функцию обратного вызова. Объект имеет следующие значения:<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>htmlBody</code></td><td>String</td><td>&lt;необязательно&gt;</td><td>Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</td></tr><tr><td><code>attachments</code></td><td>Array.&lt;Object&gt;</td><td>&lt;необязательно&gt;</td><td>Массив объектов JSON, представляющих собой вложенные файлы или элементы.<br/><br/><strong>Свойства</strong><br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Описание</th></tr></thead><tbody><tr><td><code>type</code></td><td>String</td><td>Указывает тип вложения. Требуемые значения: <code>file</code> для вложенного файла или <code>item</code> для вложенного элемента.</td></tr><tr><td><code>name</code></td><td>String</td><td>Строка, содержащая имя вложения длиной до 255 символов.</td></tr><tr><td><code>url</code></td><td>String</td><td>Используется, только если свойству <code>type</code> присвоено значение <code>file</code>. Универсальный код ресурса (URI) расположения файла.</td></tr><tr><td><code>itemId</code></td><td>String</td><td>Используется, только если свойству <code>type</code> присвоено значение <code>item</code>. Идентификатор вложения EWS. Это строка длиной до 100 символов.</td></tr></tbody></table></td></tr><tr><td><code>callback</code></td><td>function</td><td>&lt;необязательно&gt;</td><td>После выполнения метода функция, переданная в параметре <code>callback</code>, вызывается с помощью параметра <code>asyncResult</code>, который представляет собой объект <a href="simple-types.md#asyncresult"><code>AsyncResult</code></a>. Дополнительные сведения см. в статье <a href="tutorial-asynchronous.html">Использование асинхронных методов</a>.</td></tr></tbody></table>|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Примеры

Приведенный ниже код передает строку в функцию `displayReplyAllForm`.

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

Ответ с пустым текстом сообщения.

```
Office.context.mailbox.item.displayReplyAllForm({});
```

Ответ только с текстом сообщения.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

Ответ с текстом сообщения и вложенным файлом.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

Ответ с текстом сообщения и вложенным элементом.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
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

Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`formData`| String &#124; Object|Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.<br/>**ИЛИ**<br/>Объект, который содержит текст или данные вложения и функцию обратного вызова. Объект имеет следующие значения:<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>htmlBody</code></td><td>String</td><td>&lt;необязательно&gt;</td><td>Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</td></tr><tr><td><code>attachments</code></td><td>Array.&lt;Object&gt;</td><td>&lt;необязательно&gt;</td><td>Массив объектов JSON, представляющих собой вложенные файлы или элементы.<br/><br/><strong>Свойства</strong><br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Описание</th></tr></thead><tbody><tr><td><code>type</code></td><td>String</td><td>Указывает тип вложения. Требуемые значения: <code>file</code> для вложенного файла или <code>item</code> для вложенного элемента.</td></tr><tr><td><code>name</code></td><td>String</td><td>Используется, только если свойству <code>type</code> присвоено значение <code>file</code>. Строка, содержащая имя вложения, длиной до 255 символов.</td></tr><tr><td><code>url</code></td><td>String</td><td>Используется, только если свойству <code>type</code> присвоено значение <code>file</code>. Универсальный код ресурса (URI) расположения файла.</td></tr><tr><td><code>itemId</code></td><td>String</td><td>Используется, только если свойству <code>type</code> присвоено значение <code>item</code>. Идентификатор вложения EWS. Это строка длиной до 100 символов.</td></tr></tbody></table></td></tr><tr><td><code>callback</code></td><td>function</td><td>&lt;необязательно&gt;</td><td>После выполнения метода функция, переданная в параметре <code>callback</code>, вызывается с помощью параметра <code>asyncResult</code>, который представляет собой объект <a href="simple-types.md#asyncresult"><code>AsyncResult</code></a>. Дополнительные сведения см. в статье <a href="tutorial-asynchronous.html">Использование асинхронных методов</a>.</td></tr></tbody></table>|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Примеры

Приведенный ниже код передает строку в функцию `displayReplyForm`.

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

Ответ с пустым текстом сообщения.

```
Office.context.mailbox.item.displayReplyForm({});
```

Ответ только с текстом сообщения.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

Ответ с текстом сообщения и вложенным файлом.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

Ответ с текстом сообщения и вложенным элементом.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Возвращаемое значение:

Тип: [Entities](simple-types.md#entities)

##### Пример

Ниже приведен пример получения доступа к сущностям контактов в текущем элементе.

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>}

Получает массив всех сущностей указанного типа, обнаруженных в выбранном элементе.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`entityType`| [Office.MailboxEnums.EntityType](Office.MailboxEnums.md#.EntityType-string)|Одно из значений перечисления EntityType.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
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

Тип: Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>

##### Пример

В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теме или основном тексте текущего элемента.

```
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Возвращаемое значение:

Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает значение `null`. Если параметр `name` не соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.

Тип: Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>

#### getRegExMatches() → {Object}

Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.

Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующая строка должна содержаться в свойстве элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.

Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения типа `.*` для получения всего текста элемента не всегда приносит ожидаемые результаты. Вместо него используйте метод [`Body.getAsync`](Body.md#getasynccoerciontype-options-callback), чтобы получить весь текст сообщения.

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Возвращаемое значение:

Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.

<dl class="param-type">

<dt>Тип</dt>

<dd>Object</dd>

</dl>

##### Пример

В примере ниже показано, как получить доступ к массиву совпадений для элементов <rule> регулярного выражения `fruits` и `veggies`, которые указаны в манифесте.</rule>

```
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Возвращаемое значение:

Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.

<dl class="param-type">

<dt>Тип</dt>

<dd>Array.<String></dd>

</dl>

##### Пример

```
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Применимый режим Outlook| Создание|

##### Возвращаемое значение:

Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.

<dl class="param-type">

<dt>Тип</dt>

<dd>String</dd>

</dl>

##### Пример

```
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

В приведенном ниже примере показано, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. В этом примере также показано, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода используются метод `CustomProperties.get` для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` для записи настраиваемого свойства `otherProp` и метод `saveAsync` для сохранения настраиваемых свойств.

```
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
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Применимый режим Outlook| Создание|

##### Пример

Указанный ниже код удаляет вложение с идентификатором "0".

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  saveAsync([options], callback)

Асинхронно сохраняет элемент.

При вызове этот метод сохраняет текущее сообщения в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция||После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult).

После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.3|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Применимый режим Outlook| Создание|

##### Примеры

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

Ниже приведен пример параметра `result`, переданного в функцию обратного вызова. Свойство `value` содержит идентификатор элемента.

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  setSelectedDataAsync(data, [options], callback)

Асинхронно вставляет данные в текст или тему сообщения.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`data`| String||Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</td></tr><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;необязательно&gt;</td><td>Если задано значение <code>text</code>, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</td></tr></tbody></table><p>Если задано значение <code>html</code>, а поле (не тема) поддерживает HTML , в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка <code>InvalidDataFormat</code>.</p><p>Если свойство <code>coercionType</code> не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.|</p>|
|`callback`| function||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). |

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.2|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Применимый режим Outlook| Создание|

##### Пример

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
