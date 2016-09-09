

# item

## [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). item

Пространство имен item`item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен item, используя свойство itemType. Пространство имен item`item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен item, используя свойство [itemType](Office.context.mailbox.item.md#itemtype-officemailboxenumsitemtype).

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Применимый режим Outlook| Создание или чтение|

### Пример

В следующем примере кода JavaScript показано, как получить доступ к свойству subject`subject` текущего элемента в Outlook.

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

Получает массив вложений для элемента. Только в режиме чтения. Read mode only.

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

Получает или задает получателей скрытой копии сообщения. Только в режиме создания. Указывает тип сущности. Только в режиме создания.

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

*   [Основной текст](Body.md)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|
####  cc :Array.<[EmailAddressDetails](simple-types.md#emailaddressdetails)>|[Recipients](Recipients.md)

Получает или задает получателей копии сообщения.

##### Режим чтения

Свойство cc`cc` возвращает массив, содержащий объект EmailAddressDetails`EmailAddressDetails` для каждого получателя в строке **Копии** сообщения. Коллекция может включать не более 100 элементов. The collection is limited to a maximum of 100 members.

##### Режим создания

Свойство cc`cc` возвращает объект Recipients`Recipients`, предоставляющий методы для работы с получателями в строке **Копии** сообщения.

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

В случае нового элемента в форме создания это свойство принимает значение null. Это свойство имеет значение NULL для нового элемента в форме создания. Если пользователь задаст тему и сохранит элемент, свойство conversationId`conversationId` вернет значение.

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

####  end :Date Time

Получает или задает дату и время окончания встречи.

Свойство start`end` выражается в качестве значения даты и времени в формате UTC. Вы можете использовать метод convertToLocalClientTime, чтобы преобразовать значение в местные дату и время клиента. Свойство end выражается в качестве значения даты и времени в формате UTC. Вы можете использовать метод  convertToLocalClientTime , чтобы преобразовать значение свойства end в местные дату и время клиента.

##### Режим чтения

The `end` property returns a `Date` object.

##### Режим создания

The `end` property returns a `Time` object.

Если для задания времени окончания используется метод  Time.setAsync , вам следует с помощью метода  convertToUtcClientTime  преобразовать местное время в клиенте на время в формате UTC для сервера.

##### Тип:

*   Date | [Time](Time.md)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

В следующем примере с помощью метода  setAsync  объекта Time задается время окончания встречи в режиме создания.

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

Получает адрес электронной почты отправителя сообщения. Read mode only.

Свойства  from  и sender представляют одно лицо, если сообщение не отправлено делегатом. Если сообщение отправлено делегатом, то свойство from представляет делегирующее лицо, а свойство sender представляет делегата. Свойства from`from` и sender представляют одно лицо, если сообщение не отправлено делегатом. Если сообщение отправлено делегатом, то свойство from представляет делегирующее лицо, а свойство sender представляет делегата.

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

Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения. Read mode only.

Свойство itemClass`itemClass` указывает класс сообщения выбранного элемента. Далее приводятся классы сообщения по умолчанию для элемента сообщения или встречи. Свойство itemClass указывает класс сообщения выбранного элемента. Далее приводятся классы сообщения по умолчанию для элемента сообщения или встречи.

| Тип | Описание | Класс элемента |
| --- | --- | --- |
| Элементы встречи | Это элементы календаря класса элемента IPM.Appointment`IPM.Appointment` или IPM.Appointment.Occurence`IPM.Appointment.Occurence`. | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| Элементы сообщения | Сюда входят электронные сообщения с классом сообщения по умолчанию IPM.Note`IPM.Note`, а также приглашения на собрания, ответы на приглашения и отмены собраний, использующие IPM.Schedule.Meeting`IPM.Schedule.Meeting` в качестве базового класса сообщения. | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

Можно создавать настраиваемые классы сообщения, расширяющие класс сообщения по умолчанию, например настраиваемый класс сообщения о встрече IPM.Appointment.Contoso`IPM.Appointment.Contoso`.

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

Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения. Read mode only.

Идентификатор, возвращаемый свойством itemId`itemId`, совпадает с идентификатором элемента веб-служб Exchange. Свойство itemId не совпадает с кодом записи Outlook. The `itemId` property is not identical to the Outlook Entry ID.

The `itemId` property returns `null` in compose mode for items that have not been saved to the server. Свойство itemId возвращает null в режиме создания для элементов, которые не сохранены на сервере. Если требуется идентификатор элемента, метод  saveAsync  позволяет сохранить элемент на сервере, что вернет идентификатор элемента в параметре  AsyncResult.value  в функции обратного вызова.

##### Тип:

*   String

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Пример

The following code checks for the presence of an item identifier. Следующий код проверяет наличие идентификатора элемента. Если свойство itemId`itemId` возвращает null`null` или undefined`undefined`, элемент будет сохранен на сервере, а из асинхронного результата будет получен идентификатор элемента.

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

Свойство itemType`itemType` возвращает одно из значений перечисления ItemType`ItemType`, которое указывает, является ли экземпляр объекта item`item` сообщением или собранием.

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

####  location :String Location

Получает или задает место встречи.

##### Режим чтения

Свойство location`location` возвращает строку, содержащую место встречи.

##### Режим создания

Свойство location`location` возвращает объект Location`Location`, предоставляющий методы, которые используются для получения и задания места встречи.

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

Получает тему элемента со всеми удаленными префиксами (включая RE: и FWD:). Только в режиме чтения. Read mode only.

Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как RE:`RE:` и FW:`FW:`), добавляемыми почтовыми программами. Для получения темы элемента с префиксами используйте свойство subject. To get the subject of the item with the prefixes intact, use the [`subject`](Office.context.mailbox.item.md#subject) property.

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

Свойство optionalAttendees`optionalAttendees` возвращает массив, содержащий объект EmailAddressDetails`EmailAddressDetails` для каждого необязательного участника собрания.

##### Режим создания

Свойство optionalAttendees`optionalAttendees` возвращает объект Recipients`Recipients`, который предоставляет методы для получения и задания необязательных участников собрания.

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

Получает адрес электронной почты организатора указанного собрания. Read mode only.

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

Свойство requiredAttendees`requiredAttendees` возвращает массив, содержащий объект EmailAddressDetails`EmailAddressDetails` для каждого обязательного участника собрания.

##### Режим создания

Свойство requiredAttendees`requiredAttendees` возвращает объект Recipients`Recipients`, который предоставляет методы для получения и задания обязательных участников собрания.

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

Получает ресурсы, необходимые для встречи. Read mode only.

##### Тип:

*   [EmailAddressDetails](simple-types.md#emailaddressdetails)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|
#### sender :[EmailAddressDetails](simple-types.md#emailaddressdetails)

Получает адрес электронной почты отправителя сообщения. Read mode only.

Свойства  from  и sender представляют одно лицо, если сообщение не отправлено делегатом. Если сообщение отправлено делегатом, то свойство from представляет делегирующее лицо, а свойство sender представляет делегата. Свойства from`from` и sender представляют одно лицо, если сообщение не отправлено делегатом. Если сообщение отправлено делегатом, то свойство from представляет делегирующее лицо, а свойство sender представляет делегата.

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

####  start :Date Time

Получает или задает дату и время начала встречи.

Свойство start`start` выражается в качестве значения даты и времени в формате UTC. Вы можете использовать метод convertToLocalClientTime, чтобы преобразовать значение в местные дату и время клиента. Свойство start выражается в качестве значения даты и времени в формате UTC. Вы можете использовать метод  convertToLocalClientTime , чтобы преобразовать значение в местные дату и время клиента.

##### Режим чтения

The `start` property returns a `Date` object.

##### Режим создания

The `start` property returns a `Time` object.

Если для задания времени начала используется метод  Time.setAsync , вам следует с помощью метода  convertToUtcClientTime  преобразовать местное время в клиенте на время в формате UTC для сервера.

##### Тип:

*   Date | [Time](Time.md)

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

В следующем примере с помощью метода  setAsync  объекта Time задается время начала встречи в режиме создания.

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

####  subject :String Subject

Получает или задает описание, которое отображается в поле темы элемента.

Свойство subject`subject` получает или задает всю тему элемента для отправки с почтового сервера.

##### Режим чтения

The `subject` property returns a string. Свойство subject возвращает строку. Свойство  normalizedSubject  позволяет получить тему без начальных префиксов, таких как RE: и FW:.

```
var subject = Office.context.mailbox.item.subject;
```

##### Режим создания

Свойство subject`subject` возвращает объект Subject`Subject`, который предоставляет методы для получения и задания темы.

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

Свойство to`to` возвращает массив, содержащий объект EmailAddressDetails`EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов. The collection is limited to a maximum of 100 members.

##### Режим создания

Свойство to`to` возвращает объект Recipients`Recipients`, предоставляющий методы для работы с получателями в строке **Кому** сообщения.

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

Метод addFileAttachmentAsync`addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.

Идентификатор можно последовательно использовать с методом  removeAttachmentAsync , чтобы удалить вложение, добавленное во время текущего сеанса.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`uri`| String||Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов. The maximum length is 2048 characters.|
|`attachmentName`| String||Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов. Тема вкладываемого элемента. Максимальная длина — 255 символов.|
|`options`| Объект| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Объект</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). <br/>On success, the attachment identifier will be provided in the `asyncResult.value` property.<br/>Если передать вложение не удается, объект asyncResult`asyncResult` будет содержать объект Error`Error` с описанием ошибки.<br/><table class="nested-table"><thead><tr><th>Код ошибки</th><th>Описание</th></tr></thead><tbody><tr><td><code>AttachmentSizeExceeded</code></td><td>Размер вложения превышает допустимый.</td></tr><tr><td><code>FileTypeNotSupported</code></td><td>Вложение имеет неподдерживаемое расширение.</td></tr><tr><td><code>NumberOfAttachmentsExceeded</code></td><td>Сообщение содержит слишком много вложений.</td></tr></tbody></table>|

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

Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.

The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. Метод addItemAttachmentAsync`asyncResult` вкладывает элемент с указанным идентификатором Exchange в элемент в форме создания. Если указан метод обратного вызова, этот метод вызывается с помощью параметра asyncResult, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр options для передачи сведений о состоянии в метод обратного вызова. You can use the `options` parameter to pass state information to the callback method, if needed.

Идентификатор можно последовательно использовать с методом  removeAttachmentAsync , чтобы удалить вложение, добавленное во время текущего сеанса.

Если ваша надстройка Office выполняется в Outlook Web App, метод addItemAttachmentAsync`addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, поскольку оно не поддерживается.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`itemId`| String||Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов. Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.|
|`attachmentName`| String||Тема вкладываемого элемента. Максимальная длина — 255 символов. Тема вкладываемого элемента. Максимальная длина — 255 символов.|
|`options`| Объект| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Объект</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). <br/>On success, the attachment identifier will be provided in the `asyncResult.value` property.<br/>Если передать вложение не удается, объект asyncResult`asyncResult` будет содержать объект Error`Error` с описанием ошибки.<br/><table class="nested-table"><thead><tr><th>Код ошибки</th><th>Описание</th></tr></thead><tbody><tr><td><code>NumberOfAttachmentsExceeded</code></td><td>Сообщение содержит слишком много вложений.</td></tr></tbody></table>|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Применимый режим Outlook| Создание|

##### Пример

В следующем примере существующий элемент Outlook добавляется в виде вложения с именем My Attachment`My Attachment`.

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

Если любой строковый параметр превышает указанные для него ограничения, displayReplyForm`displayReplyAllForm` возвращает исключение.

Если в параметре formData.attachments`formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать все вложения и вложить их в форму ответа. Если не удается добавить какие-либо вложения, в форме пользовательского интерфейса отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`formData`| String &#124; Object|Строка, содержащая текст и HTML, которая представляет основной текст формы ответа. Эта строка должна быть не более 32 КБ. The string is limited to 32 KB.<br/>**OR**<br/>An object that contains body or attachment data and a callback function. Поведение формы при ее открытии определяется следующим образом:<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>htmlBody</code></td><td>String</td><td>&lt;необязательно&gt;</td><td>Строка, содержащая текст и HTML, которая представляет основной текст формы ответа. Эта строка должна быть не более 32 КБ. The string is limited to 32 KB.</td></tr><tr><td><code>attachments</code></td><td>Array.&lt;Object&gt;</td><td>&lt;необязательно&gt;</td><td>Необязательный параметр. Используется для метода переопределения. Массив объектов JSON, представляющих собой вложенные файлы или элементы. При использовании этого параметра не применяйте параметр "body".<br/><br/><strong>Свойства</strong><br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Описание</th></tr></thead><tbody><tr><td><code>type</code></td><td>String</td><td>Indicates the type of attachment. Must be <code>file</code> for a file attachment or <code>item</code> for an item attachment.</td></tr><tr><td><code>name</code></td><td>String</td><td>Обязательный параметр. Строка с именем вложения (до 255 символов).</td></tr><tr><td><code>url</code></td><td>String</td><td>Only used if <code>type</code> is set to <code>file</code>. Обязательный параметр. URL-адрес, по которому расположены файлы.</td></tr><tr><td><code>itemId</code></td><td>String</td><td>Only used if <code>type</code> is set to <code>item</code>. Обязательный параметр. Идентификатор элемента веб-служб Exchange во вложении. Представляет собой строку длиной не более 100 символов. Обязательный параметр. Идентификатор элемента веб-служб Exchange во вложении. Представляет собой строку длиной не более 100 символов.</td></tr></tbody></table></td></tr><tr><td><code>callback</code></td><td>функция</td><td>&lt;необязательно&gt;</td><td>После применения метода функция, переданная в параметр , вызывается с помощью параметра , который представляет собой объект   . For more information, see <a href="tutorial-asynchronous.html">Using asynchronous methods</a>.</td></tr></tbody></table>|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Примеры

Приведенный ниже код передает строку в функцию displayReplyForm`displayReplyAllForm`.

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

```JavaScript
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

```JavaScript
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

Если любой строковый параметр превышает указанные для него ограничения, displayReplyForm`displayReplyForm` возвращает исключение.

Если в параметре formData.attachments`formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать все вложения и вложить их в форму ответа. Если не удается добавить какие-либо вложения, в форме пользовательского интерфейса отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`formData`| String &#124; Object|Строка, содержащая текст и HTML, которая представляет основной текст формы ответа. Эта строка должна быть не более 32 КБ. The string is limited to 32 KB.<br/>**OR**<br/>An object that contains body or attachment data and a callback function. Поведение формы при ее открытии определяется следующим образом:<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>htmlBody</code></td><td>String</td><td>&lt;необязательно&gt;</td><td>Строка, содержащая текст и HTML, которая представляет основной текст формы ответа. Эта строка должна быть не более 32 КБ. The string is limited to 32 KB.</td></tr><tr><td><code>attachments</code></td><td>Array.&lt;Object&gt;</td><td>&lt;необязательно&gt;</td><td>Необязательный параметр. Используется для метода переопределения. Массив объектов JSON, представляющих собой вложенные файлы или элементы. При использовании этого параметра не применяйте параметр "body".<br/><br/><strong>Свойства</strong><br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Описание</th></tr></thead><tbody><tr><td><code>type</code></td><td>String</td><td>Indicates the type of attachment. Must be <code>file</code> for a file attachment or <code>item</code> for an item attachment.</td></tr><tr><td><code>name</code></td><td>String</td><td>Only used if <code>type</code> is set to <code>file</code>. Обязательный параметр. Строка с именем вложения (до 255 символов).</td></tr><tr><td><code>url</code></td><td>String</td><td>Only used if <code>type</code> is set to <code>file</code>. Обязательный параметр. URL-адрес, по которому расположены файлы.</td></tr><tr><td><code>itemId</code></td><td>String</td><td>Only used if <code>type</code> is set to <code>item</code>. Обязательный параметр. Идентификатор элемента веб-служб Exchange во вложении. Представляет собой строку длиной не более 100 символов. Обязательный параметр. Идентификатор элемента веб-служб Exchange во вложении. Представляет собой строку длиной не более 100 символов.</td></tr></tbody></table></td></tr><tr><td><code>callback</code></td><td>функция</td><td>&lt;необязательно&gt;</td><td>После применения метода функция, переданная в параметр , вызывается с помощью параметра , который представляет собой объект   . For more information, see <a href="tutorial-asynchronous.html">Using asynchronous methods</a>.</td></tr></tbody></table>|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Примеры

Приведенный ниже код передает строку в функцию displayReplyForm`displayReplyForm`.

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

```JavaScript
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

```JavaScript
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
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
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
|`entityType`| [Office.MailboxEnums.EntityType](Office.MailboxEnums.md#entitytype-string)|Одно из значений перечисления EntityType.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Применимый режим Outlook| Чтение|

##### Возвращаемое значение:

If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null. If no entities of the specified type are present on the item, the method returns an empty array. Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.

Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.

| Value of `entityType` | Тип объектов в возвращаемом массиве | Необходимый уровень разрешений |
| --- | --- | --- |
| `Address` | String | **Restricted** |
| `Contact` | Contact | **ReadItem** |
| `EmailAddress` | String | **ReadItem** |
| `MeetingSuggestion` | MeetingSuggestion | **ReadItem** |
| `PhoneNumber` | PhoneNumber | **Restricted** |
| `TaskSuggestion` | TaskSuggestion | **ReadItem** |
| `URL` | String | **Restricted** |

Type: Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>

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

Метод getFilteredEntitiesByName возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила ItemHasKnownEntity в XML-файле манифеста, с использованием указанного значения элемента FilterName.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`name`| String|Имя элемента правила ItemHasRegularExpressionMatch`ItemHasKnownEntity`, определяющее соответствующий фильтр.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Возвращаемое значение:

If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. Если в манифесте нет элемента ItemHasKnownEntity`name` со значением FilterName`ItemHasKnownEntity`, соответствующим параметру name, метод возвращает null. Если параметр name не соответствует элементу ItemHasKnownEntity в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.

Type: Array.<(String|[Contact](simple-types.md#contact)|[MeetingSuggestion](simple-types.md#meetingsuggestion)|[PhoneNumber](simple-types.md#phonenumber)|[TaskSuggestion](simple-types.md#tasksuggestion))>

#### getRegExMatches() → {Object}

Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.

Метод `ItemHasRegularExpressionMatch`getRegExMatchesByName`getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила ItemHasRegularExpressionMatch`ItemHasKnownEntity` в XML-файле манифеста, с использованием указанного значения элемента RegExName. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.

Например, рассмотрим манифест надстройки, который содержит следующий элемент Rule`Rule`:

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

Объект, возвращаемый методом getRegExMatches`getRegExMatches`, будет содержать два свойства: fruits`fruits` и veggies`veggies`.

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

Если вы указываете правило ItemHasRegularExpressionMatch`ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения типа .* для получения всего текста элемента не всегда приносит ожидаемые результаты. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](Body.md#getAsync) method to retrieve the entire body.

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Возвращаемое значение:

Массив строк, соответствующих регулярным выражениям, определяемым в XML-файле манифеста. Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута RegExName`RegExName` подходящего правила ItemHasRegularExpressionMatch`ItemHasRegularExpressionMatch` или атрибута FilterName`FilterName` соответствующего правила ItemHasKnownEntity`ItemHasKnownEntity`.

<dl class="param-type">Тип:<dd>Объект</dd>

</dl>

##### Пример

В примере ниже показано, как получить доступ к массиву совпадений для элементов <rule> регулярного выражения fruits`fruits` и veggies`veggies`, которые указаны в манифесте.</rule>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### getRegExMatchesByName(name) → (nullable) {Array.<String>}

Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.

Метод getRegExMatchesByName`getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила ItemHasRegularExpressionMatch`ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента RegExName`RegExName`.

Если вы указываете правило ItemHasRegularExpressionMatch`ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения типа .* для получения всего текста элемента не всегда приносит ожидаемые результаты. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`name`| String|Имя элемента правила ItemHasRegularExpressionMatch`ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Возвращаемое значение:

Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.

<dl class="param-type">type<dd>array<String></dd>

</dl>

##### Пример

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  getSelectedDataAsync(coercionType, [options], callback) → {String}

Асинхронно возвращает данные, выбранные в теме или тексте сообщения.

Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку InvalidSelection. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`coercionType`| [Office.CoercionType](Office.md#coerciontype-string)||Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).|
|`options`| Объект| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Объект</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult).

To access the selected data from the callback method, call `asyncResult.value.data`. Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите asyncResult.value.data. Для доступа к исходному свойству, на основе которого созданы выбранные данные, вызовите параметр asyncResult.value.sourceProperty, который может иметь значение body или subject.

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Применимый режим Outlook| Создание|

##### Возвращаемое значение:

Выбранные данные в виде строки с форматом, определенным в параметре coercionType`coercionType`.

<dl class="param-type">Тип:<dd>String</dd>

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

Custom properties are stored as key/value pairs on a per-app, per-item basis. Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект CustomProperties`CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным. Custom properties are not encrypted on the item, so this should not be used as secure storage.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`callback`| функция||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult).

The custom properties are provided as a [`CustomProperties`](CustomProperties.md) object in the `asyncResult.value` property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.| |`userContext`| Object| &lt;optional&gt;|Developers can provide any object they wish to access in the callback function. This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. В следующем примере кода JavaScript показано, как асинхронно использовать метод loadCustomPropertiesAsync`myProp` для загрузки настраиваемых свойств текущего элемента, и как использовать метод saveAsync`otherProp``CustomProperties.get` для сохранения этих свойств на сервере. После загрузки настраиваемых свойств в примере кода метод get`saveAsync``CustomProperties.set` вызывается для чтения настраиваемого свойства myProp, метод set — для записи настраиваемого свойства otherProp, а метод saveAsync — для сохранения настраиваемых свойств.

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

The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. Рекомендуется использовать идентификатор вложения для удаления, только если то же почтовое приложение добавило вложение в том же сеансе. В olwebshortolowadevices идентификатор вложения действителен только в одном сеансе. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме и затем выходит за ее пределы, продолжая работу в отдельном окне.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`attachmentId`| String||Идентификатор удаляемого вложения. Максимальная длина строки — 100 символов. Идентификатор удаляемого вложения. Максимальная длина строки — 100 символов.|
|`options`| Объект| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Объект</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). <br/>If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.<br/><table class="nested-table"><thead><tr><th>Код ошибки</th><th>Описание</th></tr></thead><tbody><tr><td><code>InvalidAttachmentId</code></td><td>Индекс вложения не существует.</td></tr></tbody></table>|

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

####  setSelectedDataAsync(data, [options], callback)

Асинхронно вставляет данные в текст или тему сообщения.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`data`| String||Обязательный параметр. Вставляемый OOXML-код. Data is not to exceed 1,000,000 characters. Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение ArgumentOutOfRange`ArgumentOutOfRange`.|
|`options`| Объект| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Объект</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;необязательно&gt;</td><td>If <code>text</code>, the current style is applied in Outlook Web App and Outlook. Если задано значение text, текущий стиль применяется в olwebshort и Outlook. Если задано значение text и поле — это HTML-редактор, вставляются только текстовые данные, даже если данные представляют собой HTML.</td></tr></tbody></table><p>If <code>html</code> and the field supports HTML (the subject doesn&#39;t), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an <code>InvalidDataFormat</code> error is returned.</p><p>Если параметр coercionType не задан, результат зависит от поля. Если поле содержит HTML, используется HTML. Если поле текстовое, используется обычный текст.|</p>|
|`callback`| функция||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). |

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.2|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Применимый режим Outlook| Создание|

##### Пример

```JavaScript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
