

# mailbox

## [Office](Office.md)[.context](Office.context.md). mailbox

Предоставляет для Microsoft Outlook и Microsoft Outlook в Интернете доступ к объектной модели надстройки Outlook.

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Применимый режим Outlook| Создание или чтение|

### Пространства имен

[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.

[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.

[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</dd>

### Элементы

#### ewsUrl :String

Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.

Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://msdn.microsoft.com/EN-US/library/office/dn148008.aspx).

Для вызова метода `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.

Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item#saveAsync). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.

##### Тип:

*   String

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

### Методы

####  convertToEwsId(itemId, restVersion) → {String}

Преобразовывает идентификатор элемента из формата REST в формат EWS.

Формат идентификаторов, извлекаемых через API REST (такие как [API Почты Outlook](https://msdn.microsoft.com/office/office365/APi/mail-rest-operations) или [Microsoft Graph](http://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`itemId`| String|Идентификатор элемента в формате REST API для Outlook|
|`restVersion`| [Office.MailboxEnums.RestVersion](Office.MailboxEnums.md#restversion)|Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.3|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Применимый режим Outlook| Создание или чтение|

##### Возвращаемое значение:

Тип: String

##### Пример

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  convertToLocalClientTime(timeValue) → {[LocalClientTime](simple-types.md#localclienttime)}

Получает словарь, содержащий сведения о локальном времени клиента.

В случае дат и времени в почтовом приложении для Outlook или Outlook Web App могут использоваться разные часовые пояса. Outlook использует часовой пояс клиентского компьютера. Outlook Web App использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.

Если почтовое приложение работает в Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook Web App, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`timeValue`| Date|Объект Date|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Возвращаемое значение:

Тип: [LocalClientTime](simple-types.md#localclienttime)

####  convertToRestId(itemId, restVersion) → {String}

Преобразовывает идентификатор элемента в формате EWS в формат REST.

Формат идентификаторов, извлекаемых через EWS или свойство `itemId` отличается от формата API REST (таких как [API Почты Outlook](https://msdn.microsoft.com/office/office365/APi/mail-rest-operations) или [Microsoft Graph](http://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`itemId`| String|Идентификатор элемента в формате EWS|
|`restVersion`| [Office.MailboxEnums.RestVersion](Office.MailboxEnums.md#restversion)|Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.3|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Применимый режим Outlook| Создание или чтение|

##### Возвращаемое значение:

Тип: String

##### Пример

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  convertToUtcClientTime(input) → {Date}

Получает объект Date из словаря, содержащего сведения о времени.

Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`input`| [LocalClientTime](simple-types.md#localclienttime)|Значение локального времени для преобразования.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Возвращаемое значение:

Объект Date со временем в формате UTC.

<dl class="param-type">

<dt>Тип</dt>

<dd>Date</dd>

</dl>

####  displayAppointmentForm(itemId)

Отображает имеющуюся встречу из календаря.

Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.

В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).

В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.

Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`itemId`| String|Идентификатор веб-служб Exchange для существующей встречи в календаре.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  displayMessageForm(itemId)

Отображает имеющееся сообщение.

Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.

В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.

Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.

Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`itemId`| String|Идентификатор веб-служб Exchange для существующего сообщения.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### displayNewAppointmentForm(parameters)

Отображает форму для создания новой встречи в календаре.

Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.

В Outlook Web App и Outlook Web App для устройств этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.

Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в полнофункциональном клиенте Outlook или Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.

Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`parameters`| Object|Словарь параметров, описывающий новую встречу.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Описание</th></tr></thead><tbody><tr><td><code>requiredAttendees</code></td><td>Array.&lt;String&gt; &#124; Array.&lt;<a href="simple-types.md#emailaddressdetails">EmailAddressDetails</a>&gt;</td><td>Массив строк, содержащий электронные адреса, или массив, содержащий объекты <code>EmailAddressDetails</code> для каждого из требуемых участников встречи. Массив может включать не более 100 записей.</td></tr><tr><td><code>optionalAttendees</code></td><td>Array.&lt;String&gt; &#124; Array.&lt;<a href="simple-types.md#emailaddressdetails">EmailAddressDetails</a>&gt;</td><td>Массив строк, содержащий адреса электронной почты, или массив, содержащий объект EmailAddressDetails для каждого необязательного участника встречи. Размер массива ограничен 100 записями.</td></tr><tr><td><code>start</code></td><td>Date</td><td>Объект Date, указывающий дату и время начала встречи.</td></tr><tr><td><code>end</code></td><td>Date</td><td>Объект Date, указывающий дату и время окончания встречи.</td></tr><tr><td><code>location</code></td><td>String</td><td>Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</td></tr><tr><td><code>resources</code></td><td>Array.&lt;String&gt;</td><td>Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</td></tr><tr><td><code>subject</code></td><td>String</td><td>Строка с темой встречи. Максимальное количество символов в строке — 255.</td></tr><tr><td><code>body</code></td><td>String</td><td>Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</td></tr></tbody></table>|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Чтение|

##### Пример

```
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

#### getCallbackTokenAsync(callback, [userContext])

Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.

Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный токен с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.

Вы можете передать сторонней системе токен и идентификатор вложения или элемента. Сторонняя система использует этот токен как токен авторизации, чтобы вызвать операцию [GetAttachment](https://msdn.microsoft.com/en-us/library/office/aa494316.aspx) или [GetItem](https://msdn.microsoft.com/en-us/library/office/aa565934.aspx) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://msdn.microsoft.com/EN-US/library/office/dn148008.aspx).

Для вызова метода `getCallbackTokenAsync` в режиме чтения манифесте приложения должно быть указано разрешение **ReadItem**.

Чтобы получить идентификатор элемента для передачи в метод `getCallbackTokenAsync`, в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item#saveAsync). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`callback`| function||После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). Токен указывается в виде строки в свойстве `asyncResult.value`.|
|`userContext`| Object| &lt;необязательно&gt;|Данные о состоянии, передаваемые в асинхронный метод.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.3|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание и чтение|

##### Пример

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  getUserIdentityTokenAsync(callback, [userContext])

Получает маркер, идентифицирующий пользователя и надстройку Office.

Метод `getUserIdentityTokenAsync` возвращает токен, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](https://msdn.microsoft.com/EN-US/library/office/fp179828.aspx).

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`callback`| function||После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult).

Токен предоставляется как строка в свойстве `asyncResult.value`.| |`userContext`| Object| &lt;дополнительно&gt;|Любые данные о состоянии, которые передаются в асинхронный метод.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  makeEwsRequestAsync(data, callback, [userContext])

Выполняет асинхронный запрос к службе Exchange Web Services (EWS) на сервере Exchange, на котором размещен почтовый ящик пользователя.

Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.

С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.

В запросе XML должна быть указана кодировка UTF-8.

```
<?xml version="1.0" encoding="utf-8"?>
```

У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операциях EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье о [разрешениях для надстроек Outlook](../../docs/outlook/understanding-outlook-add-in-permissions.md).

**ПРИМЕЧАНИЕ**. Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге Client Access Server EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.

#### Различия версий

Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`data`| String||Запрос EWS.|
|`callback`| function||После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult).

Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`. Если размер результата более 1 МБ, возвращается сообщение об ошибке.| |`userContext`| Объект | &lt;необязательно&gt;|Любые данные о состоянии, которые передаются в асинхронный метод.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteMailbox|
|Применимый режим Outlook| Создание или чтение|

##### Пример

В следующем примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item.
   var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

   return request;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;

   // Process the returned response here.
}
```
