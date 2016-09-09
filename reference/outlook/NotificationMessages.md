

# NotificationMessages

## NotificationMessages

Объект `NotificationMessages` возвращается в качестве свойства [`notificationMessages`](Office.context.mailbox.item.md#notificationmessages-notificationmessages) элемента.

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.3|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

### Методы

####  addAsync(key, JSONmessage, [options], [callback])

Добавляет уведомление к элементу.

Для каждого сообщения можно задать не более 5 уведомлений. Если задать больше, будет возвращена ошибка `NumberOfNotificationMessagesExceeded`.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`key`| String||Указанный разработчиком ключ, используемый для ссылки на это сообщение уведомления. Разработчики могут использовать его для изменения этого сообщения в дальнейшем. Его длина не должна превышать 32 символа.|
|`JSONmessage`| Object||Объект JSON, содержащий сообщение уведомления, которое необходимо добавить к элементу. Он состоит из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Описание</th></tr></thead><tbody><tr><td><code>type</code></td><td><a href="Office.MailboxEnums.md#.ItemNotificationMessageType">Office.MailboxEnums.ItemNotificationMessageType</a></td><td>Указывает тип сообщения. Если свойство type имеет значение <code>ProgressIndicator</code> или <code>ErrorMessage</code>, автоматически отображается значок, а сообщение не сохраняется. Поэтому значок и сохраняемые свойства недопустимы для этих типов сообщений. Их включение приведет к ошибке <code>ArgumentException</code>. Если параметр type имеет значение <code>ProgressIndicator</code>, разработчику следует удалить или заменить индикатор хода выполнения после завершения действия.</td></tr><tr><td><code>icon</code></td><td>String</td><td>Ссылка на значок, определенный в манифесте в разделе <code>Resource</code>. Он появляется на информационной панели. Применяется, только если параметр type имеет значение <code>InformationalMessage</code>. Если указать для этого параметра неподдерживаемый тип, будет возвращено исключение.</td></tr><tr><td><code>message</code></td><td>String</td><td>Текст сообщения уведомления. Максимальная длина составляет 150 символов. Если разработчик передает строку большей длины, возвращается исключение <code>ArgumentOutOfRange</code>.</td></tr><tr><td><code>persistent</code></td><td>Boolean</td><td>Применяется, только если параметр type имеет значение <code>InformationalMessage</code>. Если задано значение <code>true</code>, сообщение сохраняется, пока его не удалит эта надстройка или не закроет пользователь. Если задано значение <code>false</code>, оно удаляется при переходе к другому элементу. Что касается уведомлений об ошибках, сообщение сохраняется, пока пользователь не увидит его. Если указать для этого параметра неподдерживаемый тип, будет возвращено исключение.</td></tr></tbody></table>|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). |

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.3|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```
// Create three notifications, each with a different key
Office.context.mailbox.item.notificationMessages.addAsync("progress", {
  type: "progressIndicator",
  message : "An add-in is processing this message."
});
Office.context.mailbox.item.notificationMessages.addAsync("information", {
  type: "informationalMessage",
  message : "The add-in processed this message.",
  icon : "iconid",
  persistent: false
});
Office.context.mailbox.item.notificationMessages.addAsync("error", {
  type: "errorMessage",
  message : "The add-in failed to process this message."
});
```

####  getAllAsync([options], [callback])

Возвращает все ключи и сообщения для элемента.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult).

После успешного завершения свойство `asyncResult.value` будет содержать массив объектов [`NotificationMessageDetails`](simple-types.md#notificationmessagedetails).|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.3|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```
// Get all notifications
Office.context.mailbox.item.notificationMessages.getAllAsync(function (asyncResult) {
  if (asyncResult.status != "failed") {
    Office.context.mailbox.item.notificationMessages.replaceAsync( "notifications", {
      type: "informationalMessage",
      message : "Found " + asyncResult.value.length + " notifications.",
      icon : "iconid",
      persistent: false
    });
  }
});
```

####  removeAsync(key, [options], [callback])

Удаляет сообщение уведомления для элемента.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`key`| Для указания||Ключ для удаления сообщения уведомления.|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult).

Если ключ не найден, возвращается ошибка `KeyNotFound` в свойстве `asyncResult.error`.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.3|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```
// Remove a notification
Office.context.mailbox.item.notificationMessages.removeAsync("progress");
```

####  replaceAsync(key, JSONmessage, [options], [callback])

Заменяет сообщение уведомления с заданным ключом на другое сообщение.

Если сообщение уведомления с указанным ключом не существует, `replaceAsync` добавит уведомление.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`key`| Для указания||Ключ для заменяемого сообщения уведомления. Максимальная длина — 32 символа.|
|`JSONmessage`| Object||Объект JSON, содержащий новое сообщение уведомления, которое заменяет существующее сообщение. Он состоит из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Описание</th></tr></thead><tbody><tr><td><code>type</code></td><td><a href="Office.MailboxEnums.md#.ItemNotificationMessageType">Office.MailboxEnums.ItemNotificationMessageType</a></td><td>Указывает тип сообщения. Если свойство type имеет значение <code>ProgressIndicator</code> или <code>ErrorMessage</code>, автоматически отображается значок, а сообщение не сохраняется. Поэтому значок и сохраняемые свойства недопустимы для этих типов сообщений. Их включение приведет к ошибке <code>ArgumentException</code>. Если параметр type имеет значение <code>ProgressIndicator</code>, разработчику следует удалить или заменить индикатор хода выполнения после завершения действия.</td></tr><tr><td><code>icon</code></td><td>String</td><td>Ссылка на значок, определенный в манифесте в разделе <code>Resource</code>. Он появляется на информационной панели. Применяется, только если параметр type имеет значение <code>InformationalMessage</code>. Если указать для этого параметра неподдерживаемый тип, будет возвращено исключение.</td></tr><tr><td><code>message</code></td><td>String</td><td>Текст сообщения уведомления. Максимальная длина составляет 150 символов. Если разработчик передает строку большей длины, возвращается исключение <code>ArgumentOutOfRange</code>.</td></tr><tr><td><code>persistent</code></td><td>Boolean</td><td>Применяется, только если параметр type имеет значение <code>InformationalMessage</code>. Если задано значение <code>true</code>, сообщение сохраняется, пока его не удалит эта надстройка или не закроет пользователь. Если задано значение <code>false</code>, оно удаляется при переходе к другому элементу. Что касается уведомлений об ошибках, сообщение сохраняется, пока пользователь не увидит его. Если указать для этого параметра неподдерживаемый тип, будет возвращено исключение.</td></tr></tbody></table>|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). |

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.3|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```
// Replace a notification with an informational notification
Office.context.mailbox.item.notificationMessages.replaceAsync("progress", {
  type: "informationalMessage",
  message : "The message was processed successfully.",
  icon : "iconid",
  persistent: false
});
```
