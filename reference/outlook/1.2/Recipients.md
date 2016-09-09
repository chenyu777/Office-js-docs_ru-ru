

# Recipients

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание|

### Методы

####  addAsync(recipients, [options], [callback])

Добавляет список получателей к существующим получателям для встречи или сообщения.

В качестве параметра `recipients` можно задать массив из любых следующих элементов:

*   строки, содержащие электронные адреса SMTP;
*   объекты `EmailUser`;
*   объекты `EmailAddressDetails`.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`recipients`| Array.&lt;(String&#124;[EmailUser](simple-types.md#emailuser)&#124;[EmailAddressDetails](simple-types.md#emailaddressdetails))&gt;||Получатели, которых нужно добавить в список получателей.|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). <br/>Если добавить получателей не удастся, свойство `asyncResult.error` будет содержать код ошибки.<br/><table class="nested-table"><thead><tr><th>Код ошибки</th><th>Описание</th></tr></thead><tbody><tr><td>`NumberOfRecipientsExceeded</td><td>Превышено максимальное количество получателей (100).</td></tr></tbody></table>|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Применимый режим Outlook| Создание|

##### Пример

В следующем примере создается массив объектов `EmailUser`, которые добавляются к получателям сообщения, указанным в строке "Кому".

```
var newRecipients = [
  {
    "displayName": "Allie Bellew",
    "emailAddress": "allieb@contoso.com"
  },
  {
    "displayName": "Alex Darrow",
    "emailAddress": "alexd@contoso.com"
  }
];

Office.context.mailbox.item.to.addAsync(newRecipients, function(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Recipients added");
  }
});
```

####  getAsync([options], callback)

Возвращает список получателей для встречи или сообщения.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult).

После завершения вызова свойство `asyncResult.value` будет содержать массив объектов [`EmailAddressDetails`](simple-types.md#emailaddressdetails).|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание|

##### Пример

В следующем примере возвращаются необязательные участники собрания.

```js
Office.context.mailbox.item.optionalAttendees.getAsync(function(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    var msg = "";
    result.value.forEach(function(recip, index) {
      msg = msg + recip.displayName + " (" + recip.emailAddress + ");";
    });
    showMessage(msg);
  }
});
```

####  setAsync(recipients, [options], callback)

Задает список получателей для встречи или сообщения.

Метод `setAsync` перезаписывает текущий список получателей.

В качестве параметра `recipients` можно задать массив из любых следующих элементов:

*   строки, содержащие электронные адреса SMTP;
*   объекты `EmailUser`;
*   объекты `EmailAddressDetails`.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`recipients`| Array.&lt;(String&#124;[EmailUser](simple-types.md#emailuser)&#124;[EmailAddressDetails](simple-types.md#emailaddressdetails))&gt;||Получатели, которых нужно добавить в список получателей.|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). <br/>Если задать получателей не удастся, свойство `asyncResult.error` будет содержать код ошибки, произошедшей при добавлении данных.<br/><table class="nested-table"><thead><tr><th>Код ошибки</th><th>Описание</th></tr></thead><tbody><tr><td>`NumberOfRecipientsExceeded</td><td>Превышено максимальное количество получателей (100).</td></tr></tbody></table>|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Применимый режим Outlook| Создание|

##### Пример

В следующем примере создается массив объектов `EmailUser`, который заменяет получателей сообщения, указанных в строке "Копия".

```
var newRecipients = [
  {
    "displayName": "Allie Bellew",
    "emailAddress": "allieb@contoso.com"
  },
  {
    "displayName": "Alex Darrow",
    "emailAddress": "alexd@contoso.com"
  }
];

Office.context.mailbox.item.cc.setAsync(newRecipients, function(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Recipients overwritten");
  }
});
```
