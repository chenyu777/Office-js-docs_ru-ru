

# CustomProperties

The `CustomProperties` object represents custom properties that are specific to a particular item and specific to a mail add-in for Outlook. For example, there might be a need for a mail add-in to save some data that is specific to the current email message that activated the add-in. If the user revisits the same message in the future and activates the mail add-in again, the add-in will be able to retrieve the data that had been saved as custom properties.

Так как Outlook для Mac не кэширует настраиваемые свойства, в случае сбоя в сети пользователя надстройки почты не смогут получить к ним доступ.

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

### Пример

The following example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the [`saveAsync`](#saveasync) method to save these properties back to the server. В следующем примере показано, как использовать метод loadCustomPropertiesAsync для асинхронной загрузки настраиваемых свойств текущего элемента. В примере также показано, как использовать метод  saveAsync  для сохранения этих свойств на сервере. После загрузки настраиваемых свойств в примере метод  get  вызывается для чтения настраиваемого свойства myProp, метод  set  — для записи настраиваемого свойства otherProp, а метод saveAsync — для сохранения настраиваемых свойств.

```JavaScript
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var mailbox = Office.context.mailbox;
    mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
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

### Методы

####  get(name) → {String}

Возвращает значение указанного настраиваемого свойства.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`name`| String|Имя возвращаемого настраиваемого свойства.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Возвращаемое значение:

Значение указанного настраиваемого свойства.

<dl class="param-type">Тип:<dd>String</dd>

</dl>

####  remove(name)

Удаляет указанное свойство из коллекции настраиваемых свойств.

Чтобы удаление свойства было постоянным, вызовите метод  saveAsync  объекта CustomProperties.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`name`| String|Имя удаляемого свойства.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|
####  saveAsync([callback], [asyncContext])

Сохраняет настраиваемые свойства конкретного элемента на сервере.

Необходимо вызвать метод saveAsync, чтобы сохранить все изменения, внесенные с помощью метода  set  или  remove  объекта CustomProperties. Сохранение — асинхронное действие. The saving action is asynchronous.

It’s a good practice to have your callback function check for and handle errors from `saveAsync`. In particular, a read add-in can be activated while the user is in a connected state in a read form, and subsequently the user becomes disconnected. If the add-in calls `saveAsync` while in the disconnected state, `saveAsync` would return an error. Your callback method should handle this error accordingly.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`callback`| функция| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). |
|`asyncContext`| Объект| &lt;необязательно&gt;|Данные о состоянии, передаваемые в метод обратного вызова.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

В следующем примере кода JavaScript показано, как асинхронно использовать метод loadCustomPropertiesAsync для загрузки настраиваемых свойств текущего элемента, и как использовать метод  saveAsync  для сохранения этих свойств на сервере. После загрузки настраиваемых свойств в примере кода метод  get  вызывается для чтения настраиваемого свойства myProp, метод  set  — для записи настраиваемого свойства otherProp, а метод saveAsync — для сохранения настраиваемых свойств. В следующем примере кода JavaScript показано, как асинхронно использовать метод loadCustomPropertiesAsync для загрузки настраиваемых свойств текущего элемента, и как использовать метод  saveAsync  для сохранения этих свойств на сервере. После загрузки настраиваемых свойств в примере кода метод  get  вызывается для чтения настраиваемого свойства myProp, метод  set  — для записи настраиваемого свойства otherProp, а метод saveAsync — для сохранения настраиваемых свойств.

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
  if (asyncResult.status == Office.AsyncResultStatus.Failed){
    write(asyncResult.error.message);
  }
  else {
    // Async call to save custom properties completed.
    // Proceed to do the appropriate for your add-in.
  }
}

// Writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  set(name, value)

Присваивает указанному свойству заданное значение.

Метод set`set` присваивает указанному свойству заданное значение. Чтобы сохранить свойство на сервере, необходимо использовать метод saveAsync. Метод set присваивает указанному свойству заданное значение. Чтобы сохранить свойство на сервере, необходимо использовать метод  saveAsync .

Метод set`set` создает свойство, если указанное свойство не существует. В противном случае существующее значение заменяется на новое значение. Параметр value может быть любого типа, однако он всегда передается на сервер в виде строки. The `value` parameter can be of any type; however, it is always passed to the server as a string.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`name`| String|Имя свойства, которому присваивается значение.|
|`value`| Object|Значение, присваиваемое свойству.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|