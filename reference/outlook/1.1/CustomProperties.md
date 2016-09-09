

# CustomProperties

Объект `CustomProperties` представляет настраиваемые свойства, характерные для конкретного элемента и почтовой надстройки Outlook. Например, может возникнуть необходимость в почтовой надстройке, сохраняющей некоторые данные текущего электронного сообщения, которое активировало надстройку. Если впоследствии пользователь снова откроет это сообщение и активирует почтовую надстройку, она сможет извлечь данные, сохраненные в виде настраиваемых свойств.

Так как Outlook для Mac не кэширует настраиваемые свойства, в случае сбоя в сети пользователя почтовые надстройки не смогут получить к ним доступ.

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

### Пример

В приведенном ниже примере показано, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. В этом примере также показано, как сохранять эти свойства на сервере с помощью метода [`saveAsync`](CustomProperties.md#saveasynccallback-asynccontext). После загрузки настраиваемых свойств в этом примере используется метод [`get`](CustomProperties.md#get) для считывания настраиваемого свойства `myProp`, метод [`set`](CustomProperties.md#set) для записи настраиваемого свойства `otherProp` и метод `saveAsync` для сохранения настраиваемых свойств.

```
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

##### Параметры

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

<dl class="param-type">

<dt>Тип</dt>

<dd>String</dd>

</dl>

####  remove(name)

Удаляет указанное свойство из коллекции настраиваемых свойств.

Чтобы свойство было удалено безвозвратно, вызовите метод [`saveAsync`](CustomProperties.md#saveasynccallback-asynccontext) объекта `CustomProperties`.

##### Параметры

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

Необходимо вызвать метод `saveAsync`, чтобы сохранить все изменения, внесенные с помощью метода [`set`](CustomProperties.md#set) или [`remove`](CustomProperties.md#remove) объекта `CustomProperties`. Сохранение — асинхронное действие.

Рекомендуем сделать так, чтобы функция обратного вызова проверяла и обрабатывала ошибки из `saveAsync`. В частности, надстройка чтения может активироваться, когда подключенный пользователь открыл форму чтения, а затем отключился. Если надстройка вызывает `saveAsync` в отключенном состоянии, `saveAsync` возвращает ошибку. Метод обратного вызова должен обрабатывать эту ошибку соответствующим образом.

##### Параметры

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`callback`| function| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). |
|`asyncContext`| Object| &lt;необязательно&gt;|Данные о состоянии, передаваемые в метод обратного вызова.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

Приведенный ниже пример кода JavaScript иллюстрирует асинхронное использование метода `loadCustomPropertiesAsync` для загрузки настраиваемых свойств, характерных для текущего элемента, и метода [`saveAsync`](CustomProperties.md#saveasynccallback-asynccontext) для сохранения этих свойств на сервере. После загрузки настраиваемых свойств в этом примере кода используется метод [`get`](CustomProperties.md#get) для считывания настраиваемого свойства `myProp`, метод [`set`](CustomProperties.md#set) для записи настраиваемого свойства `otherProp` и метод `saveAsync` для сохранения настраиваемых свойств.

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

Метод `set` присваивает указанному свойству заданное значение. Метод [`saveAsync`](CustomProperties.md#saveasynccallback-asynccontext) необходимо использовать для сохранения свойства на сервере.

Метод `set` создает свойство, если указанное свойство не существует. В противном случае текущее значение заменяется новым. Параметр `value` может быть любого типа, но всегда передается на сервер в виде строки.

##### Параметры

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