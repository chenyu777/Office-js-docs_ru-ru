

# Body

Объект `body` обеспечивает методы добавления и обновления содержимого сообщения или встречи. Он возвращается в свойстве `body` выбранного элемента.

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

### Методы

####  getAsync(coercionType, [options], [callback])

Возвращает текущий текст в указанном формате.

Этот метод возвращает весь текущий текст в формате, заданном в параметре `coercionType`.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`coercionType`| [Office.CoercionType](Office.md#coerciontype-string)||Формат возвращаемого текста.|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult).

Текст предоставляется в запрошенном формате в свойстве `asyncResult.value`.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.3|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Примеры

В этом примере возвращается текст сообщения в формате обычного текста.

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext:"This is passed to the callback" },
  function callback(result) {
    // Do something with the result
  });
```

Ниже приведен пример параметра `result`, переданного функции обратного вызова.

```js
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  getTypeAsync([options], [callback])

Получает значение, указывающее формат содержимого: HTML или текст.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult).

Тип содержимого возвращается в виде одного из значений [CoercionType](Office.md#coerciontype-string) в свойстве `asyncResult.value`.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание|
####  prependAsync(data, [options], [callback])

Добавляет указанное содержимое в начало текста элемента.

Метод `prependAsync` вставляет указанную строку в начало текста элемента. Вызов метода `prependAsync` аналогичен вызову метода [`setSelectedDataAsync`](#setselecteddataasync) с точкой вставки в начале содержимого текста.

При включении ссылок в разметку HTML вы можете отключить предварительный просмотр ссылок в сети, задав для атрибута `id` с привязкой `<a>` значение `LPNoLP`. Например:

```js
Office.context.mailbox.item.body.prependAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`data`| String||Строка, добавляемая в начало основного текста. Максимальная длина — 1 000 000 символов.|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;необязательно&gt;</td><td>Необходимый формат текста. Строка в параметре <code>data</code> будет преобразована в этот формат.</td></tr><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). <br/>Все обнаруженные ошибки будут указаны в свойстве `asyncResult.error`.<br/><table class="nested-table"><thead><tr><th>Код ошибки</th><th>Описание</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td>Параметр <code>data</code> включает более 1 000 000 символов.</td></tr></tbody></table>|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Применимый режим Outlook| Создание|
####  setAsync(data, [options], [callback])

Заменяет весь текст указанным текстом.

Метод `setAsync` заменяет существующий текст элемента с указанной строкой или, если текст выделен в редакторе, он заменяет выделенный текст.

При включении ссылок в разметку HTML вы можете отключить предварительный просмотр ссылок в сети, задав для атрибута `id` с привязкой `<a>` значение `LPNoLP`. Например:

```js
Office.context.mailbox.item.body.setAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`data`| String||Строка, которая заменяет существующий текст. Максимальная длина — 1 000 000 символов.|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;необязательно&gt;</td><td>Необходимый формат текста. Строка в параметре <code>data</code> будет преобразована в этот формат.</td></tr><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). <br/>Все обнаруженные ошибки будут указаны в свойстве `asyncResult.error`.<br/><table class="nested-table"><thead><tr><th>Код ошибки</th><th>Описание</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td>Параметр <code>data</code> включает более 1 000 000 символов.</td></tr><tr><td><code>InvalidFormatError</code></td><td>Для параметра <code>options.coercionType</code> задано значение <code>Office.CoercionType.Html</code>, а текст сообщения указан в формате обычного текста.</td></tr></tbody></table>|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.3|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Применимый режим Outlook| Создание|

##### Примеры

В примере ниже текст заменяется на HTML-содержимое.

```js
Office.context.mailbox.item.body.setAsync(
  "<b>(replaces all body, including threads you are replying to that may be on the bottom)</b>",
  { coercionType:"html", asyncContext:"This is passed to the callback" },
  function callback(result) {
    // Process the result
  });
```

Ниже приведен пример параметра `result`, переданного функции обратного вызова.

```js
{
  "value":null,
  "status":"succeeded",
  "asyncContext":"This is passed to the callback"
}
```

####  setSelectedDataAsync(data, [options], [callback])

Заменяет выделенный фрагмент в основном тексте на заданный текст.

Метод `setSelectedDataAsync` вставляет указанную строку на месте указателя в тексте элемента. Если текст выбран в редакторе, он заменяет выделенный текст. Если указатель не появлялся в тексте элемента, или текст элемент потерял фокус в пользовательском интерфейсе, строка будет вставлена в начало содержимого текста.

При включении ссылок в разметку HTML вы можете отключить предварительный просмотр ссылок в сети, задав для атрибута `id` с привязкой `<a>` значение `LPNoLP`. Например:

```js
Office.context.mailbox.item.body.setSelectedDataAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`data`| String||Строка, добавляемая к основному тексту. Максимальная длина — 1 000 000 символов.|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;необязательно&gt;</td><td>Необходимый формат текста. Строка в параметре <code>data</code> будет преобразована в этот формат.</td></tr><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). <br/>Все обнаруженные ошибки будут указаны в свойстве `asyncResult.error`.<br/><table class="nested-table"><thead><tr><th>Код ошибки</th><th>Описание</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td>Параметр <code>data</code> включает более 1 000 000 символов.</td></tr><tr><td><code>InvalidFormatError</code></td><td>Задан тип текста HTML, а параметр data содержит обычный текст.</td></tr></tbody></table>|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Применимый режим Outlook| Создание|
