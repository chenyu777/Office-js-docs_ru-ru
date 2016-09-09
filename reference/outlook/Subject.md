

# Subject

Предоставляет методы для получения и задания темы для встречи или сообщения в надстройке Outlook.

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание|

### Методы

####  getAsync([options], callback)

Получает тему встречи или сообщения.

Метод `getAsync` выполняет асинхронный вызов на сервер Exchange Server, чтобы получить тему встречи или сообщения.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult).

Тема элемента указывается в виде строки в свойстве `asyncResult.value`.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание|
####  setAsync(subject, [options], [callback])

Задает тему встречи или сообщения.

Метод `setAsync` запускает асинхронный запрос на сервере Exchange Server, чтобы задать тему встречи или сообщения. При задании темы перезаписывается текущая тема, но сохраняются все префиксы, такие как "Fwd:" или "Re:".

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`subject`| String||Тема встречи или сообщения. Строка может содержать до 255 символов.|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). <br/>Если не удастся задать тему, свойство `asyncResult.error` будет содержать код ошибки.<br/><table class="nested-table"><thead><tr><th>Код ошибки</th><th>Описание</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td>Параметр <code>subject</code> содержит более 255 символов.</td></tr></tbody></table>|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание|
