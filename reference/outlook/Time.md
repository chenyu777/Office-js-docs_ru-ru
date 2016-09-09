

# Time

Объект `Time` возвращается как свойство встречи [`start`](Office.context.mailbox.item.md#start-datetime) или [`end`](Office.context.mailbox.item.md#end-datetime) в режиме создания.

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание|

### Методы

####  getAsync([options], callback)

Получает время начала или окончания встречи.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция||После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult).

Дата и время указываются в виде объекта Date в свойстве `asyncResult.value`. Значение приводится в формате UTC. Время в формате UTC можно преобразовать в локальное время клиента с помощью метода [`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание|
####  setAsync(dateTime, [options], [callback])

Задает время начала или окончания встречи.

Если для свойства [`setAsync`](Office.context.mailbox.item.md#start-datetime) вызывается метод `start`, свойство [`end`](Office.context.mailbox.item.md#end-datetime) будет настроено для поддержки предварительно заданной продолжительности встречи. Если для свойства `setAsync` вызывается метод `end`, продолжительность встречи будет расширена до нового времени окончания.

Время необходимо указать в формате UTC. Правильное время в формате UTC можно получить с помощью метода [`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date).

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`dateTime`| Дата||Объект Date в формате UTC.|
|`options`| Object| &lt;необязательно&gt;|Литерал объекта, содержащий один или несколько из указанных ниже свойств.<br/><br/>**Свойства**<br/><table class="nested-table"><thead><tr><th>Имя</th><th>Тип</th><th>Атрибуты</th><th>Описание</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Object</td><td>&lt;необязательно&gt;</td><td>Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</td></tr></tbody></table>|
|`callback`| функция| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). <br/>Если не удается задать дату и время, свойство `asyncResult.error` будет содержать код ошибки.<br/><table class="nested-table"><thead><tr><th>Код ошибки</th><th>Описание</th></tr></thead><tbody><tr><td><code>InvalidEndTime</code></td><td>Время окончания встречи предшествует времени начала встречи.</td></tr></tbody></table>|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.1|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Применимый режим Outlook| Создание|

##### Пример

В примере ниже устанавливается время начала встречи.

```js
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
