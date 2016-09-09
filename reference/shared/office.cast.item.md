
# Свойство Office.cast.item
Предоставляет IntelliSense для сообщений и встреч в режимах создания и чтения.

|||
|:-----|:-----|
|**Ведущие приложения:**|Outlook|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Почтовый ящик|
|**Последнее изменение в **|1.0|



|||
|:-----|:-----|
|**Применимые режимы Outlook**|Только разработка в Visual Studio|

```js
Office.cast.item.toAppointmentCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toAppointmentRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toAppointment(Office.context.mailbox.item);
```

```js
Office.cast.item.toItemCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toItemRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessageCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessageRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessage(Office.context.mailbox.item);
```


## Возвращаемое значение

Набор методов, позволяющих выбрать подходящий IntelliSense для надстройки Outlook.


## Заметки

Это свойство и его методы поддерживают IntelliSense для разработки надстроек Outlook только в Visual Studio. Они не влияют на другие средства разработки.

Во время разработки в Visual Studio используют методы **Office.cast.item**, чтобы получить подходящие варианты IntelliSense для свойства **Office.context.mailbox.item**. Например, при использовании метода **toAppointmentCompose** технология IntelliSense отобразит только те методы и свойства **Appointment**, которые применяются в режиме создания.

Методы **Office.cast.item** не влияют на надстройку Outlook во время ее работы.


## Пример

В следующем примере используется метод **toMessageCompose** для приведения свойства **Office.context.mailbox.item**, чтобы отображался IntelliSense только для объекта **Message** в режиме создания. После приведения переменная `message` будет отображать IntelliSense только для тех методов и свойств, которые можно использовать в режиме создания.


```js
var message = Office.cast.item.toMessageCompose(Office.context.mailbox.item);

```


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).

||Office для рабочего стола Windows|Office Online (в браузере)|Outlook для Mac|
|:-----|:-----|:-----|:-----|
|**Outlook**|Y|Да|Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|Почтовый ящик|
|**Минимальный уровень разрешений**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Outlook|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



|**Версия**|**Изменения**|
|:-----|:-----|
|1.0|Представлено|
