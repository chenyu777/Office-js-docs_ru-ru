

# userProfile

## [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

### Элементы

####  displayName :String

Получает отображаемое имя пользователя.

##### Тип:

*   String

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  emailAddress :String

Получает адрес электронной почты SMTP пользователя.

##### Тип:

*   String

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  timeZone :String

Получает часовой пояс пользователя по умолчанию.

##### Тип:

*   String

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```