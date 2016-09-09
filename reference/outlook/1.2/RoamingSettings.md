

# RoamingSettings

Параметры, созданные при использовании методов объекта `RoamingSettings`, сохраняются для каждой надстройки и каждого пользователя отдельно. То есть они доступны только для создавшей их надстройки и только из почтового ящика пользователя, в котором они сохранены.

> Хотя API "Надстройка Outlook" предоставляет доступ к этим параметрам только надстройке, с помощью которой они созданы, эти параметры не следует считать безопасным способом хранения. К ним можно получить доступ с помощью веб-служб Exchange или расширенного MAPI. Их не следует использовать для хранения конфиденциальных сведений, таких как учетные данные пользователя или маркеры безопасности.

Имя параметра — это String, а значение может быть String, Number, Boolean, null, Object или Array.

К объекту `RoamingSettings` можно получить доступ с помощью свойства [`roamingSettings`](Office.context.md#roamingsettings-roamingsettings) в пространстве имен `Office.context`.

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Применимый режим Outlook| Создание или чтение|

### Пример

```JavaScript
// Get the current value of the 'myKey' setting
var value = Office.context.roamingSettings.get('myKey');
// Update the value of the 'myKey' setting
Office.context.roamingSettings.set('myKey', 'Hello World!');
// Persist the change
Office.context.roamingSettings.saveAsync();
```

### Методы

####  get(name) → (nullable) {String|Number|Boolean|Object|Array}

Извлекает указанный параметр.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`name`| String|Имя извлекаемого параметра с учетом регистра.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Применимый режим Outlook| Создание или чтение|

##### Возвращаемое значение:

<dl class="param-type">

<dt>
Тип</dt>


<dd>String | Number | Boolean | Object | Array</dd>

</dl>

####  remove(name)

Удаляет указанный параметр.

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`name`| String|Имя удаляемого параметра с учетом регистра|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Применимый режим Outlook| Создание или чтение|
####  saveAsync([callback])

Сохраняет параметры.

Параметры, сохраненные надстройкой, загружаются при инициализации, поэтому во время сеанса вы можете использовать методы [`set`](RoamingSettings.md#set) и [`get`](RoamingSettings.md#get) для работы с копией контейнера свойств в памяти. Если вы хотите сохранить параметры для работы с ними в дальнейшем, используйте метод `saveAsync`.

##### Параметры:

|Имя| Тип| Атрибуты| Описание|
|---|---|---|---|
|`callback`| function| &lt;необязательно&gt;|После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](simple-types.md#asyncresult). |

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Применимый режим Outlook| Создание или чтение|
####  set(name, value)

Устанавливает или создает указанный параметр.

Метод set создает новый параметр с указанным именем, если он еще не существует, или устанавливает существующий параметр с указанным именем. Значение сохраняется в документе как сериализованное представление JSON с его типом данных.

Для каждой надстройки можно задать параметры размером не более 2 МБ. Максимальный размер каждого из параметров составляет 32 КБ.

Любые изменения, внесенные в параметры с помощью функции `set`, будут сохранены на сервере только после вызова функции [`saveAsync`](RoamingSettings.md#saveasynccallback).

##### Параметры:

|Имя| Тип| Описание|
|---|---|---|
|`name`| String|Имя устанавливаемого или создаваемого параметра с учетом регистра.|
|`value`| String &#124; Number &#124; Boolean &#124; Object &#124; Array|Значение для сохранения.|

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Применимый режим Outlook| Создание или чтение|