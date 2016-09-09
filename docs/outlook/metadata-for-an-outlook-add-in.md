
# Получение и задние метаданных для надстройки Outlook

Для управления пользовательскими данными в настройке Outlook можно использовать следующее:

- параметры перемещения, которые управляют пользовательскими данными для почтового ящика пользователя;
    
- настраиваемые свойства, которые управляют пользовательскими данными для элемента в почтовом ящике пользователя.
    
Оба этих способа предоставляют доступ к пользовательским данным, доступным только надстройке Outlook, но каждый метод хранит данные отдельно от остальных. Другими словами, данные, хранящиеся с помощью параметров перемещения, недоступны настраиваемым свойствам и наоборот. Данные хранятся на сервере этого почтового ящика и доступны в последующих сеансах Outlook на всех поддерживаемых надстройкой форм-факторах. 

## Пользовательские данные на один почтовый ящик: параметры перемещения


Вы можете указать данные, специфичные для пользователя почтового ящика Exchange, с помощью объекта [RoamingSettings](../../reference/outlook/RoamingSettings.md). Примерами таких данных являются личные данные и предпочтения пользователя. Ваша почтовая надстройка может получить доступ к параметрам перемещения, когда перемещение происходит на любом из устройств, предназначенных для работы (настольный ПК, планшет или смартфон).

 Изменения этих данных хранятся в памяти текущего сеанса Outlook. После изменения все параметры перемещения следует сохранить, чтобы они были доступны, когда пользователь откроет надстройку на том же или другом поддерживаемом устройстве в следующий раз.


### Формат параметров перемещения


Данные в объекте  **RoamingSettings** хранятся в качестве строки последовательной нотации объектов JavaScript (JSON). Ниже приведен пример структуры при условии, что используется три определенных параметра перемещения с именами `add-in_setting_name_0`,  `add-in_setting_name_1` и `add-in_setting_name_2`.


```js
{
  "add-in_setting_name_0":"add-in_setting_value_0",
  "add-in_setting_name_1":"add-in_setting_value_1",
  "add-in_setting_name_2":"add-in_setting_value_2"
}
```


### Загрузка параметров перемещения


Почтовая надстройка обычно загружает параметры перемещения в обработчик событий [Office.initialize](../../reference/shared/office.initialize.md). В следующем примере кода JavaScript показано, как выполняется загрузка существующих параметров перемещения и получение значений 2 параметров "customerName" и "customerBalance":


```js
var _mailbox;
var _settings;
var _customerName;
var _customerBalance;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Initialize instance variables to access API objects.
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;
  _customerName = _settings.get("customerName");
  _customerBalance = _settings.get("customerBalance");
}

```


### Создание или назначение параметра перемещения


Развивая предыдущий пример, следующая функция JavaScript —  `setAddInSetting` — показывает, как использовать метод [RoamingSettings.set](../../reference/outlook/RoamingSettings.md) для определения заданного параметра `cookie` с указанием сегодняшнего числа и сохранения данных с помощью метода [RoamingSettings.saveAsync](../../reference/outlook/RoamingSettings.md), который позволяет сохранить все параметры перемещения обратно на сервер. Метод  **set** создает параметр, если таковой еще не существует, и назначает для него определенное значение. Метод **saveAsync** сохраняет параметры перемещения асинхронно. Этот пример кода передает метод вызова `saveMyAddInSettingsCallback` в **saveAsync**. После завершения асинхронного вызова  `saveMyAddInSettingsCallback` вызывается с указанием одного параметра _asyncResult_. Этот параметр является объектом [AsyncResult](../../reference/outlook/simple-types.md), который содержит результат асинхронного вызова, а также любые соответствующие сведения. Вы можете использовать дополнительный параметр  _userContext_ для передачи любых сведений о состоянии из асинхронного вызова в функцию обратного вызова.


```js
// Set a roaming setting.
function setAddInSetting() {
  _settings.set("cookie", Date());
  // Save roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}

// Callback method after saving custom roaming settings.
function saveMyAddInSettingsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```


### Удаление параметра перемещения


Кроме того, в расширениях предыдущих примеров следующая функция JavaScript —  `removeAddInSetting` — показывает, как метод [RoamingSettings.remove](../../reference/outlook/RoamingSettings.md) используется для удаления параметра `cookie` и сохранения всех параметров перемещения обратно в Exchange Server.


```js
// Remove an add-in setting.
function removeAddInSetting()
{
  _settings.remove("cookie");
  // Save changes to the roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}
```


## Пользовательские данные для каждого элемента в почтовом ящике: пользовательские свойства


Вы также можете указать данные, характерные для элемента в почтовом ящике пользователя, используя объект [CustomProperties](../../reference/outlook/CustomProperties.md). Например, ваша почтовая надстройка могла бы категоризировать некоторые сообщения и отмечать категорию с помощью настраиваемого свойства  `messageCategory`. Либо, если ваша почтовая надстройка создает встречи из сообщений с предложениями о собрании, вы можете использовать настраиваемое свойство, чтобы отслеживать каждую из этих встреч. Это гарантирует, что если пользователь вновь откроет сообщение, ваша почтовая надстройка не станет во второй раз предлагать создать встречу.

Аналогично параметрам перемещения, изменения настраиваемых свойств хранятся в копии контейнера свойств для текущего сеанса Outlook. Чтобы эти настраиваемые свойства были доступны при следующем сеансе, сохраните все настраиваемые свойства на сервере.

Эти настраиваемые свойства, характерные для надстроек и объектов, доступны только при использовании объекта  **CustomProperties**. Эти свойства отличаются от настраиваемых свойств, основанных на интерфейсе MAPI, [UserProperties](http://msdn.microsoft.com/library/20b49c86-d74f-9bda-382c-559af278c148%28Office.15%29.aspx) в модели объектов Outlook, и расширенных свойств в Веб-службы Exchange (EWS). Получить доступ к **CustomProperties** с помощью модели объектов Outlook или веб-службы Exchange невозможно.

Тем не менее, почтовая надстройка может получить расширенные свойства, основанные на интерфейсе MAPI, с помощью операции веб-службы Exchange [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx). Получите доступ к  **GetItem** на стороне сервера за счет использования токена обратного вызова или на стороне клиента за счет использования метода [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md). В запросе  **GetItem** укажите настраиваемые расширенные свойства, необходимые в наборе свойств. Кроме того, почтовая надстройка может использовать **makeEwsRequestAsync** и операции веб-службы Exchange [CreateItem](http://msdn.microsoft.com/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx) и [UpdateItem](http://msdn.microsoft.com/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx) для создания и изменения расширенных свойств.




### Использование настраиваемых свойств


Перед использованием настраиваемых свойств необходимо загрузить их, вызвав метод [loadCustomPropertiesAsync](../../reference/outlook/Office.context.mailbox.item.md). Если какие-либо настраиваемые свойства уже заданы для текущего элемента, они загружаются из сервера Exchanger в этот момент. После создания контейнера свойств можно использовать методы [set](../../reference/outlook/CustomProperties.md) и [get](../../reference/outlook/CustomProperties.md) для добавления и извлечения настраиваемых свойств. Чтобы сохранить изменения, внесенные в контейнер свойств, необходимо использовать метод [saveAsync](../../reference/outlook/CustomProperties.md) для сохранения изменений на сервере Exchange.


 >**Примечание**  Так как Outlook для Mac не кэширует настраиваемые свойства, то в случае, если сеть пользователя пропадет, почтовые надстройки в Outlook для Mac не смогут получить доступ к их настраиваемым свойствам.


### Пример пользовательских свойств


Следующий пример показывает простой набор методов для надстройки Outlook, использующей настраиваемые свойства. Этот пример можно использовать в качестве отправной точки для создания надстройки, использующей настраиваемые свойства. 

Этот пример содержит следующие методы:


- [Office.initialize](../../reference/shared/office.initialize.md): инициализирует надстройку и загружает контейнер настраиваемых свойств с сервера Exchange Server.
    
-  **customPropsCallback**: получает контейнер настраиваемых свойств, возвращенный с сервера, и сохраняет его для дальнейшего использования.
    
-  **updateProperty**: задает или обновляет определенное свойство, а затем сохраняет изменения на сервере.
    
-  **removeProperty**: удаляет определенное свойство из контейнера свойств, а затем сохраняет удаление на сервере.
    



```js
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
}

// Callback function from loading custom properties.
function customPropsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
  else {
    // Successfully loaded custom properties,
    // can get them from the asyncResult argument.
    _customProps = asyncResult.value;
  }
}

// Get individual custom property.
function getProperty() {
  var myProp = customProps.get("myProp");
}

// Set individual custom property.
function updateProperty(name, value) {
  _customProps.set(name, value);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Remove a custom property.
function removeProperty(name) {
  _customProps.remove(name);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Callback function from saving custom properties.
function saveCallback() {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```


## Дополнительные ресурсы

    
- [Обзор свойств MAPI](http://msdn.microsoft.com/library/02e5b23f-1bdb-4fbf-a27d-e3301a359573%28Office.15%29.aspx)
    
- [Обзор свойств Outlook](http://msdn.microsoft.com/library/242c9e89-a0c5-ff89-0d2a-410bd42a3461%28Office.15%29.aspx)
    
- [Вызов веб-служб из надстройки Outlook](../outlook/web-services.md)
    
- [Свойства и расширенные свойства в веб-службах Exchange](http://msdn.microsoft.com/library/68623048-060e-4602-b3fa-62617a94cf72%28Office.15%29.aspx)
    
- [Наборы свойств и ответ с фигурами в веб-служб Exchange в Exchange](http://msdn.microsoft.com/library/04a29804-6067-48e7-9f5c-534e253a230e%28Office.15%29.aspx)
    


