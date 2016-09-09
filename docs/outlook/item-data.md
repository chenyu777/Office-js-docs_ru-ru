
# Считывание и запись данных элемента Outlook в формах просмотра и создания

Начиная с версии 1.1 схемы манифестов для надстроек Office, Outlook может активировать надстройки как при просмотре, так и при создании элементов. В зависимости от того, используется ли при активации форма создания или форма просмотра, в надстройке доступны разные свойства. Например, свойства [dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md) и [dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md) определяются для отправленных элементов (которые затем отображаются в форме просмотра), но не для создаваемых элементов (которые редактируются в форме создания). Еще один пример — свойство [bcc](../../reference/outlook/Office.context.mailbox.item.md), которое имеет смысл только при составлении сообщений (в форме создания) и недоступно пользователям в режиме просмотра.

В таблице 1 показаны свойства уровня элемента в API JavaScript для Office, которые доступны в почтовых надстройках в режимах просмотра и создания элементов. Как правило, свойства в формах просмотра доступны только для чтения, а свойства в формах создания доступны как для чтения, так и для записи, за исключением свойств [itemId](../../reference/outlook/Office.context.mailbox.item.md)* и [conversationId](../../reference/outlook/Office.context.mailbox.item.md)*, которые всегда доступны только для чтения. Что касается остальных свойств уровня элемента, доступных в формах создания, так как надстройка и пользователь могут одновременно считывать или записывать одно и то же свойство, методы их получения и задания асинхронные, поэтому эти свойства возвращают в формах просмотра и создания элементов объекты разных типов. Дополнительные сведения об использовании асинхронных методов, позволяющих получить или задать свойства уровня элемента в режиме создания, см. в статье [Считывание и запись данных элемента в форме создания элементов Outlook](../outlook/get-and-set-item-data-in-a-compose-form.md).


**Таблица 1. Свойства элементов, доступные в формах создания и просмотра элементов**


|**Тип элемента**|**Свойство**|**Тип свойства в формах просмотра элементов**|**Тип свойства в формах создания элементов**|
|:-----|:-----|:-----|:-----|
|Встречи и сообщения|[dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md)|Объект JavaScript  **Date**|Свойство недоступно|
|Встречи и сообщения|[dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)|Объект JavaScript  **Date**|Свойство недоступно|
|Встречи и сообщения|[itemClass](../../reference/outlook/Office.context.mailbox.item.md)|String|Свойство недоступно|
|Встречи и сообщения|[itemId](../../reference/outlook/Office.context.mailbox.item.md)|String|Свойство недоступно|
|Встречи и сообщения|[itemType](../../reference/outlook/Office.context.mailbox.item.md)|Строка в перечислении [ItemType](../../reference/outlook/Office.MailboxEnums.md)|Свойство недоступно|
|Встречи и сообщения|[вложения](../../reference/outlook/Office.context.mailbox.item.md)|[AttachmentDetails](../../reference/outlook/simple-types.md)|Свойство недоступно|
|Встречи и сообщения|[Основной текст](../../reference/outlook/Office.context.mailbox.item.md)|[Body](../../reference/outlook/Body.md)|[Body](../../reference/outlook/Body.md)|
|Встречи|[end](../../reference/outlook/Office.context.mailbox.item.md)|Объект JavaScript  **Date**|[Time](../../reference/outlook/Time.md)|
|Встречи|[location](../../reference/outlook/Office.context.mailbox.item.md)|String|[Расположение](../../reference/outlook/Location.md)|
|Встречи и сообщения|[normalizedSubject](../../reference/outlook/Office.context.mailbox.item.md)|String|Свойство недоступно|
|Встречи|[optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)|[EmailAddressDetails](../../reference/outlook/simple-types.md)|[Recipients](../../reference/outlook/Recipients.md)|
|Встречи|[organizer](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Свойство недоступно|
|Встречи|[requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Recipients|
|Встречи|[ресурсы](../../reference/outlook/Office.context.mailbox.item.md)|String|Свойство недоступно|
|Встречи|[начать](../../reference/outlook/Office.context.mailbox.item.md)|Объект JavaScript  **Date**|Time|
|Встречи и сообщения|[subject](../../reference/outlook/Office.context.mailbox.item.md)|String|[Тема](../../reference/outlook/Subject.md)|
|Сообщения|[bcc](../../reference/outlook/Office.context.mailbox.item.md)|Свойство недоступно|Recipients|
|Сообщения|[cc](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Recipients|
|Сообщения|[conversationId](../../reference/outlook/Office.context.mailbox.item.md)|String|String (только для чтения)|
|Сообщения|[от](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Свойство недоступно|
|Сообщения|[internetMessageId](../../reference/outlook/Office.context.mailbox.item.md)|Целое число|Свойство недоступно|
|Сообщения|[sender](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Свойство недоступно|
|Сообщения|[—](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Recipients|

## Использование маркеров обратного вызова Exchange Server из надстройки для просмотра элементов


Если ваша надстройка Outlook должна активироваться в формах просмотра, вы можете получить маркер обратного вызова для Exchange. Этот маркер можно использовать в серверном коде для доступа ко всему элементу с помощью веб-служб Exchange (EWS). Указав разрешение  **ReadItem** в манифесте надстройки, вы сможете с помощью метода [mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md) получить маркер обратного вызова для Exchange, с помощью свойства [mailbox.ewsUrl](../../reference/outlook/Office.context.mailbox.md) — URL-адрес конечной точки EWS для почтового ящика пользователя, а с помощью свойства [item.itemId](../../reference/outlook/Office.context.mailbox.item.md) — идентификатор EWS для выбранного элемента. Затем вы можете передать маркер обратного вызова, URL-адрес конечной точки EWS и идентификатор элемента EWS серверному коду для доступа к операции [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx), позволяющей получить дополнительные свойства элемента.


## Доступ к веб-службам EWS из надстройки для просмотра или создания элементов


Кроме того, вы можете использовать метод [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) для доступа к операциям [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) и [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx), которые выполняются в веб-службах Exchange (EWS), непосредственно из надстройки. С помощью этих операций можно получить и задать множество свойств указанного элемента. Этот метод доступен надстройкам Outlook независимо от того, в какой форме они активированы (форме просмотра или создания), если в манифесте надстройки указано разрешение  **ReadWriteMailbox**. Дополнительные сведения об использовании  **makeEwsRequestAsync** для доступа к операциям EWS см. в статье [Вызов веб-служб из надстройки Outlook](../outlook/web-services.md).


## Дополнительные ресурсы



- [Надстройки Outlook](../outlook/outlook-add-ins.md)
    
- [Считывание и запись данных элемента в форме создания элементов Outlook](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [Вызов веб-служб из надстройки Outlook](../outlook/web-services.md)
    


