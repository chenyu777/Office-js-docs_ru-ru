
# Считывание и запись данных элемента в форме создания элементов Outlook
Сведения о том, как получать и задавать различные свойства элемента в надстройке Outlook в сценарии создания, такие как сведения о получателях, тема, текст, а также место и время встречи.




## Получение и установка свойств элемента для надстройки создания


В формах создания можно получить доступ к большинству свойств, предоставляемых таким типом элемента в форме чтения (например, участники, получатели, тема и текст), а несколько дополнительных свойств доступны только в форме создания (текст, СК). 

Методы получения и задания большинства этих свойств асинхронные, так как надстройка Outlook и пользователь могут изменять одно свойство в пользовательском интерфейсе одновременно. В таблице 1 перечислены свойства уровня элемента и соответствующие асинхронные методы, позволяющие их получить и задать в форме создания. Исключение составляют свойства [item.itemType](../../reference/outlook/Office.context.mailbox.item.md) и [item.conversationId](../../reference/outlook/Office.context.mailbox.item.md), потому что пользователи не могут их менять. Их можно получить программно в форме создания так же, как и в форме чтения, напрямую из родительского объекта.

Помимо доступа к свойствам элемента в JavaScript API для Office, вы можете получить доступ к свойствам уровня элемента, используя веб-службы Exchange (EWS). С разрешением  **ReadWriteMailbox** вы можете использовать метод [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) для доступа к операциям EWS [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) и [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx), чтобы получать и устанавливать дополнительные свойства элементов в почтовом ящике пользователя. Операция  **makeEwsRequestAsync** доступна в формах создания и чтения. Дополнительные сведения о разрешении **ReadWriteMailbox** и доступе к EWS на платформе Надстройки Office см. в статьях [Указание разрешений для доступа надстройки Outlook к почтовому ящику пользователя](../outlook/understanding-outlook-add-in-permissions.md) и [Вызов веб-служб из надстройки Outlook](../outlook/web-services.md).


**Таблица 1. Асинхронные методы для получения и установки свойства элемента в форме создания**


|**Свойство**|**Тип свойства**|**Асинхронный метод для получения свойства**|**Асинхронные методы для установки свойства**|
|:-----|:-----|:-----|:-----|
|[bcc](../../reference/outlook/Office.context.mailbox.item.md)|[Получатели](../../reference/outlook/Recipients.md)|[Recipients.getAsync](../../reference/outlook/Recipients.md)|[Recipients.addAsync](../../reference/outlook/Recipients.md)[Recipients.setAsync](../../reference/outlook/Recipients.md)|
|[Основной текст](../../reference/outlook/Office.context.mailbox.item.md)|[Body](../../reference/outlook/Body.md)|[Body.getAsync](../../reference/outlook/Body.md)|[Body.prependAsync](../../reference/outlook/Body.md)[Body.setAsync](../../reference/outlook/Body.md)[Body.setSelectedDataAsync](../../reference/outlook/Body.md)|
|[cc](../../reference/outlook/Office.context.mailbox.item.md)|Получатели|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[end](../../reference/outlook/Office.context.mailbox.item.md)|[Time](../../reference/outlook/Time.md)|[Time.getAsync](../../reference/outlook/Time.md)|[Time.setAsync](../../reference/outlook/Time.md)|
|[location](../../reference/outlook/Office.context.mailbox.item.md)|[Расположение](../../reference/outlook/Location.md)|[Location.getAsync](../../reference/outlook/Location.md)|[Location.setAsync](../../reference/outlook/Location.md)|
|[optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)|Получатели|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)|Получатели|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[начать](../../reference/outlook/Office.context.mailbox.item.md)|Time|Time.getAsync|Time.setAsync|
|[subject](../../reference/outlook/Office.context.mailbox.item.md)|[Тема](../../reference/outlook/Subject.md)|[Subject.getAsync](../../reference/outlook/Subject.md)|[Subject.setAsync](../../reference/outlook/Subject.md)|
|[—](../../reference/outlook/Office.context.mailbox.item.md)|Получатели|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|



## Дополнительные ресурсы



- [Создание надстроек Outlook для форм создания](../outlook/compose-scenario.md)
    
- [Общие сведения о разрешениях для надстройки Outlook](../outlook/understanding-outlook-add-in-permissions.md)
    
- [Вызов веб-служб из надстройки Outlook](../outlook/web-services.md)
    
- [Считывание и запись данных элемента Outlook в формах просмотра и создания](../outlook/item-data.md)
    


