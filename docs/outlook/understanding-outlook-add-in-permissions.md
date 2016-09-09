
# Указание разрешений для доступа надстройки Outlook к почтовому ящику пользователя

Для надстроек Outlook указан необходимый уровень разрешений в их манифесте. Доступны следующие уровни: **Restricted**, **ReadItem**, **ReadWriteItem** или **ReadWriteMailbox**. Эти уровни разрешений можно определить как накопительные: **Restricted** — это самый низкий уровень, а каждый последующий уровень включает разрешения всех более низких уровней. Уровень **ReadWriteMailbox** включает все поддерживаемые разрешения.

Вы можете просмотреть разрешения, которые запрашивает почтовая надстройка, перед ее установкой из Магазин Office. Вы также можете просмотреть требуемые разрешения установленных надстроек в Центре администрирования Exchange.


## Разрешение ограниченного доступа


Разрешение **Restricted** — это базовый уровень разрешений. Укажите элемент **Restricted** в элементе [Permissions](http://msdn.microsoft.com/en-us/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx) манифеста, чтобы запросить это разрешение. Outlook по умолчанию назначает это разрешение почтовой надстройке, если в ее манифесте не запрашивается определенное разрешение.


### Разрешено


- [Получать только определенные сущности](../outlook/match-strings-in-an-item-as-well-known-entities.md) (номер телефона, адрес, URL-адрес) из темы или тела элемента.
    
- Указывать [правило активации ItemIs](../outlook/manifests/activation-rules.md#itemis-rule), требующее, чтобы текущий элемент в форме чтения или создания принадлежал определенному типу, или правило [ItemHasKnownEntity](../outlook/match-strings-in-an-item-as-well-known-entities.md), соответствующее малому поднабору поддерживаемых известных сущностей (номер телефона, адрес, URL-адрес) в выбранном элементе.
    
- Получать доступ к свойствам и методам, которые **не** относятся к определенной информации о пользователе или элементе. (В следующем разделе приведен список членов, которые относятся к такой информации.)
    

### Не разрешено


- Использовать правило [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) для контакта, адреса электронной почты, предложения о собрании или объекта предложения о задаче.
    
- Использовать правило [ItemHasAttachment](http://msdn.microsoft.com/en-us/library/031db7be-8a25-5185-a9c3-93987e10c6c2%28Office.15%29.aspx) или [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx).
    
- Получать доступ к участникам следующего списка, связанных с информацией о пользователе или элементом. При попытке получить доступ к членам, приведенным в этом списке, будет возвращено значение **null** и будет выведено сообщение об ошибке, говорящее, что Outlook требует у почтовой надстройки более высокого разрешения.
    
      - [item.addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.attachments](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.bcc](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.body](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.cc](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.from](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.getRegExMatches](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.getRegExMatchesByName](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.organizer](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.resources](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.sender](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.to](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.getUserIdentityTokenAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.userProfile](../../reference/outlook/Office.context.mailbox.userProfile.md)
    
  - [Body](../../reference/outlook/Body.md) и все дочерние элементы
    
  - [Location](../../reference/outlook/Location.md) и все дочерние элементы
    
  - [Recipients](../../reference/outlook/Recipients.md) и все дочерние элементы
    
  - [Subject](../../reference/outlook/Subject.md) и все дочерние элементы
    
  - [Time](../../reference/outlook/Time.md) и все дочерние элементы
    

## Разрешение ReadItem


Разрешение **ReadItem** — это следующий уровень в модели разрешений. Укажите элемент **ReadItem** в элементе **Permissions** манифеста, чтобы запросить это разрешение.


### Разрешено


- [Считывать все свойства](../outlook/item-data.md) текущего элемента в чтении или [Создавать форму](../outlook/get-and-set-item-data-in-a-compose-form.md), например [item.to](../../reference/outlook/Office.context.mailbox.item.md) в форме чтения и [item.to.getAsync](../../reference/outlook/Recipients.md) в форме создания.
    
- [Получать токен обратного звонка для получения вложений элемента](../outlook/get-attachments-of-an-outlook-item.md) или всего элемента.
    
- [Записывать пользовательские свойства](http://msdn.microsoft.com/library/30217d63-7615-4f3f-8618-c91e4e60cd43%28Office.15%29.aspx), установленные надстройкой для соответствующего элемента.
    
- [Получать все существующие известные сущности](../outlook/match-strings-in-an-item-as-well-known-entities.md) (а не только поднабор) из темы или текста элемента.
    
- Использовать все [известные сущности](../outlook/manifests/activation-rules.md#itemhasknownentity-rule) в правилах [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) или [регулярные выражения](../outlook/manifests/activation-rules.md#itemhasregularexpressionmatch-rule) в правилах [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx). Следующий пример, использующий схему версии 1.1, активирует надстройку, если обнаруживается одна или несколько известных сущностей в теме или теле выбранного сообщения:
    

```XML
<Permissions>ReadItem</Permissions>
    <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" FormType = "Read" ItemType="Message" />
    <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="PhoneNumber" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="MeetingSuggestion" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="TaskSuggestion" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="EmailAddress" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Contact" />
</Rule>
```


### Не разрешено

Получать доступ к **mailbox.makeEWSRequestAsync** или следующим методам записи:


- [item.addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.bcc.addAsync](../../reference/outlook/Recipients.md)
    
- [item.bcc.setAsync](../../reference/outlook/Recipients.md)
    
- [item.body.prependAsync](../../reference/outlook/Body.md)
    
- [item.body.setAsync](../../reference/outlook/Body.md)
    
- [item.body.setSelectedDataAsync](../../reference/outlook/Body.md)
    
- [item.cc.addAsync](../../reference/outlook/Recipients.md)
    
- [item.cc.setAsync](../../reference/outlook/Recipients.md)
    
- [item.end.setAsync](../../reference/outlook/Time.md)
    
- [item.location.setAsync](../../reference/outlook/Location.md)
    
- [item.optionalAttendees.addAsync](../../reference/outlook/Recipients.md)
    
- [item.optionalAttendees.setAsync](../../reference/outlook/Recipients.md)
    
- [item.removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.requiredAttendees.addAsync](../../reference/outlook/Recipients.md)
    
- [item.requiredAttendees.setAsync](../../reference/outlook/Recipients.md)
    
- [item.start.setAsync](../../reference/outlook/Time.md)
    
- [item.subject.setAsync](../../reference/outlook/Subject.md)
    
- [item.to.addAsync](../../reference/outlook/Recipients.md)
    
- [item.to.setAsync](../../reference/outlook/Recipients.md)
    

## Разрешение ReadWriteItem


Укажите элемент  **ReadWriteItem** в элементе **Permissions** манифеста, чтобы запросить это разрешение. Почтовые надстройки, активированные в формах создания, которые используют методы записи (**Message.to.addAsync** или **Message.to.setAsync**), должны использовать по крайней мере этот уровень разрешений.


### Разрешено


- [Считывать и записывать все свойства на уровне элемента](../outlook/item-data.md) для элемента, который просматривается или создается в Outlook.
    
- [Добавлять или удалять вложения](../outlook/add-and-remove-attachments-to-an-item-in-a-compose-form.md) для такого элемента.
    
- Использовать все другие объекты JavaScript API для Office, которые применимы к почтовым надстройкам, за исключением **Mailbox.makeEWSRequestAsync**.
    

### Не разрешено

Использовать **Mailbox.makeEWSRequestAsync**.


## Разрешение ReadWriteMailbox


Разрешение **ReadWriteMailbox** — это высший уровень разрешений. Укажите элемент **ReadWriteMailbox** в элементе **Permissions** манифеста, чтобы запросить это разрешение.

В дополнение к тому, что поддерживает уровень разрешений  **ReadWriteItem**, с помощью **Mailbox.makeEWSRequestAsync** вы можете получать доступ к поддерживаемым операциям веб-служб Exchange (EWS) для выполнения следующих действий:


- Чтение и запись всех свойств любого элемента в почтовом ящике пользователя.
    
- Создание, чтение и запись в любую папку или элемент в таком почтовом ящике.
    
- Отправка элемента из такого почтового ящика
    
С помощью  **mailbox.makeEWSRequestAsync** вы можете использовать следующие операции EWS:


- [CopyItem](http://msdn.microsoft.com/en-us/library/bcc68f9e-d511-4c29-bba6-ed535524624a%28Office.15%29.aspx)
    
- [CreateFolder](http://msdn.microsoft.com/en-us/library/6f6c334c-b190-4e55-8f0a-38f2a018d1b3%28Office.15%29.aspx)
    
- [CreateItem](http://msdn.microsoft.com/en-us/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx)
    
- [FindConversation](http://msdn.microsoft.com/en-us/library/2384908a-c203-45b6-98aa-efd6a4c23aac%28Office.15%29.aspx)
    
- [FindFolder](http://msdn.microsoft.com/en-us/library/7a9855aa-06cc-45ba-ad2a-645c15b7d031%28Office.15%29.aspx)
    
- [FindItem](http://msdn.microsoft.com/en-us/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx)
    
- [GetConversationItems](http://msdn.microsoft.com/en-us/library/8ae00a99-b37b-4194-829c-fe300db6ab99%28Office.15%29.aspx)
    
- [GetFolder](http://msdn.microsoft.com/en-us/library/355bcf93-dc71-4493-b177-622afac5fdb9%28Office.15%29.aspx)
    
- [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx)
    
- [MarkAsJunk](http://msdn.microsoft.com/en-us/library/1f71f04d-56a9-4fee-a4e7-d1034438329e%28Office.15%29.aspx)
    
- [MoveItem](http://msdn.microsoft.com/en-us/library/dcf40fa7-7796-4a5c-bf5b-7a509a18d208%28Office.15%29.aspx)
    
- [SendItem](http://msdn.microsoft.com/en-us/library/337b89ef-e1b7-45ed-92f3-8abe4200e4c7%28Office.15%29.aspx)
    
- [UpdateFolder](http://msdn.microsoft.com/en-us/library/3494c996-b834-4813-b1ca-d99642d8b4e7%28Office.15%29.aspx)
    
- [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx)
    
Попытка использования неподдерживаемой операции приведет к возврату ошибки.


## Дополнительные ресурсы



- [Конфиденциальность, разрешения и безопасность для надстроек Outlook](../outlook/../../docs/develop/privacy-and-security.md)
    
- [Сопоставление строк в элементе Outlook как известных сущностей](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
