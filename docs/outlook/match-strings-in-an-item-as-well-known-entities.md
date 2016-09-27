﻿

# Сопоставление строк в элементе Outlook как известных сущностей


Прежде чем отправить сообщение или приглашение на собрание, Exchange Server анализирует содержимое элемента, определяет в теме и тексте строки, похожие на известные Exchange сущности, например адреса электронной почты, номера телефонов и URL-адреса, и добавляет соответствующие метки. Exchange Server отправляет сообщения и приглашения на собрания с помеченными известными сущностями в папку "Входящие" приложения Outlook. 

С помощью API JavaScript для Office можно получить строки, которые соответствуют известным сущностям, для дальнейшей обработки. Известную сущность также можно указать в правиле в манифесте надстройки, чтобы программа Outlook смогла активировать надстройку, когда пользователь просматривает элемент, который содержит совпадения с этой сущностью. Затем вы можете извлечь и обработать совпадения с этой сущностью. 

Возможность найти или извлечь такие экземпляры из выбранного сообщения или встречи очень полезна. Например, вы можете реализовать службу определения владельца номера как надстройку Outlook. Она может извлекать строки из темы или текста элемента, похожие на номер телефона, и показывать зарегистрированного владельца номера телефона.

В этой статье приводятся общие сведения об этих известных сущностях, приводятся примеры основанных на них правил активации, а также описывается, как извлекать совпадения с сущностями независимо от того, используются ли они в правилах активации.


## Поддержка известных сущностей


После отправки элемента и до его доставки получателю Exchange Server добавляет метки к известным сущностям в сообщении или приглашении на собрание. Поэтому метки добавляются только к тем элементам, которые транспортировались через Exchange, и приложение Outlook может активировать надстройки на основе этих меток, когда пользователь просматривает такие элементы. С другой стороны, когда пользователь создает или просматривает в папке "Отправленные" элемент, который не прошел процесс транспортировки, Outlook не может активировать надстройки на основе известных сущностей. 

По этой же причине невозможно извлечь известные сущности из элементов, которые создаются или находятся в папке "Отправленные", так как эти элементы не прошли процесс транспортировки и не содержат метки. Дополнительные сведения о типах элементов, которые поддерживают активацию, см. в статье [Правила активации надстроек Outlook](../outlook/manifests/activation-rules.md#activation-rules-for-outlook-add-ins).

В следующей таблице перечислены сущности, которые поддерживаются и распознаются приложениями Exchange Server и Outlook (отсюда и название "известные сущности"), а также типы экземпляров каждой сущности. Распознавание строки в качестве одной из этих сущностей основано на модели обучения восприятию естественного языка, выработанной путем анализа большого объема данных. Поэтому распознавание является недетерминированным. Дополнительные сведения об условиях распознавания см. в статье [Советы по использованию известных сущностей](#Советы-по-использованию-известных-сущностей).

 **Таблица 1. Поддерживаемые сущности и их типы**



|**Тип сущности**|**Условия для распознавания**|**Тип объекта**|
|:-----|:-----|:-----|
|**Address**|Адреса в США, например:1234 Main Street, Redmond, WA 07722.В целом, чтобы адрес был распознан, он должен соответствовать структуре почтовых адресов в США с присутствием большинства элементов, таких как номер дома, название улицы, города, штата и почтовый индекс. Адрес может быть указан на одной или нескольких строках.|Объект JavaScript  **String**|
|**Контакт**|Ссылка на сведения о человеке распознаются на естественном языке.Распознавание контакта зависит от контекста. Например, подпись в конце сообщения или имя человека, которое присутствует в непосредственной близости от таких сведений, как номер телефона, адрес, адрес эл. почты и URL-адрес.|[Объект ](../../reference/outlook/simple-types.md)Contact|
|**EmailAddress**|SMTP-адреса электронной почты.|Объект JavaScript  **String**|
|**MeetingSuggestion**|Ссылка на событие или собрание. Например, Exchange 2013 распознает следующий текст как приглашение на собрание: _Let's meet tomorrow for lunch (Встретимся завтра за обедом)._|[Объект ](../../reference/outlook/simple-types.md)MeetingSuggestion|
|**PhoneNumber**|Телефонные номера в США, например: _(235) 555-0110._|[Объект](../../reference/outlook/simple-types.md)PhoneNumber|
|**TaskSuggestion**|Предложения в сообщении эл. почты, являющиеся призывами к действию. Например: _Please update the spreadsheet (Просьба обновить таблицу)._|[Объект ](../../reference/outlook/simple-types.md)TaskSuggestion|
|**Url**|Веб-адрес, который в явной форме указывает сетевое расположение и идентификатор веб-ресурса. Exchange Server не требует указывать в веб-адресе протокол доступа и не распознает URL-адреса, вставленные в текст ссылки как экземпляры сущности  **Url**. Примеры сущностей, которые Exchange Server может распознавать: _www.youtube.com/user/officevideos_ _http://www.youtube.com/user/officevideos_|Объект JavaScript  **String**|
На рис. 1 показано, как Exchange Server и Outlook поддерживают известные сущности для надстроек, а также действия, которые надстройки могут совершать с ними. Дополнительные сведения см. в разделах [Получение сущностей в надстройке](#Получение-сущностей-в-надстройке) и [Активация надстройки в зависимости от наличия сущности](#Активация-надстройки-в-зависимости-от-наличия-сущности).


**Рис. 1. Поддержка известных сущностей в Exchange Server, Outlook и надстройках**

![Поддержка и использование известных сущностей в почтовом приложении](../../images/mod_off15_mailapp_wellknownentities_curvedlines.png)


## Разрешения на извлечение сущностей


Чтобы извлечь сущности в коде JavaScript или активировать надстройку в зависимости от наличия определенных известных сущностей, запросите нужные разрешения в манифесте приложения.

Ограниченное разрешение по умолчанию позволяет надстройке извлекать сущности  **Address**,  **MeetingSuggestion** и **TaskSuggestion**. Чтобы извлечь какую-либо другую сущность, необходимо указать разрешение на чтение элемента или чтение и запись элемента либо почтового ящика. Для этого в манифесте следует использовать элемент [Permissions](http://msdn.microsoft.com/en-us/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx) и указать соответствующее разрешение ( **Restricted**,  **ReadItem**,  **ReadWriteItem** или **ReadWriteMailbox**), как в приведенном ниже примере.




```XML
<Permissions>ReadItem</Permissions>
```


## Получение сущностей в надстройке


Если тема или текст элемента, которые просматривает пользователь, содержат строки, которые Exchange и Outlook могут распознать как известные сущности, то эти экземпляры будут доступны надстройке, даже если она была активирована без использования известных сущностей. Обладая соответствующим разрешением, можно использовать метод  **getEntities** или **getEntitiesByType**, чтобы извлечь известные сущности в текущем сообщении или встрече. Метод  **getEntities** возвращает массив объектов [Entities](../../reference/outlook/simple-types.md), который содержит все известные сущности в элементе. Если вас интересует определенный тип сущностей, используйте метод  **getEntitiesByType**, возвращающий массив нужных сущностей. Перечисление [EntityType](../../reference/outlook/Office.MailboxEnums.md) представляет все типы известных сущностей, которые можно извлечь.

После вызова  **getEntities** можно использовать соответствующее свойство объекта **Entities**, чтобы получить массив экземпляров типа сущности. В зависимости от типа сущности экземпляры массива могут быть просто строками или сопоставлениями с определенными объектами. Например, как показано на рис. 1, для получения адресов в элементе нужно обратиться к массиву, возвращаемому  `getEntities().addresses[]`. Свойство  **Entities.addresses** возвращает массив строк, которые распознаются в Outlook как почтовые адреса, а свойство **Entities.contacts** возвращает массив объектов **Contact**, которые распознаются в Outlook как контактная информация. В таблице 1 приведен список типов объектов экземпляра каждой поддерживаемой сущности.

В следующем примере показано, как получить любые адреса, найденные в сообщении.




```
// Get the address entities from the item.
var entities = Office.context.mailbox.item.getEntities();
// Check to make sure that address entities are present.
if (null != entities &amp;&amp; null != entities.addresses &amp;&amp; undefined != entities.addresses) {
   //Addresses are present, so use them here.
}

```


## Активация надстройки в зависимости от наличия сущности


Другой способ использовать известные сущности — активировать надстройку Outlook в зависимости от наличия одного или нескольких типов сущностей в теме или тексте элемента, который вы просматриваете в данный момент. Для этого в манифесте надстройки можно указать правило  **ItemHasKnownEntity**. Простой тип [KnownEntityType](http://msdn.microsoft.com/en-us/library/432d413b-9fcc-eb50-cfea-0ed10a43bd52%28Office.15%29.aspx) представляет разные типы известных сущностей, поддерживаемых правилами **ItemHasKnownEntity**. После активации надстройки вы также сможете получить экземпляры таких сущностей для ваших целей, как описано в предыдущем разделе [Получение сущностей в надстройке](#Получение-сущностей-в-надстройке). 

Вы также можете применить регулярное выражение в правиле  **ItemHasKnownEntity**, чтобы дополнительно отфильтровать экземпляры сущности и задействовать надстройку в Outlook только для поднабора экземпляров сущности. Например, можно указать сущность адреса в сообщении, содержащем почтовый индекс штата Вашингтон, который начинается с "98". Чтобы применить фильтр к экземплярам сущности, воспользуйтесь атрибутами  **RegExFilter** и **FilterName** в элементе [Rule](http://msdn.microsoft.com/en-us/library/56dfc32e-2b8c-1724-05be-5595baf38aa3%28Office.15%29.aspx) типа [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx).

Как и для других правил активации, можно указать несколько правил, составляющих коллекцию для вашей надстройки. Следующий пример применяет оператор "И" для двух правил:  **ItemIs** и **ItemHasKnownEntity**. Эта коллекция правил активирует надстройку, если текущий элемент является сообщением и Outlook распознает адрес в теме или тексте элемента.




```XML
<Rule xsi:type="RuleCollection" Mode="And">
   <Rule xsi:type="ItemIs" ItemType="Message" />
   <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

В следующем примере используется правило  **getEntitiesByType** текущего элемента для определения переменной `addresses` применительно к результатам предыдущей коллекции правил.




```
var addresses = Office.context.mailbox.item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
```

В следующем примере правила  **ItemHasKnownEntity** активация надстройки выполняется при наличии URL-адреса в теме или основном тексте текущего элемента и строки "youtube" в этом адресе независимо от регистра.




```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="Url" 
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

В следующем примере используется правило  **getFilteredEntitiesByName(name)** текущего элемента для определения переменной `videos` для получения массива результатов, которые соответствуют регулярному выражению в предыдущем правиле **ItemHasKnownEntity**.




```
var videos = Office.context.mailbox.item.getFilteredEntitiesByName(youtube);
```


## Советы по использованию известных сущностей


Необходимо помнить о некоторых аспектах и ограничениях при использовании известных сущностей в надстройке. Следующее справедливо, если почтовая надстройка активируется, когда пользователь читает элемент, содержащий соответствия известных сущностей (вне зависимости от того, используется ли правило  **ItemHasKnownEntity**).


1. Вы можете извлечь строки, представляющие известные сущности, только если они на английском языке.
    
2. Вы можете извлечь известные сущности только из первых 2000 символов в тексте элемента. Это помогает сбалансировать функциональность и производительность, чтобы Exchange Server и Outlook не тратили слишком много ресурсов на анализ и поиск экземпляров известных сущностей в крупных сообщениях и встречах. Обратите внимание на то, что этот предел не зависит от того, указывает ли надстройка правило  **ItemHasKnownEntity**. Если она использует такое правило, обратите внимание на предел обработки правил в элементе 2 ниже для полнофункционального клиента Outlook.
    
3. Вы можете извлекать сущности из встреч — собраний, организованных не владельцем почтового ящика. Нельзя извлекать сущности из элементов календарей, которые не являются собраниями, или собраний, организованных владельцем почтового ящика.
    
4. Вы можете извлекать сущности типа  **MeetingSuggestion** только из сообщений, но не встреч.
    
5. Вы можете извлекать URL-адреса, явно указанные в теле элемента, но не URL-адреса, встроенные в текст гиперссылок в HTML-элементе. Чтобы получать явные и встроенные URL-адреса, используйте правило  **ItemHasRegularExpressionMatch**. Укажите  **BodyAsHTML** как значение _PropertyName_ и регулярное выражение, выбирающее URL-адреса, как _RegExValue_.
    
6. Невозможно извлечь сущности из элементов в папке "Отправленные".
    
Кроме того, если вы используете правило [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx), применяются следующие условия, которые могут повлиять на ситуации, связанные с активацией надстройки:


1. При использовании правила  **ItemHasKnownEntity**Outlook сравнивает строки сущностей только на английском языке независимо от языкового стандарта по умолчанию, указанного в манифесте.
    
2. Если надстройка работает в полнофункциональном клиенте Outlook, Outlook будет применять правило  **ItemHasKnownEntity** только к первому мегабайту основного текста элемента.
    
3. Правило  **ItemHasKnownEntity** нельзя использовать для активации надстройки для элементов в папке "Отправленные".
    

## Дополнительные ресурсы



- [Создание надстроек Outlook для форм чтения](../outlook/read-scenario.md)
    
- [Извлечение строк сущности из элемента Outlook](../outlook/extract-entity-strings-from-an-item.md)
    
- [Правила активации для надстроек Outlook](../outlook/manifests/activation-rules.md)
    
- [Использование регулярных правил активации выражений для отображения надстройки Outlook](../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
- [Общие сведения о разрешениях для надстройки Outlook](../outlook/understanding-outlook-add-in-permissions.md)
    