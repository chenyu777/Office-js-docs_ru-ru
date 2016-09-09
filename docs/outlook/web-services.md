
# Вызов веб-служб из надстройки Outlook

Ваша надстройка может использовать веб-службы Exchange (EWS) на компьютере с Exchange Server 2013; веб-службу, доступную на сервере, предоставляющем исходное расположение для пользовательского интерфейса надстройки; или веб-службу, доступную через Интернет. В этой статье приведен пример того, как надстройка Outlook может запрашивать данные из EWS.

Способы вызова веб-службы различаются в зависимости от расположения службы. В таблице 1 приведены различные способы вызова веб-службы в зависимости от расположения.


**Таблица 1. Способы вызова веб-служб из надстройки Outlook**


|**Расположение веб-службы**|**Способ вызова веб-службы**|
|:-----|:-----|
|Сервер Exchange, на котором размещен почтовый ящик клиента|Используйте метод [makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) для вызова операций EWS, поддерживаемых надстройками. Сервер Exchange Server, на котором размещен почтовый ящик, также предоставляет доступ к EWS.|
|Веб-сервер, предоставляющий исходное расположение для пользовательского интерфейса надстроек.|Вызывайте веб-службу с помощью стандартных методик JavaScript. Код JavaScript в пределах пользовательского интерфейса работает в контексте веб-сервера, предоставляющего пользовательский интерфейс. Поэтому он сможет вызывать веб-службы на этом сервере, не создавая ошибки межсайтового скрипта.|
|Все другие расположения|Создайте прокси для веб-службы на веб-сервере, предоставляющем исходное расположение для пользовательского интерфейса. Если не указать прокси, надстройка не запустится из-за ошибок межсайтовых сценариев. Один из способов указать такой прокси — это использовать JSON/P. Дополнительные сведения см. в статье [Конфиденциальность и безопасность надстроек для Office](../../docs/develop/privacy-and-security.md).|

## Получение доступа к операциям веб-служб Exchange с помощью метода makeEwsRequestAsync


С помощью метода [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) вы можете отправить запрос EWS на сервер Exchange Server, на котором размещается почтовый ящик пользователя.

Веб-службы Exchange поддерживают различные операции на сервере Exchange. Например, операции копирования, поиска, обновления или отправки на уровне элемента, а также операции создания, получения или обновления на уровне папки. Чтобы выполнить операцию веб-служб Exchange, создайте для нее SOAP-запрос в формате XML. После завершения операции будет возвращен SOAP-ответ в формате XML с необходимыми данными. SOAP-запросы к веб-службам Exchange и их SOAP-ответы соответствуют схеме, определенной в файле Messages.xsd. Как и другие файлы схемы веб-служб Exchange, файл Message.xsd расположен в виртуальном каталоге IIS, в котором размещены веб-службы Exchange. 

Чтобы использовать метод  **makeEwsRequestAsync** для запуска операции веб-служб Exchange, предоставьте вот что:


- XML-код SOAP-запроса для соответствующей операции EWS в качестве аргумента для параметра  _data_;
    
- метод обратного вызова (в качестве аргумента  _callback_);
    
- все необязательные входные данные для этого метода обратного вызова (в качестве аргумента  _userContext_).
    
Когда SOAP-запрос к веб-службам Exchange выполнен, Outlook вызывает метод обратного вызова с аргументом в виде объекта [AsyncResult](../../reference/outlook/simple-types.md). Такой метод позволяет получить доступ к двум свойствам объекта  **AsyncResult**. Вот они: свойство  **value**, содержащее SOAP-ответ в формате XML (получен при выполнении операции веб-служб Exchange), и свойство  **asyncContext** (необязательное), содержащее все данные, переданные в виде параметра **userContext**. Как правило, затем метод обратного вызова анализирует XML-код в SOAP-ответе, чтобы получить необходимые сведения и обработать их соответствующим образом.


## Советы по анализу ответов веб-служб Exchange


При анализе SOAP-ответа, полученного при выполнении операции веб-служб Exchange, обратите внимание на приведенные ниже особенности, связанные с типом браузера.


- При использовании метода DOM  **getElementsByTagName** укажите префикс имени тега, чтобы включить поддержку браузера Internet Explorer.
    
     Метод **getElementsByTagName** работает по-разному в зависимости от типа браузера. Например, ответ EWS может содержать следующий XML-код (отформатированный и сокращенный для наглядности):
    
```XML
      <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
    PropertyName="MyProperty" 
    PropertyType="String"/>
    <t:Value>{
    ...
    }</t:Value></t:ExtendedProperty>
```

 Приведенный ниже код позволит получить XML-код, заключенный в теги **ExtendedProperty**, в таком браузере, как Chrome.

```js
    var mailbox = Office.context.mailbox;
    mailbox.makeEwsRequestAsync(mailbox.item.itemId), function(result) {
        var response = $.parseXML(result.value);
        var extendedProps = response.getElementsByTagName("ExtendedProperty");
```


   
 В Internet Explorer необходимо включить префикс `t:` имени тега, как показано ниже:

```js
    var mailbox = Office.context.mailbox;
    mailbox.makeEwsRequestAsync(mailbox.item.itemId), function(result) {
        var response = $.parseXML(result.value);
        var extendedProps = response.getElementsByTagName("t:ExtendedProperty");
```

- Чтобы получить содержимое тега в ответе веб-служб Exchange, используйте свойство DOM  **textContent**:
    
```
      content = $.parseJSON(value.textContent);
```

 Другие свойства, например **innerHTML** могут не работать в Internet Explorer для некоторых тегов в ответе веб-служб Exchange.
    

## Пример


Следующий пример вызывает  **makeEwsRequestAsync** для использования операции [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx), чтобы получить тему элемента. Этот пример содержит три следующие функции:


-  `getSubjectRequest` принимает в качестве входных данных идентификатор элемента и возвращает XML-код SOAP-запроса, чтобы вызвать операцию **GetItem** для заданного элемента.
    
-  `sendRequest` вызывает функцию `getSubjectRequest`, чтобы получить SOAP-запрос для выбранного элемента. Затем передает этот запрос и метод обратного вызова, `callback`, в **makeEwsRequestAsync**, чтобы получить тему выбранного элемента.
    
-  `callback` обрабатывает SOAP-ответ, включающий тему и другие сведения об указанном элементе.
    

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item. 
   var result = 
'<?xml version="1.0" encoding="utf-8"?>' +
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
'               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
'               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
'               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
'  <soap:Header>' +
'    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
'  </soap:Header>' +
'  <soap:Body>' +
'    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
'      <ItemShape>' +
'        <t:BaseShape>IdOnly</t:BaseShape>' +
'        <t:AdditionalProperties>' +
'            <t:FieldURI FieldURI="item:Subject"/>' +
'        </t:AdditionalProperties>' +
'      </ItemShape>' +
'      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
'    </GetItem>' +
'  </soap:Body>' +
'</soap:Envelope>';

   return result;
}





function sendRequest() {
   // Create a local variable that contains the mailbox.
   var mailbox = Office.context.mailbox;

   mailbox.makeEwsRequestAsync(getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.context;

   // Process the returned response here.
}


```


## Операции веб-служб Exchange, которые надстройки поддерживают


Надстройки Outlook могут получать доступ к подмножеству операций EWS с помощью метода  **makeEwsRequestAsync**. Если вы не знакомы с операциями EWS и не знаете, как использовать метод  **makeEwsRequestAsync** для доступа к операциям, начните с примера SOAP-запроса для настройки аргумента _data_. В следующем примере показано, как применить метод  **makeEwsRequestAsync**:


1. В XML-коде замените все идентификаторы элементов и релевантные атрибуты операций EWS на соответствующие значения.
    
2. Включите SOAP-запрос в качестве аргумента для параметра  _data_ метода **makeEwsRequestAsync**.
    
3. Укажите метод обратного вызова и вызовите  **makeEwsRequestAsync**.
    
4. В методе обратного вызова проверьте результаты операции в SOAP-ответе.
    
5. Используйте результаты операции EWS в соответствии с вашими потребностями.
    
В следующей таблице указаны операции EWS, которые надстройки поддерживают. Чтобы просмотреть примеры SOAP-запросов и SOAP-ответов, выберите ссылку для каждой операции. Дополнительные сведения об операциях EWS см. в статье [Операции EWS в Exchange](http://msdn.microsoft.com/library/cf6fd871-9a65-4f34-8557-c8c71dd7ce09%28Office.15%29.aspx).


**Таблица 2. Поддерживаемые операции EWS**


|**Операция служб EWS**|**Описание**|
|:-----|:-----|
|[CopyItem Operation](http://msdn.microsoft.com/library/bcc68f9e-d511-4c29-bba6-ed535524624a%28Office.15%29.aspx)|Копирует выбранные элементы и размещает новые элементы в выделенной папке в хранилище Exchange.|
|[CreateFolder Operation](http://msdn.microsoft.com/library/6f6c334c-b190-4e55-8f0a-38f2a018d1b3%28Office.15%29.aspx)|Создает папки в выбранном расположении в хранилище Exchange.|
|[CreateItem Operation](http://msdn.microsoft.com/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx)|Создает заданные элементы в хранилище Exchange.|
|[FindConversation Operation](http://msdn.microsoft.com/library/2384908a-c203-45b6-98aa-efd6a4c23aac%28Office.15%29.aspx)|Перечисляет список бесед в определенной папке в хранилище Exchange.|
|[FindFolder Operation](http://msdn.microsoft.com/library/7a9855aa-06cc-45ba-ad2a-645c15b7d031%28Office.15%29.aspx)|Ищет вложенные папки заданной папки и возвращает набор свойств, описывающих вложенные папки.|
|[FindItem Operation](http://msdn.microsoft.com/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx)|Определяет элементы, расположенные в определенной папке в хранилище Exchange.|
|[GetConversationItems operation](http://msdn.microsoft.com/library/8ae00a99-b37b-4194-829c-fe300db6ab99%28Office.15%29.aspx)|Получает один или несколько наборов элементов, упорядоченных в узлы в беседе.|
|[GetFolder Operation](http://msdn.microsoft.com/library/355bcf93-dc71-4493-b177-622afac5fdb9%28Office.15%29.aspx)|Получает определенные свойства и содержимое папок из хранилища Exchange.|
|[GetItem Operation](http://msdn.microsoft.com/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx)|Получает определенные свойства и содержимое элементов из хранилища Exchange.|
|[MarkAsJunk Operation](http://msdn.microsoft.com/library/1f71f04d-56a9-4fee-a4e7-d1034438329e%28Office.15%29.aspx)|Перемещает сообщения электронной почты в папку "Нежелательная почта" и соответствующим образом добавляет или удаляет отправителей сообщений в списке заблокированных отправителей.|
|[MoveItem Operation](http://msdn.microsoft.com/library/dcf40fa7-7796-4a5c-bf5b-7a509a18d208%28Office.15%29.aspx)|Перемещает элементы в одну целевую папку в хранилище Exchange.|
|[SendItem Operation](http://msdn.microsoft.com/library/337b89ef-e1b7-45ed-92f3-8abe4200e4c7%28Office.15%29.aspx)|Отправляет сообщения электронной почты, расположенные в хранилище Exchange.|
|[Операцию UpdateFolder](http://msdn.microsoft.com/library/3494c996-b834-4813-b1ca-d99642d8b4e7%28Office.15%29.aspx)|Изменяет свойства существующих папок в хранилище Exchange.|
|[UpdateItem Operation](http://msdn.microsoft.com/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx)|Изменяет свойства существующих элементов в хранилище Exchange.|

## Разрешения и проверка подлинности для метода makeEwsRequestAsync


Когда используется метод  **makeEwsRequestAsync**, запрос проходит проверку подлинности с помощью данных учетной записи электронной почты текущего пользователя. Метод  **makeEwsRequestAsync** автоматически управляет учетными записями, поэтому в запросе не требуется предоставлять учетные данные для проверки подлинности.


 >
  **Примечание**  Администратор сервера должен использовать командлет [New-WebServicesVirtualDirctory](http://technet.microsoft.com/en-us/library/bb125176.aspx) или [Set-WebServicesVirtualDirecory](http://technet.microsoft.com/en-us/library/aa997233.aspx), чтобы установить для параметра  _OAuthAuthentication_ значение **true** в каталоге Client Access Server EWS, чтобы метод **makeEwsRequestAsync** мог выполнять запросы EWS.

В манифесте надстройки должно быть указано разрешение **ReadWriteMailbox**, чтобы эта надстройка могла использовать метод **makeEwsRequestAsync**. Сведения об использовании разрешения **ReadWriteMailbox** см. в разделе [Разрешение ReadWriteMailbox](../outlook/understanding-outlook-add-in-permissions.md#readwritemailbox-permission) статьи [Указание разрешений для доступа надстройки Outlook к почтовому ящику пользователя](../outlook/understanding-outlook-add-in-permissions.md).


## Дополнительные ресурсы



- [Надстройки Outlook](../outlook/outlook-add-ins.md)
    
- [Конфиденциальность и безопасность надстроек для Office](../../docs/develop/privacy-and-security.md)
    
- [Решение ограничений политик одинакового происхождения в надстройках для Office](../../docs/develop/addressing-same-origin-policy-limitations.md)
    
- [Справка по веб-служб Exchange для Exchange](http://msdn.microsoft.com/library/2a873474-1bb2-4cb1-a556-40e8c4159f4a%28Office.15%29.aspx)
    
- [Приложения электронной почты для Outlook и EWS в Exchange](http://msdn.microsoft.com/library/821c8eb9-bb58-42e8-9a3a-61ca635cba59%28Office.15%29.aspx)
    
Сведения о создании внутренних служб для надстроек с помощью веб-API ASP.NET см. в следующих статьях:


- [Создание веб-службы надстройки для Office с использованием веб-API ASP.NET](http://blogs.msdn.com/b/officeapps/archive/2013/06/10/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api.aspx)
    
- [Основы создания службы HTTP с использованием веб-API ASP.NET](http://www.asp.net/web-api)
    
