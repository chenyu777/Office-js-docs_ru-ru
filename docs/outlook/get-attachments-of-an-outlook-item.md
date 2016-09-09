
# Получение вложений элемента Outlook с сервера

Надстройка Outlook не может передавать вложения для выбранного элемента непосредственно в удаленную службу, работающую на сервере. Вместо этого она может использовать API вложений для отправки информации о вложениях в такую удаленную службу. Затем эта служба может обратиться напрямую к серверу Exchange для получения вложений.

Чтобы отправить информацию о вложениях в удаленную службу, используйте следующие свойства и функцию:


- Свойство [Office.context.mailbox.ewsUrl](../../reference/outlook/Office.context.mailbox.md) предоставляет URL-адрес веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик. Служба использует этот URL-адрес, чтобы вызвать метод [ExchangeService.GetAttachments](http://msdn.microsoft.com/en-us/library/office/dn600509%28v=exchg.80%29.aspx) для [управляемого API EWS](http://msdn.microsoft.com/library/c2267733-6f4f-49e5-9614-1e4a24c3af1a%28Office.15%29.aspx) или операцию [GetAttachment](http://msdn.microsoft.com/en-us/library/24d10a15-b942-415e-9024-a6375708f326%28Office.15%29.aspx) для EWS.
    
- Свойство [Office.context.mailbox.item.attachments](../../reference/outlook/Office.context.mailbox.item.md) получает массив объектов [AttachmentDetails](../../reference/outlook/simple-types.md) (по одному для каждого вложения в элемент).
    
- Функция [Office.context.mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md) асинхронно вызывает сервер Exchange Server с почтовым ящиком, чтобы получить маркер обратного вызова, который клиентский сервер отправит обратно на сервер Exchange Server для проверки подлинности запроса на получение вложения.
    

## Использование API вложений


Чтобы использовать API вложений для получения вложений из почтового ящика Exchange, выполните следующие действия. 


1. Отобразите надстройку, когда пользователь просматривает сведения о встрече или сообщение, которые содержат вложение.
    
2. Получите маркер обратного вызова с сервера Exchange.
    
3. Отправьте маркер обратного вызова и сведения о вложениях в удаленную службу.
    
4. Получите вложения с сервера Exchange с помощью метода  **ExchangeService.GetAttachments** операции **GetAttachment**.
    
Более подробно каждое из этих действий рассматривается в последующих разделах с использованием кода из примера [Outlook-Add-in-JavaScript-GetAttachments](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-GetAttachments).


 >**Примечание**  Код в этих примерах был сокращен, чтобы уделить основное внимание информации о вложениях. Пример содержит дополнительный код для проверки подлинности надстройки на удаленном сервере и управления состоянием запроса.


### Активация надстройки


С помощью правила [ItemHasAttachment](http://msdn.microsoft.com/en-us/library/031db7be-8a25-5185-a9c3-93987e10c6c2%28Office.15%29.aspx) в файле манифеста надстройки вы можете обеспечить отображение этой надстройки, если выбранный элемент содержит вложения, как показано в следующем примере:


```XML
<Rule xsi:type="ItemHasAttachment" />
```


### Получение маркера обратного вызова


Объект [Office.context.mailbox](../../reference/outlook/Office.context.mailbox.md) предоставляет функцию **getCallbackTokenAsync** для получения маркера, который удаленный сервер использует для проверки подлинности на сервере Exchange Server. В следующем коде показана функция надстройки, которая отправляет асинхронный запрос на получение маркера обратного вызова, а также функция обратного вызова, получающая ответ. Маркер обратного вызова сохраняется в объекте запроса службы, описанном в следующем разделе.


```
function getAttachmentToken() {
    if (serviceRequest.attachmentToken == "" {
        Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
    }
};
function attachmentTokenCallback(asyncResult, userContext) {
    if (asyncResult.status === "succeeded") {
        // Cache the result from the server.
        serviceRequest.attachmentToken = asyncResult.value;
        serviceRequest.state = 3;
        testAttachments();
    } else {
        showToast("Error", "Could not get callback token: " + asyncResult.error.message);
    }
};
```


### Отправка информации о вложениях в удаленную службу


От удаленной службы, которую вызывает ваша надстройка, зависит способ отправки информации о вложениях в эту службу. В данном примере такой удаленной службой является приложение веб-API, созданное с помощью Visual Studio 2013. Удаленная служба ожидает получить сведения о вложениях в объекте JSON. Следующий код инициализирует объект, содержащий информацию о вложениях.


```
// Initialize a context object for the add-in.
//   Set the fields that are used on the request
//   object to default values.
serviceRequest = new Object();
serviceRequest.attachmentToken = "";
serviceRequest.ewsUrl = Office.context.mailbox.ewsUrl;
serviceRequest.attachments = new Array();
```

Свойство  `Office.context.mailbox.item.attachments` содержит коллекцию объектов **AttachmentDetails** — по одному для каждого вложения элемента. В большинстве случаев надстройка может передать в удаленную службу только свойство идентификатора вложения, принадлежащее объекту **AttachmentDetails**. Если удаленной службе требуются дополнительные сведения о вложении, вы можете передать объект  **AttachmentDetails** полностью или частично. Указанный ниже код определяет метод, который помещает весь массив **AttachmentDetails** в объект `serviceRequest` и отправляет запрос в удаленную службу.




```js
    function makeServiceRequest() {
      // Format the attachment details for sending.
      for (var i = 0; i < mailbox.item.attachments.length; i++) {
        serviceRequest.attachments[i] = JSON.parse(JSON.stringify(mailbox.item.attachments[i].$0_0));
      }

      $.ajax({
        url: '../../api/Default',
        type: 'POST',
        data: JSON.stringify(serviceRequest),
        contentType: 'application/json;charset=utf-8'
      }).done(function (response) {
        if (!response.isError) {
          var names = "<h2>Attachments processed using " +
                        serviceRequest.service +
                        ": " +
                        response.attachmentsProcessed +
                        "</h2>";
          for (i = 0; i < response.attachmentNames.length; i++) {
            names += response.attachmentNames[i] + "<br />";
          }
          document.getElementById("names").innerHTML = names;
        } else {
          app.showNotification("Runtime error", response.message);
        }
      }).fail(function (status) {

      }).always(function () {
        $('.disable-while-sending').prop('disabled', false);
      })
    };

```


### Получение вложений с сервера Exchange


Ваша удаленная служба может использовать метод [GetAttachments](http://msdn.microsoft.com/en-us/library/office/dn600509%28v=exchg.80%29.aspx) управляемого API веб-служб Exchange или операцию [GetAttachment](http://msdn.microsoft.com/library/24d10a15-b942-415e-9024-a6375708f326%28Office.15%29.aspx) управляемого API веб-служб Exchange для получения вложений с сервера. Приложению-службе необходимы два объекта, чтобы выполнить десериализацию строки JSON в объекты .NET Framework, которые можно использовать на сервере. В следующем коде показаны определения объектов десериализации.


```C#



namespace AttachmentsSample
{
  public class AttachmentSampleServiceRequest
  {
    public string attachmentToken { get; set; }
    public string ewsUrl { get; set; }
    public string service { get; set; }
    public AttachmentDetails [] attachments { get; set; }
  }

  public class AttachmentDetails
  {
    public string attachmentType { get; set; }
    public string contentType { get; set; }
    public string id { get; set; }
    public bool isInline { get; set; }
    public string name { get; set; }
    public int size { get; set; }
  }
}
```


#### Использование управляемого API EWS для получения вложений

Если вы используете в своей удаленной службе [управляемый API EWS](http://go.microsoft.com/fwlink/?LinkID=255472), вы можете воспользоваться методом [GetAttachments](http://msdn.microsoft.com/en-us/library/office/dn600509%28v=exchg.80%29.aspx), который создаст, отправит и получит SOAP-запрос EWS для получения вложений. Рекомендуем использовать управляемый API EWS, поскольку он требует меньше строк кода и обеспечивает более интуитивный интерфейс для вызовов EWS. Приведенный ниже код отправляет один запрос на получение всех вложений, а также возвращает количество и имена обработанных вложений.


```C#
    private AttachmentSampleServiceResponse GetAtttachmentsFromExchangeServerUsingEWSManagedApi(AttachmentSampleServiceRequest request)
    {
      var attachmentsProcessedCount = 0;
      var attachmentNames = new List<string>();

      // Create an ExchangeService object, set the credentials and the EWS URL.
      ExchangeService service = new ExchangeService();
      service.Credentials = new OAuthCredentials(request.attachmentToken);
      service.Url = new Uri(request.ewsUrl);

      var attachmentIds = new List<string>();

      foreach (AttachmentDetails attachment in request.attachments)
      {
        attachmentIds.Add(attachment.id);
      }

      // Call the GetAttachments method to retrieve the attachments on the message.
      // This method results in a GetAttachments EWS SOAP request and response
      // from the Exchange server.
      var getAttachmentsResponse =
        service.GetAttachments(attachmentIds.ToArray(),
                               null,
                               new PropertySet(BasePropertySet.FirstClassProperties,
                                               ItemSchema.MimeContent));

      if (getAttachmentsResponse.OverallResult == ServiceResult.Success)
      {
        foreach (var attachmentResponse in getAttachmentsResponse)
        {
          attachmentNames.Add(attachmentResponse.Attachment.Name);

          // Write the content of each attachment to a stream.
          if (attachmentResponse.Attachment is FileAttachment)
          {
            FileAttachment fileAttachment = attachmentResponse.Attachment as FileAttachment;
            Stream s = new MemoryStream(fileAttachment.Content);
            // Process the contents of the attachment here.
          }

          if (attachmentResponse.Attachment is ItemAttachment)
          {
            ItemAttachment itemAttachment = attachmentResponse.Attachment as ItemAttachment;
            Stream s = new MemoryStream(itemAttachment.Item.MimeContent.Content);
            // Process the contents of the attachment here.
          }

          attachmentsProcessedCount++;
        }
      }

      // Return the names and number of attachments processed for display
      // in the add-in UI.
      var response = new AttachmentSampleServiceResponse();
      response.attachmentNames = attachmentNames.ToArray();
      response.attachmentsProcessed = attachmentsProcessedCount;

      return response;
    }


```


#### Использование EWS для получения вложений

Если вы используете EWS в своей удаленной службе, то чтобы получить вложения с сервера Exchange, вам требуется создать запрос SOAP [GetAttachment](http://msdn.microsoft.com/library/24d10a15-b942-415e-9024-a6375708f326%28Office.15%29.aspx). Следующий код возвращает строку, которая предоставляет запрос SOAP. Удаленная служба использует метод  **String.Format** для вставки идентификатора вложения в эту строку.


```C#
    private const string GetAttachmentSoapRequest =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
xmlns:xsd=""http://www.w3.org/2001/XMLSchema""
xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
<soap:Header>
<t:RequestServerVersion Version=""Exchange2013"" />
</soap:Header>
  <soap:Body>
    <GetAttachment xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""
    xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
      <AttachmentShape/>
      <AttachmentIds>
        <t:AttachmentId Id=""{0}""/>
      </AttachmentIds>
    </GetAttachment>
  </soap:Body>
</soap:Envelope>";

```

Наконец, следующий метод обеспечивает использование запроса  **GetAttachment** EWS для получения вложений с сервера Exchange. При такой реализации каждому вложению соответствует отдельный запрос, и возвращается число обработанных вложений. Каждый ответ обрабатывается в отдельном методе **ProcessXmlResponse**, определение которого приведено ниже.




```C#
    private AttachmentSampleServiceResponse GetAttachmentsFromExchangeServerUsingEWS(AttachmentSampleServiceRequest request)
    {
      var attachmentsProcessedCount = 0;
      var attachmentNames = new List<string>();

      foreach (var attachment in request.attachments)
      {
        // Prepare a web request object.
        HttpWebRequest webRequest = WebRequest.CreateHttp(request.ewsUrl);
        webRequest.Headers.Add("Authorization",
          string.Format("Bearer {0}", request.attachmentToken));
        webRequest.PreAuthenticate = true;
        webRequest.AllowAutoRedirect = false;
        webRequest.Method = "POST";
        webRequest.ContentType = "text/xml; charset=utf-8";

        // Construct the SOAP message for the GetAttachment operation.
        byte[] bodyBytes = Encoding.UTF8.GetBytes(
          string.Format(GetAttachmentSoapRequest, attachment.id));
        webRequest.ContentLength = bodyBytes.Length;

        Stream requestStream = webRequest.GetRequestStream();
        requestStream.Write(bodyBytes, 0, bodyBytes.Length);
        requestStream.Close();

        // Make the request to the Exchange server and get the response.
        HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

        // If the response is okay, create an XML document from the reponse
        // and process the request.
        if (webResponse.StatusCode == HttpStatusCode.OK)
        {
          var responseStream = webResponse.GetResponseStream();

          var responseEnvelope = XElement.Load(responseStream);

          // After creating a memory stream containing the contents of the 
          // attachment, this method writes the XML document to the trace output.
          // Your service would perform it's processing here.
          if (responseEnvelope != null)
          {
            var processResult = ProcessXmlResponse(responseEnvelope);
            attachmentNames.Add(string.Format("{0} {1}", attachment.name, processResult));

          }

          // Close the response stream.
          responseStream.Close();
          webResponse.Close();

        }
        // If the response is not OK, return an error message for the 
        // attachment.
        else
        {
          var errorString = string.Format("Attachment \"{0}\" could not be processed. " +
            "Error message: {1}.", attachment.name, webResponse.StatusDescription);
          attachmentNames.Add(errorString);
        }
        attachmentsProcessedCount++;
      }

      // Return the names and number of attachments processed for display
      // in the add-in UI.
      var response = new AttachmentSampleServiceResponse();
      response.attachmentNames = attachmentNames.ToArray();
      response.attachmentsProcessed = attachmentsProcessedCount;

      return response;
    }

```

Каждый ответ от операции  **GetAttachment** отправляется методу **ProcessXmlResponse**. Этот метод проверяет ответ на наличие ошибок. Если ошибки не найдены, он обрабатывает вложенные файлы и элементы. Метод  **ProcessXmlResponse** выполняет большую часть работы по обработке вложения.




```C#
    // This method processes the response from the Exchange server.
    // In your application the bulk of the processing occurs here.
    private string ProcessXmlResponse(XElement responseEnvelope)
    {
      // First, check the response for web service errors.
      var errorCodes = from errorCode in responseEnvelope.Descendants
                       ("{http://schemas.microsoft.com/exchange/services/2006/messages}ResponseCode")
                       select errorCode;
      // Return the first error code found.
      foreach (var errorCode in errorCodes)
      {
        if (errorCode.Value != "NoError")
        {
          return string.Format("Could not process result. Error: {0}", errorCode.Value);
        }
      }

      // No errors found, proceed with processing the content.
      // First, get and process file attachments.
      var fileAttachments = from fileAttachment in responseEnvelope.Descendants
                        ("{http://schemas.microsoft.com/exchange/services/2006/types}FileAttachment")
                            select fileAttachment;
      foreach(var fileAttachment in fileAttachments)
      {
        var fileContent = fileAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Content");
        var fileData = System.Convert.FromBase64String(fileContent.Value);
        var s = new MemoryStream(fileData);
        // Process the file attachment here. 
      }

      // Second, get and process item attachments.
      var itemAttachments = from itemAttachment in responseEnvelope.Descendants
                            ("{http://schemas.microsoft.com/exchange/services/2006/types}ItemAttachment")
                            select itemAttachment;
      foreach(var itemAttachment in itemAttachments)
      {
        var message = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Message");
        if (message != null)
        {
         // Process a message here.
          break;
        }
        var calendarItem = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}CalendarItem");
        if (calendarItem != null)
        {
          // Process calendar item here.
          break;
        }
        var contact = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Contact");
        if (contact != null)
        {
          // Process contact here.
          break;
        }
        var task = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Tontact");
        if (task != null)
        {
          // Process task here.
          break;
        }
        var meetingMessage = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingMessage");
        if (meetingMessage != null)
        {
          // Process meeting message here.
          break;
        }
        var meetingRequest = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingRequest");
        if (meetingRequest != null)
        {
          // Process meeting request here.
          break;
        }
        var meetingResponse = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingResponse");
        if (meetingResponse != null)
        {
          // Process meeting response here.
          break;
        }
        var meetingCancellation = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingCancellation");
        if (meetingCancellation != null)
        {
          // Process meeting cancellation here.
          break;
        }
      }
     
      return string.Empty;
    }

```


## Дополнительные ресурсы



- [Создание надстроек Outlook для форм чтения](../outlook/read-scenario.md)
    
- [Сведения об управляемом API EWS, EWS и веб-службах в Exchange](http://msdn.microsoft.com/library/0bc6f81d-cc10-42b0-ba5d-6f22ff55d51c%28Office.15%29.aspx)
    
- [Начало работы с клиентскими приложениями, использующими управляемый API EWS](http://msdn.microsoft.com/library/c2267733-6f4f-49e5-9614-1e4a24c3af1a%28Office.15%29.aspx)
    
- [Примеры кода Outlook Power Hour](https://github.com/OfficeDev/Outlook-Power-Hour-Code-Samples):  `MyAttachments` и `AttachmentsDemo`
    
