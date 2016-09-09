
# Проверка маркера удостоверения Exchange

Надстройка Outlook может отправить пользователю маркер удостоверения. Чтобы можно было доверять запросу, необходимо проверить маркер и убедиться, что он создан на нужном сервере Exchange. В примерах, приведенных в этой статье, показано, как проверить маркер удостоверения Exchange с помощью объекта проверки, написанного на языке C#. Для проверки можно использовать любой язык программирования. Действия, необходимые для проверки маркера, описаны в статье [Веб-маркер JSON (JWT), интернет-проект](http://self-issued.info/docs/draft-goland-json-web-token-00.mdl). 

Рекомендуется использовать процесс проверки маркера удостоверения и получения уникального идентификатора пользователя, включающий четыре этапа. Первый этап: извлечение веб-маркера JSON (JWT) из строки, закодированной в формате URL-адреса Base64. Второй этап: проверка правильности маркера, то есть его предназначения для вашей надстройки Outlook, его актуальности и возможности извлечения допустимого URL-адреса для документа метаданных проверки подлинности. Затем необходимо получить документ метаданных проверки подлинности с сервера Exchange и проверить подпись, приложенную к маркеру удостоверения. Наконец, необходимо вычислить уникальный идентификатор пользователя путем хэширования идентификатора Exchange пользователя с URL-адресом документа метаданных проверки подлинности. В целом процесс может показаться сложным, однако каждый отдельный этап довольно прост. Решение, содержащее эти примеры, можно скачать с веб-страницы [Outlook-Add-in-JavaScript-ValidateIdentityToken](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ValidateIdentityToken).
 




## Настройка проверки маркера удостоверения


Примеры кода в этой статье зависят от Windows Identity Foundation (WIF) вместе с DLL-библиотекой, расширяющей WIF обработчиками для маркеров JSON. Можно загрузить необходимые сборки из следующих мест:


- [Windows Identity Foundation](http://msdn.microsoft.com/en-us/security/aa570351)
    
- [Windows.IdentityModel.Extensions.dll для 32-разрядных приложений](http://download.microsoft.com/download/0/1/D/01D06854-CA0C-46F1-ADBA-EBF86010DCC6/MicrosoftIdentityExtensions-32.msi)
    
- [Windows.IdentityModel.Extensions.dll для 64-разрядных приложений](http://download.microsoft.com/download/0/1/D/01D06854-CA0C-46F1-ADBA-EBF86010DCC6/MicrosoftIdentityExtensions-64.msi)
    

## Извлечение веб-маркера JSON


Фабричный метод **Decode** разбивает JWT с сервера Exchange на три строки, которые составляют маркер, а затем с помощью метода **Base64Decode** (показан во втором примере) декодирует заголовок JWT и полезные данные в строки JSON. Строки передаются в конструктор **JsonToken**, который проверяет содержимое JWT и возвращает новый экземпляр объекта **JsonToken**.


```C#
    public static JsonToken Decode(string rawToken)
    {
      string[] tokenParts = rawToken.Split('.');

      if (tokenParts.Length != 3)
      {
        throw new ApplicationException("Token must have three parts separated by '.' characters.");
      }

      string encodedHeader = tokenParts[0];
      string encodedPayload = tokenParts[1];
      string signature = tokenParts[2];

      string decodedHeader = Base64Decode(encodedHeader);
      string decodedPayload = Base64Decode(encodedPayload);

      JavaScriptSerializer serializer = new JavaScriptSerializer();

      Dictionary<string, string> header = serializer.Deserialize<Dictionary<string, string>>(decodedHeader);
      Dictionary<string, string> payload = serializer.Deserialize<Dictionary<string, string>>(decodedPayload);

      return new JsonToken(header, payload, signature);
    }
```

Метод **Base64Decode** реализует логику декодирования, описанную в приложении "Заметки о реализации кодирования Base64Decode без заполнения" к статье [Веб-маркер JSON (JWT), интернет-проект](http://self-issued.info/docs/draft-goland-json-web-token-00.mdl).




```C#
    public static Encoding TextEncoding = Encoding.UTF8;

    private static char Base64PadCharacter = '=';
    private static char Base64Character62 = '+';
    private static char Base64Character63 = '/';
    private static char Base64UrlCharacter62 = '-';
    private static char Base64UrlCharacter63 = '_';

    private static byte[] DecodeBytes(string arg)
    {
      if (String.IsNullOrEmpty(arg))
      {
        throw new ApplicationException("String to decode cannot be null or empty.");
      }

      StringBuilder s = new StringBuilder(arg);
      s.Replace(Base64UrlCharacter62, Base64Character62);
      s.Replace(Base64UrlCharacter63, Base64Character63);

      int pad = s.Length % 4;
      s.Append(Base64PadCharacter, (pad == 0) ? 0 : 4 - pad);

      return Convert.FromBase64String(s.ToString());
    }

    private static string Base64Decode(string arg)
    {
      return TextEncoding.GetString(DecodeBytes(arg));
    }
```


## Обработка JWT


Конструктор для объекта **JsonToken** проверяет структуру и содержимое JWT для определения допустимости этого маркера. Рекомендуется делать это перед запросом документа метаданных проверки подлинности. Если JWT не содержит необходимых утверждений или истекло время его существования, вы можете избежать вызова к серверу Exchange и соответствующей задержки.

Конструктор вызывает служебные методы, чтобы определить, имеются ли какие-либо утверждения и находятся ли они в необходимой области. При наличии проблемы служебный метод создает исключение приложения. При отсутствии исключений свойство **IsValid** принимает значение **true**, и маркер можно использовать для проверки подписи.

Каждый служебный метод описывается далее в этой статье.




```C#
    public JsonToken(Dictionary<string, string> header, Dictionary<string, string> payload, string signature)
    {

      // Assume that the token is invalid to start out.
      this.IsValid = false;

      // Set the private dictionaries that contain the claims.
      this.headerClaims = header;
      this.payloadClaims = payload;
      this.signature = signature;

      // If there is no "appctx" claim in the token, throw an ApplicationException.
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.AppContext))
      {
        throw new ApplicationException(String.Format("The {0} claim is not present.", AuthClaimTypes.AppContext));
      }

      appContext = new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(payload[AuthClaimTypes.AppContext]);


      // Validate the header fields.
      this.ValidateHeader();

      // Determine whether the token is within its valid time.
      this.ValidateLifetime();

      // Validate that the token was sent to the correct URL.
      this.ValidateAudience();

      // Validate the token version.
      this.ValidateVersion();

      // Make sure that the appctx contains an authentication
      // metadata location.
      this.ValidateMetadataLocation();

      // If the token passes all the validation checks, we
      // can assume that it is valid.
      this.IsValid = true;
    }
```


### Метод ValidateHeader

Метод **ValidateHeader** проверяет, есть ли в заголовке маркера нужные утверждения и имеют ли они правильные значения. Необходимо задать заголовок, как показано ниже; в противном случае метод создаст исключение приложения и завершит работу.

```js
{ "typ" : "JWT", "alg" : "RS256", "x5t" : "<thumbprint>" }
```

```C#
    private void ValidateHeaderClaim(string key, string value)
    {
      if (!this.headerClaims.ContainsKey(key))
      {
        throw new ApplicationException(String.Format("Header does not contain \"{0}\" claim.", key));
      }

      if (!value.Equals(this.headerClaims[key]))
      {
        throw new ApplicationException(String.Format("\"{0}\" claim must be \"{0}\".", key, value));
      }
    }

    private void ValidateHeader()
    {
      ValidateHeaderClaim(AuthClaimTypes.TokenType, Config.TokenType);
      ValidateHeaderClaim(AuthClaimTypes.Algorithm, Config.Algorithm);
    
      if (!this.headerClaims.ContainsKey(AuthClaimTypes.x509Thumprint))
      {
        throw new ApplicationException(String.Format("Header does not contain \"{0}\" claim.", AuthClaimTypes.x509Thumprint));
      }
    }


```


### Метод ValidateLifetime

В маркере JWT предоставлены две даты: "nbf" ("не до") предоставляет дату и время, когда маркер становится допустимым, и "exp" — истечение срока действия маркера. Допустимыми считаются маркеры, приходящиеся на промежуток времени между этими двумя датами. Чтобы учесть незначительные различия в настройке часов на сервере и клиенте, данный метод будет проверять маркеры до пяти минут до и пяти минут после времени, указанного в маркере.


```C#
    private void ValidateLifetime()
    {
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.ValidFrom))
      {
        throw new ApplicationException(
          String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.ValidFrom));
      }

      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.ValidTo))
      {
        throw new ApplicationException(
          String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.ValidTo));
      }

      DateTime unixEpoch = new DateTime(1970, 1, 1, 0, 0, 0,DateTimeKind.Utc);

      TimeSpan padding = new TimeSpan(0, 5, 0);

      DateTime validFrom = unixEpoch.AddSeconds(int.Parse(this.payloadClaims[AuthClaimTypes.ValidFrom]));
      DateTime validTo = unixEpoch.AddSeconds(int.Parse(this.payloadClaims[AuthClaimTypes.ValidTo]));

      DateTime now = DateTime.UtcNow;

      if (now < (validFrom - padding))
      {
        throw new ApplicationException(String.Format("The token is not valid until {0}.", validFrom));
      }

      if (now > (validTo + padding))
      {
        throw new ApplicationException(String.Format("The token is not valid after {0}.", validFrom));
      }
    }
```

Даты **validFrom** (nbf) и **validTo** (exp) передаются в виде количества секунд с начала "эпохи Unix" (то есть с 1 января 1970 г.). Дата и время вычисляются с использованием времени в формате UTC, чтобы избежать проблем с разницей в часовых поясах между сервером Exchange и сервером, выполняющим код проверки.


### Метод ValidateAudience

Маркер удостоверения является допустимым только для запросившей его надстройки. Метод **ValidateAudience** проверяет утверждение аудитории в маркере, чтобы удостовериться, что он соответствует ожидаемому URL-адресу надстройки Outlook.


```C#
    private void ValidateAudience()
    {
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.Audience))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the application context.", AuthClaimTypes.Audience));
      }

      string location = Config.Audience.Replace("/", "-").Replace("\\", "-");
      string audience = this.payloadClaims[AuthClaimTypes.Audience].Replace("/", "-").Replace("\\", "-");

      if (!location.Equals(audience))
      {
        throw new ApplicationException(String.Format(
          "The audience URL does not match. Expected {0}; got {1}.",
          Config.Audience, this.payloadClaims[AuthClaimTypes.Audience]));
      }
    }

```


### Метод ValidateVersion

Метод **ValidateVersion** проверяет версию маркера удостоверения и соответствие маркера ожидаемой версии. Различные версии маркера могут содержать разные утверждения. Проверка версии гарантирует присутствие ожидаемых утверждений в маркере удостоверения.


```js
    private void ValidateVersion()
    {
      if (!this.appContext.ContainsKey(AuthClaimTypes.MsExchExtensionVersion))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.MsExchExtensionVersion));
      }

      if (!Config.Version.Equals(this.appContext[AuthClaimTypes.MsExchExtensionVersion]))
      {
        throw new ApplicationException(String.Format(
          "The version does not match. Expected {0}; got {1}.",
          Config.Version, this.appContext[AuthClaimTypes.MsExchExtensionVersion]));
      }
    }

```


### Метод ValidateMetadataLocation

Объект метаданных проверки подлинности, хранящийся на сервере Exchange, содержит сведения, которые необходимы для проверки подписи, включенной в маркер удостоверения. Метод **ValidateMetadataLocation** обеспечивает наличие утверждения URL-адреса метаданных проверки подлинности в маркере удостоверения, при этом фактическая проверка подписи осуществляется на следующем этапе.


```C#
    private void ValidateMetadataLocation()
    {
      if (!this.appContext.ContainsKey(AuthClaimTypes.MsExchAuthMetadataUrl))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.MsExchAuthMetadataUrl));
      }
    }

```


## Проверка подписи маркера удостоверения


После подтверждения наличия в маркере JWT утверждений, необходимых для проверки подписи, можно использовать Windows Identity Foundation (WIF) и расширения WIF для проверки подписи маркера. Для проверки подписи необходимы следующие сведения:


- Строка исходного удостоверения с кодированием URL-адреса base-64, отправленная с сервера Exchange.
    
- Расположение документа метаданных проверки подлинности из JWT.
    
- URL-адрес аудитории из JWT.
    
В этом примере конструктор объекта **IdentityToken** получает документ метаданных проверки подлинности с сервера Exchange и проверяет подпись маркера удостоверения. Если маркер удостоверения является допустимым, вы можете использовать экземпляр объекта **IdentityToken** для получения уникального идентификатора пользователя, включенного в маркер удостоверения.




```C#
    public IdentityToken(string rawToken, string audience, string authMetadataEndpoint)
    {
      X509Certificate2 currentCertificate = null;

      currentCertificate = AuthMetadata.GetSigningCertificate(new Uri(authMetadataEndpoint));

      JsonWebSecurityTokenHandler jsonTokenHandler =
          GetSecurityTokenHandler(audience, authMetadataEndpoint, currentCertificate);

      SecurityToken jsonToken = jsonTokenHandler.ReadToken(rawToken);
      JsonWebSecurityToken webToken = (JsonWebSecurityToken)jsonToken;

      SigningCertificateThumbprint = currentCertificate.Thumbprint;
      Issuer = webToken.Issuer;
      Audience = webToken.Audience;
      ValidTo = webToken.ValidTo;
      ValidFrom = webToken.ValidFrom;
      foreach (JsonWebTokenClaim claim in webToken.Claims)
      {
        if (claim.ClaimType.Equals(AuthClaimTypes.AppContextSender))
        {
          ApplicationContextSender = claim.Value;
        }

        if (claim.ClaimType.Equals(AuthClaimTypes.IsBrowserHostedApp))
        {
          IsBrowserHostedApp = claim.Value == "true";
        }

        if (claim.ClaimType.Equals(AuthClaimTypes.AppContext))
        {
          string[] appContextClaims = claim.Value.Split(',');
          Dictionary<string, string> appContext =
              new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(claim.Value);
          AuthenticationMetaDataUrl = appContext[AuthClaimTypes.MsExchAuthMetadataUrl];
          ExchangeID = appContext[AuthClaimTypes.MsExchImmutableId];
          TokenVersion = appContext[AuthClaimTypes.MsExchTokenVersion];
        }
      }
    }


```

Большая часть кода в конструкторе объекта **IdentityToken** задает свойства экземпляра с использованием утверждений с сервера Exchange. Конструктор вызывает метод **GetSecurityTokenHandler** для получения обработчика маркера, который проверит маркер удостоверения Exchange. Метод **GetSecurityTokenHandler** вызывает два служебных метода **GetMetadataDocument** и **GetSigningCertificate**, которые получают сертификат подписи с сервера Exchange. Каждый из этих методов описан в приведенных ниже разделах.


### Метод GetSecurityTokenHandler

Метод **GetSecurityTokenHandler** возвращает обработчик маркера WIF, который проверяет маркер удостоверения. Большая часть кода в этом методе инициализирует обработчик маркеров для выполнения проверки, однако этот метод вызывает метод **GetSigningCertificate** для получения сертификата X.509, который используется для подписывания маркера с сервера Exchange.


```C#
    private JsonWebSecurityTokenHandler GetSecurityTokenHandler(string audience,
        string authMetadataEndpoint,
        X509Certificate2 currentCertificate)
    {
      JsonWebSecurityTokenHandler jsonTokenHandler = new JsonWebSecurityTokenHandler();
      jsonTokenHandler.Configuration = new SecurityTokenHandlerConfiguration();

      jsonTokenHandler.Configuration.AudienceRestriction = new AudienceRestriction(AudienceUriMode.Always);
      jsonTokenHandler.Configuration.AudienceRestriction.AllowedAudienceUris.Add(
        new Uri(audience, UriKind.RelativeOrAbsolute));

      jsonTokenHandler.Configuration.CertificateValidator = X509CertificateValidator.None;

      jsonTokenHandler.Configuration.IssuerTokenResolver =
        SecurityTokenResolver.CreateDefaultSecurityTokenResolver(
          new ReadOnlyCollection<SecurityToken>(new List<SecurityToken>(
            new SecurityToken[]
            {
              new X509SecurityToken(currentCertificate)
            })), false);

      ConfigurationBasedIssuerNameRegistry issuerNameRegistry = new ConfigurationBasedIssuerNameRegistry();
      issuerNameRegistry.AddTrustedIssuer(currentCertificate.Thumbprint, Config.ExchangeApplicationIdentifier);
      jsonTokenHandler.Configuration.IssuerNameRegistry = issuerNameRegistry;

      return jsonTokenHandler;
    }
```


### Метод GetSigningCertificate

Метод **GetSigningCertificate** вызывает метод **GetMetadataDocument** для получения метаданных проверки подлинности с сервера Exchange, а затем возвращает первый сертификат X.509 из документа метаданных проверки подлинности. Если документ не существует, метод создает исключение приложения.


```C#
    private X509Certificate2 GetSigningCertificate(Uri authMetadataEndpoint)
    {
      JsonAuthMetadataDocument document = GetMetadataDocument(authMetadataEndpoint);

      if (null != document.keys &amp;&amp; document.keys.Length > 0)
      {
        JsonKey signingKey = document.keys[0];

        if (null != signingKey &amp;&amp; null != signingKey.keyValue)
        {
          return new X509Certificate2(Encoding.UTF8.GetBytes(signingKey.keyValue.value));
        }
      }

      throw new ApplicationException("The metadata document does not contain a signing certificate.");
    }

```


### Метод GetMetadataDocument

Документ метаданных проверки подлинности содержит сведения, которые понадобятся вам для проверки подписи маркера удостоверения Exchange. Документ передается в виде строки JSON. Метод **GetMetatDataDocument** запрашивает документ из указанного местоположения в маркере удостоверения Exchange и возвращает объект, который инкапсулирует строку JSON в качестве объекта. Если URL-адрес не содержит документ метаданных проверки подлинности, этот метод создает исключение приложения.


```C#
    private JsonAuthMetadataDocument GetMetadataDocument(Uri authMetadataEndpoint)
    {
      // Uncomment the next line if your Exchange server uses the default
      // self-signed certificate.
      // ServicePointManager.ServerCertificateValidationCallback = Config.CertificateValidationCallback;

      byte[] acsMetadata;
      using (WebClient webClient = new WebClient())
      {
        acsMetadata = webClient.DownloadData(authMetadataEndpoint);
      }
      string jsonResponseString = Encoding.UTF8.GetString(acsMetadata);

      JsonAuthMetadataDocument document = new JavaScriptSerializer().Deserialize<JsonAuthMetadataDocument>(jsonResponseString);

      if (null == document)
      {
        throw new ApplicationException(String.Format("No authentication metadata document found at {0}.", authMetadataEndpoint));
      }

      return document;
    }
```

По умолчанию сервер Exchange использует самозаверяющий сертификат X.509 для проверки подлинности запросов документа метаданных проверки подлинности. Если не установить сертификат, который может выполнять обратную трассировку до корневого сервера, необходимо создать метод обратного вызова проверки сертификата или запрос документа метаданных проверки подлинности; в противном случае запрос документа метаданных проверки подлинности завершится со сбоем. 

Класс **ServicePointManager** в пространстве имен System.Net платформы .NET Framework позволяет подключать метод обратного вызова проверки путем задания свойства **ServerCertificateValidationCallback**. Пример метода обратного вызова проверки сертификата, который подходит для разработки и тестирования, см. в статье [Проверка сертификатов X509](http://msdn.microsoft.com/en-us/library/dd633677%28EXCHG.80%29.aspx).


 **Примечание по безопасности.** Если вы используете метод обратного вызова проверки сертификата, то необходимо, чтобы он соответствовал требованиям к безопасности, принятым в вашей организации.


## Вычисление уникального идентификатора для учетной записи Exchange


Вы можете создать уникальный идентификатор для учетной записи Exchange путем хэширования URL-адреса документа метаданных проверки подлинности с идентификатором Exchange для учетной записи. Когда вы получите этот уникальный идентификатор, вы сможете использовать его для создания системы единого входа, предназначенной для веб-службы вашей надстройки Outlook. Дополнительные сведения об использовании уникальных идентификаторов для единого входа см. в статье [Проверка подлинности пользователя с помощью маркера удостоверения для Exchange](../outlook/authenticate-a-user-with-an-identity-token.md).

Свойство **UniqueUserIdentification** создает "соленый" хэш SHA256 идентификатора Exchange и URL-адреса метаданных проверки подлинности с помощью стандартного поставщика SHA256 из пространства имен **System.Security.Cryptography**.


 **Примечание по безопасности.** Чтобы создать уникальный идентификатор для учетной записи, необходимо хэшировать документ метаданных проверки подлинности с идентификатором Exchange. Если вы будете использовать только идентификатор Exchange, то к вашей службе смогут получить доступ неавторизованные пользователи. Как всегда при работе с функциями проверки подлинности и обеспечения безопасности, необходимо убедиться, что использование уникального идентификатора, созданного описанным методом, соответствует требованиям к безопасности, используемым в вашем приложении.




```C#
    // Salt to apply when creating unique ID.
    private byte[] Salt = new byte[] {<Provide random salt bytes here };

    private string ComputeUniqueIdentification()
    {
      byte[] inputBytes = Encoding.ASCII.GetBytes(string.Concat(ExchangeID, AuthenticationMetaDataUrl));

      // Combine input bytes and salt.
      byte[] saltedInput = new byte[Salt.Length + inputBytes.Length];
      Salt.CopyTo(saltedInput, 0);
      inputBytes.CopyTo(saltedInput, Salt.Length);

      // Compute the unique key.
      byte[] hashedBytes = SHA256CryptoServiceProvider.Create().ComputeHash(saltedInput);

      // Convert the hashed value to a string and return.
      return BitConverter.ToString(hashedBytes);
    }

    public string UniqueUserIdentification
    {
      get { return ComputeUniqueIdentification(); }
    }


```


## Объекты приложений


Примеры кода в этой статье зависят от нескольких служебных объектов, которые предоставляют понятные имена используемым ими константам. В следующей таблице приведены эти служебные объекты.


**Таблица 1. Служебные объекты**


|**Object**|**Описание**|
|:-----|:-----|
|**AuthClaimsType**|Централизованно собирает идентификаторы удостоверений, которые используются кодом проверки маркеров.|
|**Config**|Предоставляет константы для проверки маркера удостоверения. |
|**JsonAuthMetadataDocument**|Включает документ метаданных проверки подлинности JSON, отправленный с сервера Exchange.|

### Объект AuthClaimTypes

Объект **AuthClaimTypes** централизованно собирает идентификаторы утверждений, которые используются кодом проверки маркеров. Он включает как стандартные утверждения JWT, так и определенные утверждения в маркере удостоверения Exchange.


```C#
  public class AuthClaimTypes
  {
    public const string NameIdentifier =
        JsonWebTokenConstants.ReservedClaims.NameIdentifier;
    public const string MsExchImmutableId = "msexchuid";
    public const string MsExchTokenVersion = "version";
    public const string MsExchAuthMetadataUrl = "amurl";

    public const string AppContext =
        JsonWebTokenConstants.ReservedClaims.AppContext;
    public const string Audience =
        JsonWebTokenConstants.ReservedClaims.Audience;
    public const string Issuer =
        JsonWebTokenConstants.ReservedClaims.Issuer;
    public const string ValidFrom =
        JsonWebTokenConstants.ReservedClaims.NotBefore;
    public const string ValidTo =
        JsonWebTokenConstants.ReservedClaims.ExpiresOn;

    public const string AppContextSender = "appctxsender";
    public const string IsBrowserHostedApp = "isbrowserhostedapp";

    public const string TokenType = "typ";
    public const string Algorithm = "alg";
    public const string x509Thumbprint = "x5t";      
  }
```


### Объект Config

Объект **Config** содержит константы, которые используются для проверки маркера удостоверения, а также метод обратного вызова проверки сертификата, который вы можете использовать, если на сервере нет сертификата X509, трассируемого обратно до корневого сертификата.


 
  **Примечание по безопасности.** Метод обратного вызова сертификата безопасности требуется только в том случае, если сервер применяет самозаверяющий сертификат, используемый по умолчанию. В этом примере метод обратного вызова возвращает значение **false**, если сертификат является самозаверяющим, поэтому следует заменить его на метод обратного вызова, который соответствует требованиям к безопасности, принятым в вашей организации. Пример метода обратного вызова проверки сертификата, который подходит для разработки и тестирования, см. в статье [Проверка сертификатов X509](http://msdn.microsoft.com/en-us/library/dd633677%28EXCHG.80%29.aspx).


```C#
  public static class Config
  {
    public static string Algorithm = "RS256";
    public static string Audience = @"https:\\localhost:44300\Pages\IdentityTest.html";
    public static string TokenType = "JWT";
    public static string Version = "ExIdTok.V1";

    public static string ExchangeApplicationIdentifier = "Exchange";

    internal static bool CertificateValidationCallback(
    object sender,
    System.Security.Cryptography.X509Certificates.X509Certificate certificate,
    System.Security.Cryptography.X509Certificates.X509Chain chain,
    System.Net.Security.SslPolicyErrors sslPolicyErrors)
    {
      // If the certificate is a valid, signed certificate, return true.
      if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
      {
        return true;
      }

      // If there are errors in the certificate chain, look at each error to determine the cause.
      else
      {
        return false;
      }
    }
  }
```


### Объект JsonAuthMetadataDocument

Объект **JsonAuthMetadataDocument** предоставляет содержимое документа метаданных проверки подлинности через свои свойства.


```C#
using System;

namespace IdentityTest
{
  public class JsonAuthMetadataDocument
  {
    public string id { get; set; }
    public string version { get; set; }
    public string name { get; set; }
    public string realm { get; set; }
    public string serviceName { get; set; }
    public string issuer { get; set; }
    public string [] allowedAudiences { get; set; }
    public JsonKey[] keys;
    public JsonEndpoint[] endpoints;
  }

  public class JsonEndpoint
  {
    public string location { get; set; }
    public string protocol { get; set; }
    public string usage { get; set; }
  }

  public class JsonKey
  {
    public string usage { get; set; }
    public JsonKeyValue keyValue { get; set; }
  }

  public class JsonKeyValue
  {
    public string type { get; set; }
    public string value { get; set; }
  }
}

```


## Дополнительные ресурсы



- [Проверка подлинности надстройки Outlook с помощью маркеров удостоверения Exchange](../outlook/authentication.md)
    
- [Подробные сведения о маркере удостоверения Exchange](../outlook/inside-the-identity-token.md)
    
