
# Проверка подлинности пользователя с помощью маркера удостоверения для Exchange

Вы можете реализовать схему проверки подлинности единого входа для информационной службы, которая позволяет клиентам, использующим надстройки Outlook, подключаться к вашей службе с помощью своих учетных данных на сервере Exchange Server. В этой статье показано, как можно сопоставить учетные данные с помощью простого хранилища данных пользователя на основе объектов **Dictionary**.

 >**Примечание.** Это простой пример единого входа, и ее не следует использовать в производственном коде. Как всегда, когда вы имеете дело с удостоверениями и проверкой подлинности, необходимо убедиться, что код соответствует требованиям безопасности вашей организации.


## Необходимые условия для использования проверки подлинности единого входа


Чтобы использовать маркер удостоверения для единого входа, приложению-службе требуется допустимый маркер удостоверения. Сведения о маркерах удостоверений, а также о порядке запроса и проверки маркера удостоверения см. в следующих статьях.


- [Подробные сведения о маркере удостоверения Exchange](../outlook/inside-the-identity-token.md)
    
- [Вызов службы из надстройки Outlook с использованием маркера удостоверения в Exchange](../outlook/call-a-service-by-using-an-identity-token.md)
    
- [Использование библиотеки проверки маркеров Exchange](../outlook/use-the-token-validation-library.md), если используется управляемый код, или [Проверка маркера удостоверения Exchange](../outlook/validate-an-identity-token.md), если создается собственный метод проверки маркеров.
    

## Проверка подлинности пользователя


В следующем примере кода показан простой объект проверки подлинности, который сопоставляет уникальное удостоверение, представленное маркером удостоверения, с набором учетных данных для службы. Класс **TokenAuthentication** предоставляет метод **GetResponseFromService**, который будет возвращать ответ для ранее проверенных маркеров или запрашивать у пользователя учетные данные, которые можно проверить на подлинность и сопоставить с маркером удостоверения. Этот код не завершен. Предполагается, что вы предоставите следующие объекты и методы.



|**Объект или метод**|**Описание**|
|:-----|:-----|
|Объект **LocalCredentials**|Представляет учетные данные пользователя для службы. Структура объекта зависит от требований вашей службы.|
|Объект **IdentityToken**|Содержит маркер удостоверения пользователя, отправленный в службу надстройкой Outlook. Этот объект должен содержать как минимум уникальный идентификатор пользователя Exchange и URL-адрес метаданных аутентификации для сервера, выдавшего этот маркер. В данном примере используется объект маркера удостоверения, определенный в статье [Проверка маркера удостоверения Exchange](../outlook/validate-an-identity-token.md).|
|Объект **JsonResponse**|Представляет ответ от службы. Этот объект может быть сериализован в объект JSON.|
|Метод **CallService**|Вызывает службу с объектом **LocalCredentials**, содержащим учетные данные пользователя для службы, и с объектом, содержащим данные для запроса службы. Если учетные данные допустимы, то этот метод возвращает объект **JsonReponse**, содержащий результаты запроса. Если учетные данные недопустимы, то метод возвращает значение **null**.|
|Метод **GetCredentialsResponse**|Возвращает объект **JsonReponse**, который ваша почтовая надстройка для Office будет распознавать как запрос учетных данных для службы.|
|Метод **LocalCredentialsAreValid**|Возвращает значение **true**, если предоставленные для службы учетные данные допустимы; в противном случае возвращает значение **false**.|

 >**Примечание.** Ниже приводится только один вариант использования маркера удостоверения. Как обычно, когда вы имеете дело с удостоверениями и проверкой подлинности, необходимо убедиться, что код соответствует требованиям безопасности вашей организации.


```C#
    public class TokenAuthentication
    {
        // This example uses a Dictionary object to store local credentials. Your application should use
        // a data store that is appropriate to the security requirements of your organization.
        private Dictionary<string, LocalCredentials> AuthenticationCache = new Dictionary<string, LocalCredentials>();

        // Salt to apply when creating unique ID.
        private byte[] Salt = new byte[] {25, 139, 201, 13};

        private JsonResponse CallService(LocalCredentials credentials, object data)
        {
            // Calls the local service to get the response for the user.
            return null;
        }

        private JsonResponse GetCredentialsResponse()
        {
            // Creates a response that tells the Outlook add-in to
            // request the user's credentials for the service.
            return null;
        }

        private bool LocalCredentialsAreValid(LocalCredentials credentials)
        {
            // Returns true if the service recognizes the credentials provided.
            return false;
        }

        private string ComputeSHA256Hash(string uniqueId, string authenticationMetadataUrl, byte[] salt)
        {
            byte[] inputBytes = Encoding.ASCII.GetBytes(string.Concat(uniqueId, authenticationMetadataUrl));

            // Combine input bytes and salt.
            byte[] saltedInput = new byte[salt.Length + inputBytes.Length];
            salt.CopyTo(saltedInput, 0);
            inputBytes.CopyTo(saltedInput, salt.Length);

            // Compute the unique key.
            byte[] hashedBytes = SHA256CryptoServiceProvider.Create().ComputeHash(saltedInput);

            // Convert the hashed value to a string and return.
            return BitConverter.ToString(hashedBytes);
        }

        public JsonResponse GetResponseFromService(IdentityToken token, LocalCredentials credentials, object data)
        {
            JsonResponse response = null;
            // This method should never be called with a null token.
            if (null == token)
            {
                throw new ArgumentNullException("token");
            }

            if (null == credentials)
            {
                string uniqueKey = ComputeSHA256Hash(token.ExchangeID, token.AuthenticationMetadataUrl, Salt);
                if (!AuthenticationCache.ContainsKey(uniqueKey))
                {
                    // The user's credentials are not in the authentication cache. Ask
                    // for the credentials.
                    response = GetCredentialsResponse();
                }
                else
                {
                    // The user's credentials are in the cache; make a request.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials. For example,
                        // the user has ended their subscription to the service, or the
                        // credentials have expired. Get new credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // The service returned a response to the user. Return the
                        // service response.
                        response = serviceResponse;
                    }
                }
            }
            else
            {
                // If the credentials are not null, it's a request to add an identity
                // to the authentication cache. Check to determine whether the local credentials
                // sent to the service are known.
                if (LocalCredentialsAreValid(credentials))
                {
                    // The local credentials are known. Add them to the 
                    // cached credentials.
                    string uniqueKey = ComputeSHA256Hash(token.ExchangeID, token.AuthenticationMetadataUrl, Salt);
                    AuthenticationCache.Add(uniqueKey, credentials);

                    // Get a response from the service.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // Return the service response to the user.
                        response = serviceResponse;
                    }
                }
            }

            return response;
        }
    }}
```


## Проверка подлинности пользователя с помощью управляемой библиотеки проверки


При использовании управляемой библиотеки для проверки маркеров удостоверений нет необходимости вычислять уникальный ключ. В качестве уникального ключа для пользователя можно непосредственно использовать свойство **UniqueUserIdentification** класса **AppIdentityToken**. В следующем примере кода показаны изменения метода **GetResponseFromService** из предыдущего примера, которые необходимы для использования класса **AppIdentityToken**.


```js
        public JsonResponse GetResponseFromService(AppIdentityToken token, LocalCredentials credentials, object data)
        {
            JsonResponse response = null;
            // This method should never be called with a null token.
            if (null == token)
            {
                throw new ArgumentNullException("token");
            }

            if (null == credentials)
            {
                string uniqueKey = token.UniqueUserIdentitification;
                if (!AuthenticationCache.ContainsKey(uniqueKey))
                {
                    // The user's credentials are not in the authentication cache. Ask
                    // for the credentials.
                    response = GetCredentialsResponse();
                }
                else
                {
                    // User's credentials are in the cache. Make a request.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials. For example,
                        // the user has ended their subscription to the service, or the
                        // credentials have expired. Get new credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // The service returned a response to the user. Return the
                        // service response.
                        response = serviceResponse;
                    }
                }
            }
            else
            {
                // If the credentials are not null, it's a request to add an identity
                // to the authentication cache. Check to determine whether the local credentials
                // sent to the service are known.
                if (LocalCredentialsAreValid(credentials))
                {
                    // The local credentials are known. Add them to the 
                    // cached credentials. 
                    string uniqueKey = token.UniqueUserIdentitification;
                    AuthenticationCache.Add(uniqueKey, credentials);

                    // Get a response from the service.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // Return the service response to the user.
                        response = serviceResponse;
                    }
                }
            }

            return response;
        }
```


## Дополнительные ресурсы



- [Проверка подлинности надстройки Outlook с помощью маркеров удостоверения Exchange](../outlook/authentication.md)
    
- [Вызов службы из надстройки Outlook с использованием маркера удостоверения в Exchange](../outlook/call-a-service-by-using-an-identity-token.md)
    
- [Использование библиотеки проверки маркеров Exchange](../outlook/use-the-token-validation-library.md)
    
- [Проверка маркера удостоверения Exchange](../outlook/validate-an-identity-token.md)
    
