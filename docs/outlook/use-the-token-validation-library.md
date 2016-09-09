
# Использование библиотеки проверки маркеров управляемого API веб-служб Exchange

Можно идентифицировать клиентов в надстройке Outlook с помощью маркера удостоверения, который надстройка запрашивает с сервера, где работает Exchange Server 2013 или Exchange Online. Маркер в формате JSON Web Token — уникальный идентификатор для учетной записи электронной почты на сервере Exchange Server. Управляемый API веб-служб Exchange (EWS) предоставляет вспомогательные классы для упрощения использования маркера удостоверения.

## Предварительные требования для использования библиотеки проверки

Чтобы проверить маркер удостоверения Exchange, необходимо установить [библиотеку управляемого API EWS](https://www.nuget.org/packages/Microsoft.Exchange.WebServices).

## Проверка маркера удостоверения Exchange

Библиотека проверки управляемого API EWS предоставляет класс **AppIdentityToken** для управления маркерами удостоверения Exchange. Следующий метод показывает, как создать экземпляр **AppIdentityToken** и вызвать метод **Validate** для проверки допустимости маркера. Этот метод принимает следующие параметры:

- *rawToken*. Строковое представление маркера, возвращаемое в надстройке Outlook из метода [**Office.context.mailbox.getUserIdentityTokenAsync**](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox).
- *hostUri*. Полный универсальный код ресурса (URI) к странице в надстройке Outlook, которая вызвала метод **getUserIdentityTokenAsync**.

```C#
// Required to use the validation library.
using Microsoft.Exchange.WebServices.Auth.Validate;

private AppIdentityToken CreateAndValidateIdentityToken(string rawToken, string hostUri)
{
    try
    {
        AppIdentityToken token = (AppIdentityToken)AuthToken.Parse(rawToken);
        token.Validate(new Uri(hostUri));

        return token;
    }
    catch (TokenValidationException ex)
    {
        throw new ApplicationException("A client identity token validation error occurred.", ex);
    }
}
```

## Дополнительные ресурсы

- [Проверка подлинности надстройки Outlook с помощью маркеров удостоверения Exchange](../outlook/authentication.md)  
- [Подробные сведения о маркере удостоверения Exchange](../outlook/inside-the-identity-token.md)
- [Проверка маркера удостоверения Exchange](../outlook/validate-an-identity-token.md)
    
