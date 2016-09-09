
# Вызов службы из надстройки Outlook с использованием маркера удостоверения в Exchange

Маркер удостоверения предоставляет уникальный идентификатор для каждого из клиентов и может использоваться для персонализации службы, которую вы предоставляете. Ваш код может запрашивать маркер удостоверения на сервере Exchange Server с помощью вызова асинхронного метода, возвращающего строку в надстройку Outlook. Строка содержит маркер удостоверения JSON Web Token (JWT). Надстройке не требуется распаковывать маркер. Вместо этого она передает маркер в веб-службу, чтобы веб-служба могла выполнять проверку подлинности при получении запроса от надстройки.

Веб-служба, которая поддерживает вашу надстройку, должна работать на том же сервере, на котором размещаются исходные HTML- и JavaScript-файлы надстройки. Это предотвращает появление ошибок при выполнении межсайтовых сценариев. Ваш сервер может передать запрос другим веб-службам, если это требуется для надстройки.

Добавить маркер удостоверения в запрос службы, отправляемый надстройкой, очень просто — достаточно запросить маркер, применить его и получить ответ веб-службы. Вот как это выглядит в простом XML-документе, отправляемом на сервер с помощью метода **XmlHttpRequest**.

## Запрос токена на сервере Exchange


В этом простом способе инициализации надстройки с помощью метода **getUserIdentityTokenAsync** запрашивается маркер удостоверения на сервере Exchange Server. Параметр _getUserIdentityToken_ — это функция, которая вызывается при возврате асинхронного запроса, отправленного серверу. Описание метода обратного вызова представлено в следующем разделе.


```js
var _mailbox;
var _xhr;
// The initialize function is required for all add-ins.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
        _mailbox = Office.context.mailbox;
    _mailbox.getUserIdentityTokenAsync(getUserIdentityTokenCallback);
    });
}

```


## Использование токенов удостоверений


Функция обратного вызова для метода **getUserIdentityTokenAsync** имеет один параметр, содержащий токен удостоверения пользователя в своем свойстве **value**.

Эта функция обратного вызова создает объект **XMLHttpRequest** для вызова веб-службы. В качестве свойства **onreadystatechange** объекта **XMLHttpRequest** укажите имя функции, которая должна запускаться, когда надстройка получит ответ от веб-службы.




```js
function getUserIdentityTokenCallback(asyncResult) {
    var token = asyncResult.value;

    _xhr = new XMLHttpRequest();
    _xhr.open("POST", "https://localhost:44300/IdentityTestService/UnpackTokenJSON");
    _xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    _xhr.onreadystatechange = readyStateChange;

    var request = new Object();
    request.token = token;
    request.phoneNumbers = _mailbox.item.getEntities().phoneNumbers;

    _xhr.send(JSON.stringify(request));
}
```


## Использование ответа веб-службы


Это другая простая функция, которая обрабатывает ответ от веб-службы. Она соответствует стандартному шаблону для функций обратного вызова **XHMHttpResponse**. Она ожидает прихода от веб-службы полного ответа, а затем передает содержимое ответа пользовательскому интерфейсу надстройки. Ответ, который анализирует эта функция — это ответ от веб-службы. Сведения об ответе см. в разделе [Проверка маркера удостоверения Exchange](../outlook/validate-an-identity-token.md). 


```js
function readyStateChange() {
    if (_xhr.readyState == 4 &amp;&amp; _xhr.status == 200) {

        var response = JSON.parse(_xhr.responseText);

        if (undefined == response.error) {
            document.getElementById("msexchuid").value = response.token.msexchuid;
            document.getElementById("amurl").value = response.token.amurl;
            document.getElementById("uniqueID").value = response.token.uniqueID;
            document.getElementById("iss").value = response.token.iss;
            document.getElementById("x5t").value = response.token.x5t;
            document.getElementById("nbf").value = response.token.nbf;
            document.getElementById("exp").value = response.token.exp;
        }
        else {
            document.getElementById("error").value = response.error;
        }
    }
}
```


## Пример вызова веб-службы с помощью токена удостоверения


Токены удостоверений передают сведения удостоверений клиентов, вызывающих службу, веб-службе, работающей на сервере. Чтобы использовать токены удостоверений, необходимо следующее:


- Надстройка Outlook, которая запрашивает маркер удостоверения на сервере Exchange Server и отправляет его в веб-службу. Сведения, приведенные в данной статье, помогут создать такую надстройку.
    
- Веб-служба, которая проверят маркер удостоверения и работает на сервере, предоставляющем пользовательский интерфейс для вашей надстройки. Сведения, необходимые для создания веб-службы, представлены в следующих статьях:
    
      - [Использование библиотеки проверки маркеров Exchange](../outlook/use-the-token-validation-library.md) -- (если используется предоставленная нами библиотека проверки).
    
  - [Проверка маркера удостоверения Exchange](../outlook/validate-an-identity-token.md) (если вы пишете свой код проверки).
    

### Пример кода надстройки


Следующие файлы необходимы для надстройки, описанной в этой статье:


- IdentityTest.js — JavaScript-файлы, которые предоставляют бизнес-логику для надстройки.
    
- IdentityTest.html — HTML-файл, который предоставляет пользовательский интерфейс для надстройки.
    
Вам также потребуется веб-служба для тестирования удостоверения. Сведения об этой веб-службе см. в разделе [Проверка маркера удостоверения Exchange](../outlook/validate-an-identity-token.md).


#### IdentityTest.js

Следующий пример показывает файл IdentityTest.js.


```js
var _mailbox;
var _xhr;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    _mailbox = Office.context.mailbox;
    _mailbox.getUserIdentityTokenAsync(getUserIdentityTokenCallback);
    });
}
function getUserIdentityTokenCallback(asyncResult) {
    var token = asyncResult.value;

    _xhr = new XMLHttpRequest();
    _xhr.open("POST", "https://localhost:44300/IdentityTestService/UnpackTokenJSON");
    _xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    _xhr.onreadystatechange = readyStateChange;

    var request = new Object();
    request.token = token;
    request.phoneNumbers = _mailbox.item.getEntities().phoneNumbers;

    _xhr.send(JSON.stringify(request));
}

function readyStateChange() {
    if (_xhr.readyState == 4 &amp;&amp; _xhr.status == 200) {

        var response = JSON.parse(_xhr.responseText);

        if (undefined == response.error) {
            document.getElementById("msexchuid").value = response.token.msexchuid;
            document.getElementById("amurl").value = response.token.amurl;
            document.getElementById("uniqueID").value = response.token.uniqueID;
            document.getElementById("iss").value = response.token.iss;
            document.getElementById("x5t").value = response.token.x5t;
            document.getElementById("nbf").value = response.token.nbf;
            document.getElementById("exp").value = response.token.exp;
        }
        else {
            document.getElementById("error").value = response.error;
        }
    }
}
```


#### IdentityTest.html

Следующий пример показывает файл IdentityTest.html.


```HTML
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>Identity Test</title>

    <link rel="stylesheet" type="text/css" href="../Content/Office.css" />
    <link rel="stylesheet" type="text/css" href="../Content/App.css" />

    <script src="../Scripts/jquery-1.6.2.js"></script>
    <script src="../Scripts/Office/MicrosoftAjax.js"></script>
    <script src="../Scripts/Office/Office.js"></script>

    <!-- Add your JavaScript to the following JavaScript file -->
    <script src="../Scripts/IdentityTest.js"></script>
</head>
<body>
    <div id="SectionContent">
        <table style="width: 80%;">
            <tr>
                <th>Claim
                </th>
                <th>Contents
                </th>
            </tr>
            <tr>
                <td style="width: 25%;">Error:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="error" value="None" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">User Exchange ID:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="msexchuid" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Authentication Metadata URL:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="amurl" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Unique identifier:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="uniqueID" />
                </td>
            </tr>
          </tr>
            <tr>
                <td style="width: 25%;">Audience:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="aud" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Issuer:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="iss" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Certificate thumbprint:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="x5t" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Valid from:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="nbf" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Valid to:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="exp" />
                </td>
            </tr>
        </table>
    </div>
</body>
</html>
```


## Дальнейшие действия


Теперь, когда вы знаете, как запрашивать маркер идентификации, вам необходимо использовать этот маркер на серверной стороне запроса. Следующие статьи помогут начать работу:


- [Использование библиотеки проверки маркеров Exchange](../outlook/use-the-token-validation-library.md)
    
- [Проверка маркера удостоверения Exchange](../outlook/validate-an-identity-token.md)
    
- [Проверка подлинности пользователя с помощью маркера удостоверения для Exchange](../outlook/authenticate-a-user-with-an-identity-token.md)
    

## Дополнительные ресурсы



- [Проверка подлинности надстройки Outlook с помощью маркеров удостоверения Exchange](../outlook/authentication.md)
    
- [Подробные сведения о маркере удостоверения Exchange](../outlook/inside-the-identity-token.md)
    
