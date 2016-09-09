
# Метод CustomXmlNode.setTextAsync
Асинхронно задает текст узла XML в настраиваемой XML-части.

|||
|:-----|:-----|
|**Ведущие приложения:**|Word|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Добавлен в версии**|1.2|

```
customXmlNodeObj.setTextAsync(text, [asyncContext,]callback(asyncResult);
```


## Параметры



|**Имя**|**Тип**|**Описание**|
|:-----|:-----|:-----|
| _text_|**string**|Обязательный параметр. Текстовое значение XML-узла.|
| _asyncContext_|**object**|Необязательный. Пользовательский объект, доступный в свойстве asyncContext объекта [AsyncResult](../../reference/shared/asyncresult.md). С его помощью можно указать объект или значение **AsyncResult**, если функция обратного вызова является именованной.|
| _callback_|**object**|Необязательный. Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.|

## Значение обратного вызова

Когда выполняется функция, переданная в параметр _callback_, она получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ исключительно с помощью параметра функции обратного вызова.

Если функция обратного вызова передана методу **setTextAsync**, можно использовать свойства объекта **AsyncResult** для возврата следующей информации.



|**Свойство**|**Применение**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Не используется.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Указывает, удалось ли выполнить операцию.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Получает доступ к объекту [Error](../../reference/shared/error.md), который содержит сведения об ошибке, если операция завершилась неудачно.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Получает доступ к определенному пользователем **object** или значению, если они передаются как параметр _asyncContext_. Если параметр _asyncContext_ не задан, это свойство возвращает undefined.|

## Пример

Узнайте, как задать текстовое значение узла в настраиваемой XML-части.


```js
// Get the built-in core properties XML part by using its ID. This results in a call to Word.
Office.context.document.customXmlParts.getByIdAsync("{6C3C8BC8-F283-45AE-878A-BAB7291924A1}", function (getByIdAsyncResult) {
    
    // Access the XML part.
    var xmlPart = getByIdAsyncResult.value;
    
    // Add namespaces to the namespace manager. These two calls result in two calls to Word.
    xmlPart.namespaceManager.addNamespaceAsync('cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', function () {
        xmlPart.namespaceManager.addNamespaceAsync('dc', 'http://purl.org/dc/elements/1.1/', function () {

            // Get XML nodes by using an Xpath expression. This results in a call to the host.
            xmlPart.getNodesAsync("/cp:coreProperties/dc:subject", function (getNodesAsyncResult) {
                
                // Get the first node returned by using the Xpath expression. This will be the subject element in this example.
                var subjectNode = getNodesAsyncResult.value[0];
                
                // Set the text value of the subject node and use the asyncContext. This results in a call to the host. 
                // The results are logged to the browser console. 
                subjectNode.setTextAsync("newSubject", {asyncContext: "StateNormal"}, function (setTextAsyncResult) {
                   console.log("The status of the call: " + setTextAsyncResult.status);
                   console.log("The asyncContext value = " + setTextAsyncResult.asyncContext);
                });
            });
        });
    });
});
```


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y|Да|Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|CustomXmlParts|
|**Минимальный уровень разрешений**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Область задач|
|**Библиотека**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



****


|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлен метод setTextAsync.|
