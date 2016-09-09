
# Метод Document.getFilePropertiesAsync
Получает свойства текущего документа.

|||
|:-----|:-----|
|**Ведущие приложения:**|Excel, PowerPoint, Word|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Не в наборе|
|**Добавлено в версии**|1.1|

```js
Office.context.document.getFilePropertiesAsync([, options], callback);
```


## Параметры



|**Имя**|**Тип**|**Описание**|**Примечания по вопросам поддержки**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|Задает следующие [необязательные параметры](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** или **undefined**|Определяемый пользователем элемент любого типа, который возвращается в объекте **AsyncResult** без изменения.||
| _callback_|**object**|Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.||

## Значение обратного вызова

Когда выполняется функция, переданная в параметр _callback_, она получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью единственного параметра функции обратного вызова.

В функции обратного вызова, переданной методу **getFilePropertiesAsync**, вы можете использовать свойства объекта **AsyncResult**, чтобы получить следующие сведения:



|**Свойство**|**Применение**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Возвращает URL-адрес файла в свойство `asyncResult.value.url`|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Определяет, удалось ли выполнить операцию.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Получает доступ к объекту [Error](../../reference/shared/error.md), который содержит сведения об ошибке, если операция завершилась неудачно.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Получает доступ к определенному пользователем объекту **object** или значению, если они передаются как параметр _asyncContext_.|

## Заметки

В этом выпуске API URL-адрес файла можно получить только с помощью свойства **url** (`asyncResult.value.url`).


## Пример

Для чтения URL-адреса текущего файла необходимо написать функцию обратного вызова, которая возвращает URL-адрес. В следующем примере показано, как:


 - **передать функцию обратного вызова**, которая возвращает значение URL-адреса файла в параметр _callback_ метода **getFilePropertiesAsync**;
    
 - **показать значение** на странице надстройки.
    

```js
function getFileUrl() {
    //Get the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;
        if (fileUrl == "") {
            showMessage("The file hasn't been saved yet. Save the file and try again");
        }
        else {
            showMessage(fileUrl);
        }
    });
}
```




## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Поддерживаемые ведущие приложения по платформе**


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Да|Y|
|**PowerPoint**|Y|Да|Y|
|**Word**|Y|Да|Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|Не в наборе|
|**Минимальный уровень разрешений**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки




|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel, PowerPoint и Word в Office для iPad.|
|1.1|Представлен.|
