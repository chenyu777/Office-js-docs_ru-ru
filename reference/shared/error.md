
# Объект Error
Предоставляет конкретные сведения об ошибке, возникшей во время выполнения асинхронной операции с данными.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Последнее изменение в **|1.1|

```
asyncResult.error
```


## Элементы


**Свойства**


|**Имя**|**Описание**|
|:-----|:-----|
|[code](../../reference/shared/error.code.md)|Получает цифровой код ошибки.|
|[name](../../reference/shared/error.name.md)|Получает имя ошибки.|
|[сообщение](../../reference/shared/error.message.md)|Получает подробное описание ошибки.|

## Заметки

К объекту **Error** можно обратиться из объекта [AsyncResult](../../reference/shared/asyncresult.md), который возвращается функцией, переданной в качестве аргумента _callback_ асинхронной операции с данными, такой как метод [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) объекта **Document**.


## Пример

В следующем примере используется метод **setSelectedDataAsync**, чтобы вставить строку "Hello World!" в выбранный фрагмент текста. Если метод завершается ошибкой, отображаются значения свойств **name** и **message** объекта **Error**.


```js
function setText() {

    Office.context.document.setSelectedDataAsync("Hello World!", {},
        function (asyncResult) {
            if (asyncResult.status === "failed")
            var err = asyncResult.error; 
                write(err.name + ": " + err.message);
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).

||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|**OWA для устройств**|**Outlook для Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Да|Y|||
|**Outlook**|Y|Да||Да|Y|
|**PowerPoint**|Y|Да|Y|||
|**Project**|Y|||||
|**Word**|Y|Да|Y|||

|||
|:-----|:-----|
|**Минимальный уровень разрешений**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Надстройки области задач, надстройки Outlook, контентные надстройки|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



****


|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel, PowerPoint и Word в Office для iPad.|
|1.1|Добавлена поддержка контентных надстроек для Access.|
|1.0|Представлено|
