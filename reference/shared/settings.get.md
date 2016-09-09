
# Метод Settings.get
Извлекает указанный параметр.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, PowerPoint, Word|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Параметры|
|**Последнее изменение в **|1.1|

```js
var mySetting = Office.context.document.settings.get(name);
```


## Параметры



_name_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Тип: **string**

&nbsp;&nbsp;&nbsp;&nbsp;Имя извлекаемого параметра с учетом регистра.

    



## Возвращаемое значение

Объект **object**, имена свойств которого сопоставлены сериализованным значениям JSON.


## Пример




```js
function displayMySetting() {
    write('Current value for mySetting: ' + Office.context.document.settings.get('mySetting'));
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что это свойство поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает это свойство.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).



||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Да|Y|
|**PowerPoint**|Y|Да|Y|
|**Word**|Y|Да|Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|Параметры|
|**Минимальный уровень разрешений**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки




|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка PowerPoint Online.|
|1.1|Добавлена поддержка Excel, PowerPoint и Word в Office для iPad.|
|1.1|Добавлена поддержка создания параметров в контентных надстройках для Access.|
|1.0|Представлено|
