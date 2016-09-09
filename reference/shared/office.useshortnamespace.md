

# Метод Office.useShortNamespace
Включает или отключает псевдоним `Office` для всего пространства имен `Microsoft.Office.WebExtension`.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Последнее изменение в **|1.1|

```js
Office.useShortNamespace(useShortcut);
```


## Параметры



_useShortcut_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Тип: **boolean**

    
&nbsp;&nbsp;&nbsp;&nbsp;**true**, чтобы использовать псевдоним; **false** чтобы отключить его. Значение по умолчанию: **true**.
    


## Пример



```js
function startUsingShortNamespace() {
    if (typeof Office === 'undefined') {
        Microsoft.Office.WebExtension.useShortNamespace(true);
    }
    else {
        Office.useShortNamespace(true);
    }
    write('Office alias is now ' + typeof Office);
}

function stopUsingShortNamespace() {
    if (typeof Office === 'undefined') {
        Microsoft.Office.WebExtension.useShortNamespace(false);
    }
    else {
        Office.useShortNamespace(false);
    }
    write('Office alias is now ' + typeof Office);
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


|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка PowerPoint Online.|
|1.1|Добавлена поддержка Excel, PowerPoint и Word в Office для iPad.|
|1.1|Добавлена поддержка для вызова этого метода в контентных надстройках для Access.|
|1.0|Представлено|
