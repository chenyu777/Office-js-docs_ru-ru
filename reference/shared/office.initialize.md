
# Событие Office.initialize
Происходит, когда среда выполнения загружена и надстройка готова начать взаимодействие с приложением и размещенным документом. 

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Последнее изменение в **|1.1|

```js
Office.initialize = function (reason) {/* initialization code */}
```


## Заметки

Параметр _reason_ функции прослушивателя событий **initialize** возвращает значение перечисления [InitializationReason](../../reference/shared/initializationreason-enumeration.md), которое указывает, как произошла инициализация. Надстройку области задач или контентную надстройку можно инициализировать двумя способами:


- пользователь может вставить надстройку из раздела **Недавно использовавшиеся надстройки** раскрывающегося списка **Надстройка** на вкладке **Вставка** на ленте в ведущем приложении Office или из диалогового окна **Вставка надстройки**;
    
- пользователь может открыть документ, который уже содержит надстройку.
    

 >**Параметр**. Параметр reason функции прослушивателя событий **initialize** возвращает значение перечисления **InitializationReason** только для надстроек области задач и контентных надстроек. Он не возвращает значение для надстроек Outlook.


## Пример

Значение **InitializationEnumeration** можно использовать для реализации логики, которая отличает надстройку, добавленную впервые, от надстройки, которая уже была частью документа. В следующем примере показана простая логика, которая использует значение параметра _reason_ для отображения способа инициализации надстройки области задач или контентной надстройки.


```js
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, code specific to the add-in can run.
    // Display initialization reason.
    if (reason == "inserted")
    write("The add-in was just inserted.");

    if (reason == "documentOpened")
    write("The add-in is already part of the document.");
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает, что данное событие поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает это событие.

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
|1.1|Добавлена возможность инициализации контентных надстроек в Access.|
|1.0|Представлено|
