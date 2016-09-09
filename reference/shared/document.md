
# Объект Document
Абстрактный класс, представляющий документ, с которым взаимодействует надстройка.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, PowerPoint, Project, Word|
|**Добавлен в версии**|1.0|
|**Последнее изменение в **|1.1|

```
Office.context.document
```


## Элементы


**Свойства**


|**Имя**|**Описание**|**Примечания по вопросам поддержки**|
|:-----|:-----|:-----|
|[bindings](../../reference/shared/document.bindings.md)|Получает объект, предоставляющий доступ к привязкам, определенным в документе.|В версии 1.1 добавлена поддержка контентных надстроек для Access.|
|[customXmlParts](../../reference/shared/document.customxmlparts.md)|Получает объект, представляющий настраиваемые XML-части в документе.||
|[mode](../../reference/shared/document.mode.md)|Получает режим, в котором находится документ.|В версии 1.1 добавлена поддержка контентных надстроек для Access.|
|[параметры](../../reference/shared/document.settings.md)|Получает объект, который представляет сохраненные настраиваемые параметры надстройки области задач или контентной надстройки для текущего документа.|В версии 1.1 добавлена поддержка контентных надстроек для Access.|
|[url](../../reference/shared/document.url.md)|Получает URL-адрес документа, открытого ведущим приложением.|В версии 1.1 добавлена поддержка контентных надстроек для Access.|

**Методы**


|**Имя**|**Описание**|**Примечания по вопросам поддержки**|
|:-----|:-----|:-----|
|[addHandlerAsync](../../reference/shared/document.addhandlerasync.md)|Добавляет обработчик для события объекта **Document**.||
|[getActiveViewAsync](../../reference/shared/document.getactiveviewasync.md)|Возвращает текущее представление презентации.|В версии 1.1 добавлена поддержка [надстроек PowerPoint](../../docs/powerpoint/powerpoint-add-ins.md).|
|[getFileAsync](../../reference/shared/document.getfileasync.md)|Возвращает полный файл документа фрагментами размером до 4194304 байт (4 МБ).|В версии 1.1 добавлена поддержка считывания файлов в формате PDF в надстройках для PowerPoint и Word.|
|[getFilePropertiesAsync](../../reference/shared/document.getfilepropertiesasync.md)|Получает свойства текущего документа. В этом выпуске можно считать только URL-адрес документа.|В версии 1.1 добавлена возможность получить URL-адрес документа в надстройках для Excel, Word и PowerPoint.|
|[getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md)|Читает данные, содержащиеся в выбранном разделе документа.|В версии 1.1 добавлена поддержка считывания идентификатора, заголовка и индекса выбранного диапазона слайдов в надстройках PowerPoint.|
|[goToByIdAsync](../../reference/shared/document.gotobyidasync.md)|Переходит к указанному объекту или месту в документе.|В версии 1.1 добавлена поддержка навигации по документу в надстройках для Excel и PowerPoint.|
|[removeHandlerAsync](../../reference/shared/document.removehandlerasync.md)|Удаляет обработчик события объекта **Document**.||
|[setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md)|Записывает данные в текущий фрагмент в документе.|В версии 1.1 добавлена поддержка [настройка форматирования выбранной таблицы при записи данных в надстройках Excel](../../docs/excel/format-tables-in-add-ins-for-excel.md).|

**События**


|**Имя**|**Описание**|**Примечания по вопросам поддержки**||
|:-----|:-----|:-----|:-----|
|[ActiveViewChanged](../../reference/shared/document.activeviewchanged.md)|Возникает, когда пользователь изменяет текущее представление документа.|В версии 1.1 добавлена поддержка надстроек PowerPoint.||
|[SelectionChanged](../../reference/shared/document.selectionchanged.event.md)|Происходит при изменении выбора в документе.|||

## Заметки

В скрипте не нужно непосредственно создавать экземпляр объекта **Document**. Чтобы вызвать элементы объекта **Document** для взаимодействия с текущим документом или листом, используйте объект `Office.context.document`.


## Пример

В следующем примере используется метод **getSelectedDataAsync** объекта **Document** для получения выбранного фрагмента в виде текста и его отображения на странице надстройки.


```js

// Display the user's current selection.
function showSelection() {
    Office.context.document.getSelectedDataAsync(
        "text",                        // coercionType
        {valueFormat: "unformatted",   // valueFormat
        filterType: "all"},            // filterType
        function (result) {            // callback
            var dataValue; 
            dataValue = result.value;
            write('Selected data is: ' + dataValue);
        });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Сведения о поддержке


Поддержка каждого элемента API объекта **Document** зависит от ведущего приложения Office. Информацию о поддержке элемента в том или ином приложении см. в соответствующем разделе "Сведения о поддержке".

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


|||
|:-----|:-----|
|**Добавлен в версии**|1.0|
|**Последнее изменение в **|1.1|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|
