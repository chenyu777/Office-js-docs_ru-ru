
# Метод Document.getActiveViewAsync
 Возвращает состояние текущего представления презентации (редактирование или чтение).

|||
|:-----|:-----|
|**Ведущие приложения:** Excel, PowerPoint, Word|**Типы надстроек:** надстройки области задач и контентные надстройки.|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|ActiveView|
|**Добавлен в ActiveView**|1.1|

```
Office.context.document.getActiveViewAsync([,options], callback);
```


## Параметры



|**Имя**|**Тип**|**Описание**|**Примечания по вопросам поддержки**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|Задает следующие [необязательные параметры](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** или **undefined**|Определяемый пользователем элемент любого типа, который возвращается в объекте **AsyncResult** без изменения.||
| _callback_|**object**|Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.||

## Значение обратного вызова

Когда выполняется функция, переданная в параметр _callback_, она получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью единственного параметра функции обратного вызова.

В функции обратного вызова, переданной методу **getActiveViewAsync**, свойство [AsyncResult.value](../../reference/shared/asyncresult.value.md) возвращает состояние текущего представления презентации. Может возвращено значение `edit` или `read`. `edit` соответствует всем представлениям, в которых можно редактировать слайды, таким как **обычный режим** или **режим структуры**. `read` соответствует **режиму показа слайдов** или **режиму чтения**.


## Заметки

Может активировать событие при изменении представления.


## Пример

Для определения представления текущей презентации необходимо написать функцию обратного вызова, которая возвращает значение. В следующем примере показано, как:


-  **передать функцию обратного вызова**, которая возвращает тип представления в параметр _callback_ метода **getActiveViewAsync**;
    
-  **показать значение** на странице надстройки.
    

```js
function getFileView() {
    // Get whether the current view is edit or read.
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage(asyncResult.value);
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
|**Excel**|||Y|
|**PowerPoint**|Y|Да|Y|
|**Word**|||Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|ActiveView|
|**Добавлен в ActiveView**|1.1|
|**Минимальный уровень разрешений**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки





****


|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel, PowerPoint и Word в Office для iPad.|
|1.1|Представлен.|
