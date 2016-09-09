
# Свойство BindingSelectionChangedEventArgs.startColumn
Получает индекс первого столбца текущего выбора (с отсчетом от нуля).

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Word|
|**Последнее изменение в **|1.1|

```
var startCol = eventArgsObj.startColumn;
```


## Возвращаемое значение

Отсчитываемый от нуля индекс первого столбца текущего выбора, начиная с самого левого столбца в привязке.


## Заметки

Если пользователь выбирает не смежные столбцы, то возвращаются координаты последнего сплошного выбора. 

В Word это свойство будет работать только для привязок с [BindingType](../../reference/shared/bindingtype-enumeration.md) "table". Если привязка имеет тип "matrix", возвращается значение **null**. Кроме того, вызов завершится неудачно, если в таблице имеются объединенные ячейки, поскольку для правильной работы этого свойства структура таблицы должна быть однородной.


## Пример

В следующем примере добавляется обработчик событий для события [SelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md) в привязку с [id](../../reference/shared/binding.id.md)`myTable`. Когда пользователь изменяет выбор, обработчик отображает координаты первой ячейки в выборе, а также количество выбранных строк и столбцов.


```js
function addSelectionHandler() {
    Office.context.document.bindings.getByIdAsync("myTable", function (result) {
        result.value.addHandlerAsync("bindingSelectionChanged", myHandler);
    });
}

// Display selection start coordinates and row/column count.
function myHandler(bArgs) {
    write("Selection start row/col: " + bArgs.startRow + "," + bArgs.startColumn);
    write("Selection row count: " + bArgs.rowCount);
    write("Selection col count: " + bArgs.columnCount);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что это свойство поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает это свойство.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Поддерживаемые ведущие приложения по платформе**


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Да|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Минимальный уровень разрешений**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



****


|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel и Word в Office для iPad.|
|1.1|Добавлена поддержка надстроек для Access.|
|1.0|Представлено|
