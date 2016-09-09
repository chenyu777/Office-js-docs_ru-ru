
# Метод Binding.getDataAsync
Возвращает данные, содержащиеся в привязке.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Word|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, TableBindings, TextBindings|
|**Последнее изменение в TableBindings**|1.1|

```
bindingObj.getDataAsync([, options] , callback );
```


## Параметры



|**Имя**|**Тип**|**Описание**|**Примечания по вопросам поддержки**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|Задает следующие [необязательные параметры](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods).||
| _coercionType_|**[CoercionType](../../reference/shared/coerciontype-enumeration.md)**|Указывает способ приведения задаваемых данных. ||
| _valueFormat_|[ValueFormat](../../reference/shared/valueformat-enumeration.md)|Определяет, применяется ли форматирование к возвращаемым значениям, таким как числа и даты.||
| _filterType_|[FilterType](../../reference/shared/filtertype-enumeration.md)|Указывает, должен ли применяться фильтр к полученным данным.||
| _rows_|**Office.TableRange.ThisRow**| Указывает предопределенную строку "thisRow" для получения данных в текущей выбранной строке. Необязательный параметр.|Только для привязок таблиц в контентных надстройках для Access.|
| _startRow_|**number**|Для привязки таблицы или матрицы задает начальную строку (с отсчетом от нуля) для подмножества данных в привязке. ||
| _startColumn_|**number**|Для привязки таблицы или матрицы задает начальный столбец (с отсчетом от нуля) для подмножества данных в привязке. ||
| _rowCount_|**number**|Для табличной или матричной привязки задает количество строк смещения от _startRow_. ||
| _columnCount_|**number**|Для привязки таблицы или матрицы задает число столбцов смещения от _startColumn_.||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** или **undefined**|Определяемый пользователем элемент любого типа, который возвращается в объекте **AsyncResult** без изменения.||
| _callback_|**object**|Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.||

## Значение обратного вызова

Когда выполняется функция, переданная в параметр _callback_, она получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ исключительно с помощью параметра функции обратного вызова.

Если функция обратного вызова передана методу **Binding.getDataAsync**, можно использовать свойства объекта **AsyncResult** для возврата следующей информации.



|**Свойство**|**Применение**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Получает доступ к значениям в указанной привязке. Если указан параметр _coercionType_ (и вызов выполнен успешно), данные возвращаются в формате, описанном в разделе о перечислении [CoercionType](../../reference/shared/coerciontype-enumeration.md).|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Определяет, удалось ли выполнить операцию.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Получает доступ к объекту [Error](../../reference/shared/error.md), который содержит сведения об ошибке, если операция завершилась неудачно.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Получает доступ к определенному пользователем объекту **object** или значению, если они передаются как параметр _asyncContext_.|

## Заметки

Если необязательный параметр опущен, используется следующее значение по умолчанию (если применимо к типу и формату данных).



|**Параметр**|**По умолчанию**|
|:-----|:-----|
| _coercionType_|Исходный, или неприведенный, тип привязки.|
| _valueFormat_|Неформатированные данные.|
| _filterType_|Все значения (без фильтрации).|
| _startRow_|Первая строка.|
| _startColumn_|Первый столбец.|
| _rowCount_|Все строки.|
| _columnCount_|Все столбцы.|
При вызове из [MatrixBinding](../../reference/shared/binding.matrixbinding.md) или [TableBinding](../../reference/shared/binding.tablebinding.md) метод **getDataAsync** вернет подмножество значений привязки, если указаны необязательные параметры _startRow_, _startColumn_, _rowCount_ и _columnCount_ (и эти параметры задают непрерывный и допустимый диапазон).


## Пример




```
function showBindingData() {
    Office.select("bindings#MyBinding").getDataAsync(function (asyncResult) {
        write(asyncResult.value)
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



Существует существенное отличие в режимах между использованием `"table"` и `"matrix"`_coercionType_ с методом **Binding.getDataAsync**, в отношении форматирования данных в строках заголовка, как указано в двух следующих примерах. Указанные примеры кода отображают функции обработчика событий для события [Binding.SelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md).

Если задать `"table"` _coercionType_, свойство [TableData.rows](../../reference/shared/tabledata.rows.md) (`result.value.rows` в следующем примере кода) возвращает массив, содержащий только строки таблицы. Таким образом, нулевая строка является первой строкой таблицы, не являющейся заголовком.




```js
function selectionChanged(evtArgs) { 
    Office.select("bindings#TableTranslate").getDataAsync({ coercionType: 'table', startRow: evtArgs.startRow, startCol: 0, rowCount: 1, columnCount: 1 },  
        function (result) { 
            if (result.status == 'succeeded') { 
                write("Image to find: " + result.value.rows[0][0]); 
            } 
            else 
                write(result.error.message); 
    }); 
}     
// Function that writes to a div with id='message' on the page. 
function write(message){ 
    document.getElementById('message').innerText += message; 
}
```

Но если указать `"matrix"` _coercionType_, свойство `result.value` в следующем примере кода возвращает массив, содержащий заголовок таблицы в нулевой строке. Если заголовок таблицы содержит множество строк, все они включаются в матрицу `result.value` в качестве отдельных строк перед включением строк таблицы.




```js
function selectionChanged(evtArgs) { 
    Office.select("bindings#TableTranslate").getDataAsync({ coercionType: 'matrix', startRow: evtArgs.startRow, startCol: 0, rowCount: 1, columnCount: 1 },  
        function (result) { 
            if (result.status == 'succeeded') { 
                write("Image to find: " + result.value[1][0]); 
            } 
            else 
                write(result.error.message); 
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


**Поддерживаемые ведущие приложения по платформе**


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Да|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|MatrixBindings, TableBindings, TextBindings|
|**Минимальный уровень разрешений**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



****


|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel и Word в Office для iPad.|
|1.1|Добавлена поддержка табличной привязки в надстройках для Access.|
|1.0|Представлено|

## См. также



#### Другие ресурсы


[Привязка к областям в документе или электронной таблице](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)
