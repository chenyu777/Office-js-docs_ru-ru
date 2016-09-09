
# Метод Binding.setDataAsync
Записывает данные в привязанный раздел документа, представленный указанным объектом привязки.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Word|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, TableBindings, TextBindings|
|**Последнее изменение в TableBindings**|1.1|

```js
bindingObj.setDataAsync(data [, options] ,callback);
```


## Параметры



|**Имя**|**Тип**|**Описание**|**Примечания по вопросам поддержки**|
|:-----|:-----|:-----|:-----|
| _data_|<table><tr><td><b>string</b></td><td>только Excel, Excel Online, Word и Word Online</td></tr><tr><td><b>array</b> (массив массивов — "матрица")</td><td>только Excel и Word</td></tr><tr><td><a href="https://msdn.microsoft.com/en-us/library/office/fp161002"><b>TableData</b></a></td><td>только Access, Excel и Word</td></tr><tr><td><b>HTML</b></td><td>только Word и Word Online</td></tr><tr><td><b>Office Open XML</b></td><td>только Word</td></tr></table>|Данные, записываемые в текущий фрагмент. Обязательный.|**Изменено в версии:** 1.1. Для поддержки контентных надстроек для Access требуется набор требований **TableBinding** версии 1.1 или более поздней.|
| _options_|**object**|Задает следующие [необязательные параметры](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods).||
| _coercionType_|**[CoercionType](../../reference/shared/coerciontype-enumeration.md)**|Указывает способ приведения задаваемых данных. ||
| _столбцы_|**массив строк**| Задает имена столбцов.|**Добавлено в версии:** 1.1. Только для привязок таблиц в контентных надстройках для Access.|
| _rows_|**Office.TableRange.ThisRow**|Указывает предопределенную строку "thisRow" для задания данных в текущей выбранной строке. Необязательный параметр. |**Добавлено в версии:** 1.1. Только для привязок таблиц в контентных надстройках для Access.|
| _startColumn_|**number**|Задает начальный столбец подмножества данных с отсчетом от нуля. |Только для привязки таблицы или матрицы. Если опущен, данные записываются начиная с первого столбца.|
| _startRow_|**number**|Указывает начальную строку (с отсчетом от нуля) для подмножества данных в привязке. |Только для привязки таблицы или матрицы. Если опущен, данные записываются начиная с первой строки.|
| _tableOptions_|**object**|Для вставленной таблицы это список пар "ключ-значение", которые задают [параметры форматирования таблицы](../../docs/excel/format-tables-in-add-ins-for-excel.md), такие как строку заголовков, строку итогов и полосы строк. |**Добавлено в версии:** 1.1. **Поддерживается в:** Excel.|
| _cellFormat_|**object**|Для вставленной таблицы это список пар "ключ-значение", которые указывают диапазон столбцов, строк, ячеек и [их форматирования](../../docs/excel/format-tables-in-add-ins-for-excel.md).|**Добавлено в версии** 1.1. **Поддерживается в:** Excel, Excel Online.|
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** или **undefined**|Определяемый пользователем элемент любого типа, который возвращается в объекте **AsyncResult** без изменения.||
| _callback_|**object**|Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.||

## Значение обратного вызова

Когда выполняется функция, переданная в параметр _callback_, она получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ исключительно с помощью параметра функции обратного вызова.

В функции обратного вызова, переданной в метод **setDataAsync**, можно использовать свойства объекта **AsyncResult**, чтобы возвратить такие сведения, как следующие:



|**Свойство**|**Применение**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Всегда возвращает значение **undefined**, так как объекты и данные не извлекаются.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Определяет, удалось ли выполнить операцию.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Получает доступ к объекту [Error](../../reference/shared/error.md), который содержит сведения об ошибке, если операция завершилась неудачно.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Получает доступ к определенному пользователем объекту **object** или значению, если они передаются как параметр _asyncContext_.|

## Заметки

Значение, передаваемое для параметра _data_, содержит данные, записываемые в привязку. Тип передаваемых значений определяет, что будет записано, как описано в следующей таблице.



|**Значение _data_**|**Записываемые данные**|
|:-----|:-----|
|Значение типа **string**|Записывается обычный текст или другие данные, которые могут быть приведены к типу **string**.|
|Массив массивов (матрица)|Будут вставлены табличные данные без заголовков. Например, для записи данных в три строки по два столбца можно передать массив, подобный следующему: ` [["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]`. Для записи одного столбца из трех строк можно передать массив, подобный следующему: `[["R1C1"], ["R2C1"], ["R3C1"]]`|
|Объект [TableData](../../reference/shared/tabledata.md)|Записываются табличные данные с заголовками.|
Кроме того, при записи данных в привязку применяются следующие действия, соответствующие конкретному приложению.

 **В Word** указанное значение параметра _data_ записывается в привязку следующим образом:



|**Значение _data_**|**Записываемые данные**|
|:-----|:-----|
|Значение типа **string**|Записывается указанный текст.|
|Массив массивов (матрица) или объект **TableData**|Записывается таблица Word.|
|HTML|Записывается указанный HTML-код.
 >**Важно!** Если указанный HTML-код содержит недопустимые фрагменты, Word не вызовет ошибку. Word запишет весь допустимый HTML-код и пропустит недопустимые данные.

|
|Office Open XML (Open XML)|Записывается указанный XML-код.|  **В Excel** указанное значение параметра _data_ записывается в привязку следующим образом:



|**Значение _data_**|**Записываемые данные**|
|:-----|:-----|
|Значение типа **string**|Указанный текст вставляется в качестве значения первой привязанной ячейки. Вы также можете указать допустимую формулу, чтобы добавить ее в привязанную ячейку. Например, если для _data_ указать `"=SUM(A1:A5)"`, будут получены итоговые значения в указанном диапазоне. Но после того как вы задали формулу для привязанной ячейки, вы не сможете прочитать добавленную формулу (или уже существующие формулы) из привязанной ячейки. Когда вызывается метод [Binding.getDataAsync](../../reference/shared/binding.getdataasync.md) для привязанной ячейки для чтения этих данных, метод может вернуть только данные, отображенные в ячейке (результат формулы).|
|Массив массивов (матрица) и форма точно соответствует форме указанной привязки|Записывается заданный набор строк и столбцов. Вы также можете указать массив массивов, содержащих допустимые формулы, чтобы добавить их в привязанные ячейки. Например, если для _data_ указать `[["=SUM(A1:A5)","=AVERAGE(A1:A5)"]]`, две эти формулы будут добавлены в привязку, содержащую две ячейки. Точно так же, как при указании формулы для одной привязанной ячейки, вы не можете читать добавленные формулы (или уже существующие формулы) из связки с помощью метода **Binding.getDataAsync**, поскольку он возвращает только данные, отображенные в привязанных ячейках.|
|Объект **TableData** и форма таблицы соответствуют форме привязанной таблицы.|Записывается заданный набор строк и/или заголовков, если при этом не будут перезаписаны другие данные в соседних ячейках. **Примечание.** Если вы укажете формулы в объекте **TableData**, который вы передаете для параметра _data_, можно не получить ожидаемые результаты, потому что функция "вычисляемые столбцы" Excel автоматически копирует формулы в столбце. Чтобы обойти это, когда вы хотите записать _data_ с формулами в привязанную таблицу, попробуйте указать данные как массив массивов (вместо объекта **TableData**) и для _coercionType_ указать **Microsoft.Office.Matrix** или "matrix".|
 **Дополнительные заметки для приложения Excel Online**


- Общее количество ячеек в значении, переданном параметру _data_ в одном вызове этого метода, не может превышать 20 000.
    
- Количество _групп форматирования_, переданных параметру _cellFormat_, не может превышать 100. Одна группа форматирования состоит из набора форматов, примененного к указанному диапазону ячеек. Например, приведенный ниже вызов передает параметру _cellFormat_ две группы форматирования.
    
```js
  Office.select("bindings#myBinding").setDataAsync([['Berlin'],['Munich'],['Duisburg']],
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});

```

Во всех остальных случаях возвращается ошибка.

Метод **setDataAsync** запишет данные в подмножество привязки таблицы или матрицы, если указаны необязательные параметры _startRow_ и _startColumn_, которые задают допустимый диапазон.


## Пример




```js
function setBindingData() {
    Office.select("bindings#MyBinding").setDataAsync('Hello World!', function (asyncResult) { });
}
```

Указав необязательный параметр _coercionType_, вы можете задать тип данных, которые требуется записать в привязку. Например, если в текстовую привязку в Word записывается HTML-код, можно указать параметр _coercionType_ со значением `"html"`, как показано в приведенном ниже примере, где используются теги HTML `<b>`, выделяющие слово "Hello" полужирным шрифтом.




```js
function writeHtmlData() {
    Office.select("bindings#myBinding").setDataAsync("<b>Hello</b> World!", {coercionType: "html"}, function (asyncResult) {
        if (asyncResult.status == "failed") {
            write('Error: ' + asyncResult.error.message);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

В этом примере при вызове **setDataAsync** передается параметр _data_ в виде массива массивов (для создания одного столбца из трех строк) и с помощью параметра _coercionType_ указывается структура данных `"matrix"`.




```js
function writeBoundDataMatrix() {
    Office.select("bindings#myBinding").setDataAsync([['Berlin'],['Munich'],['Duisburg']],{ coercionType: "matrix" }, function (asyncResult) {
        if (asyncResult.status == "failed") {
            write('Error: ' + asyncResult.error.message);
        } else {
            write('Bound data: ' + asyncResult.value);
        }
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

В этом примере в функции `writeBoundDataTable` при вызове **setDataAsync** передается параметр _data_ в виде объекта **TableData** (для записи трех столбцов и трех строк) и с помощью параметра _coercionType_ указывается структура данных `"table"`. 

В функции `updateTableData` при вызове **setDataAsync** снова передается параметр _data_ как объект **TableData**, но в виде одного столбца с новым заголовком и трех строк, чтобы обновить значения в последнем столбце таблицы, созданной с помощью функции `writeBoundDataTable`. Необязательный параметр _startColumn_, отсчитывающийся с нуля, равен 2, чтобы заменить значения в третьем столбце таблицы.




```js
function writeBoundDataTable() {
    // Create a TableData object.
    var myTable = new Office.TableData();
    myTable.headers = ['First Name', 'Last Name', 'Grade'];
    myTable.rows = [['Kim', 'Abercrombie', 'A'], ['Junmin','Hao', 'C'],['Toni','Poe','B']];

    // Set myTable in the binding.
    Office.select("bindings#myBinding").setDataAsync(myTable, { coercionType: "table" }, 
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write('Error: '+ asyncResult.error.message);
        } else {
            write('Bound data: ' + asyncResult.value);
        }
    });
}

// Replace last column with different data.
function updateTableData() {
     var newTable = new Office.TableData();
     newTable.headers = ["Gender"];
     newTable.rows = [["M"],["M"],["F"]];
     Office.select("bindings#myBinding").setDataAsync(newTable, { coercionType: "table", startColumn:2 }, 
         function (asyncResult) {
             if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                 write('Error: '+ asyncResult.error.message);
         } else {
            write('Bound data: ' + asyncResult.value);
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
|**Access**||Y||
|**Excel**|Y|Да|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|MatrixBindings, TableBindings, TextBindings|
|**Минимальный уровень разрешений**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel и Word в Office для iPad.|
|1.1|<ul><li>В надстройки для Access добавлена поддержка записи таблиц данных.</li><li>В надстройки для Excel добавлена поддержка <a href="http://msdn.microsoft.com/library/46b05707-b350-41be-b6b8-311799c71a33(Office.15).aspx" target="_blank">настройки форматирования при записи данных в привязку таблицы</a> с помощью дополнительных параметров <span class="parameter" sdata="paramReference">tableOptions</span> и <span class="parameter" sdata="paramReference">cellFormat</span>.</li></ul>|
|1.0|Представлено|

## См. также



#### Другие ресурсы


[Привязка к областям в документе или электронной таблице](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)
