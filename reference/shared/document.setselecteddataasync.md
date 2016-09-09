
# Метод Document.setSelectedDataAsync
Записывает данные в текущий фрагмент в документе.

|||
|:-----|:-----|
|**Ведущие приложения:** Access, Excel, PowerPoint, Project, Word, Word Online|**Типы надстроек: ** контентные, надстройки области задач|
|**Доступно в [наборе требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Версия последнего изменения**|1.1|

```js
Office.context.document.setSelectedDataAsync(data [, options], callback(asyncResult));
```


## Параметры

|**Имя**|**Тип**|**Описание**|**Примечания по вопросам поддержки**|
|:-----|:-----|:-----|:-----|
| _data_|Данные могут быть любого из следующих типов:<ul><li><b>Строка.</b> (Office.CoercionType.Text.) Применяется только в Excel, Excel Online, PowerPoint, PowerPoint Online, Word и Word Online.</li><li><b>Массив массивов.</b> (Office.CoercionType.Matrix.) Применяется только в Excel, Word и Word Online.</li><li>[TableData.](../../reference/shared/tabledata.md) (Office.CoercionType.Table.) Применяется только в Access, Excel, Word и Word Online.</li><li><b>HTML.</b> (Office.CoercionType.Html.) Применяется только в Word и Word Online.</li><li><b>Office Open XML.</b> (Office.CoercionType.Ooxml.) Применяется только в Word и Word Online.</li><li><b>Поток образа с кодировкой Base64.</b> (Office.CoercionType.Image.) Применяется только в Excel, PowerPoint, Word и Word Online.</li></ul>|Данные, записываемые в текущий выделенный фрагмент. Обязательный.|**Версия изменения:** 1.1. Для поддержки контентных надстроек для Access необходим набор требований **Selection** версии 1.1 или более поздней. Чтобы можно было задать данные образа, необходим набор требований **ImageCoercion** версии 1.1 или более поздней. Чтобы задать его для активации приложения, используйте такой код:<br/><br/>`<Requirements>`<br/>&nbsp;&nbsp;`<Sets DefaultMinVersion="1.1">`<br/>&nbsp;&nbsp;&nbsp;&nbsp;`<Set Name="ImageCoercion"/>`<br/>&nbsp;&nbsp;`</Sets>`<br/>`</Requirements>`<br/><br/>Для обнаружения возможности использовать ImageCoercion в среде выполнения подойдет такой код:<br/><br/>`if (Office.context.requirements.isSetSupported('ImageCoercion', '1.1')) {)) {`<br/>&nbsp;&nbsp;&nbsp;&nbsp;`// insertViaImageCoercion();`<br/>`} else {`<br/>&nbsp;&nbsp;&nbsp;&nbsp;`// insertViaOoxml();`<br/>`}`|
| _options_|**object**|Задает набор [необязательных параметров](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods). Объект options может содержать следующие свойства для задания параметров:<br/><ul><li>coercionType (<b><a href="735eaab6-5e31-4bc2-add5-9d378900a31b.htm">CoercionType</a></b>). Указывает способ приведения задаваемых данных. Если параметр coercionType не задан, используется стандартное значение Office.CoercionType.Text.</li><li>tableOptions (<b>object</b>). Для вставленной таблицы это список пар "ключ-значение", которые задают <a href="http://msdn.microsoft.com/library/46b05707-b350-41be-b6b8-311799c71a33(Office.15).aspx" target="_blank">параметры форматирования таблицы</a> (например, для отображения строки заголовков, строки итогов и чередующихся строк). </li><li>cellFormat (<b>object</b>). Для вставленной таблицы это список пар "ключ-значение", которые определяют диапазон столбцов, строк, ячеек, а также <a href="http://msdn.microsoft.com/library/46b05707-b350-41be-b6b8-311799c71a33(Office.15).aspx" target="_blank">форматирование последних</a>. </li><li>imageLeft (<b>number</b>). Этот параметр подходит для вставки изображений. Обозначает место вставки относительно левого края слайда в PowerPoint или выделенной в данный момент ячейки в Excel. Это значение не учитывается в Word. Это значение представлено в точках.</li><li>imageTop (<b>number</b>). Этот параметр подходит для вставки изображений. Обозначает место вставки относительно верхнего края слайда в PowerPoint или выделенной в данный момент ячейки в Excel. Это значение не учитывается в Word. Это значение представлено в точках.</li><li>imageWidth (<b>number</b>). Этот параметр подходит для вставки изображений. Обозначает ширину изображения. Если этот параметр указывается без параметра imageHeight, изображение масштабируется в соответствии с указанной шириной. Если указаны и ширина, и высота, то размер изображения изменяется соответствующим образом. Если не указана ни высота, ни ширина, используются исходные размер и пропорции изображения. Это значение представлено в точках.</li><li>imageHeight (<b>number</b>). Этот параметр подходит для вставки изображений. Обозначает высоту изображения. Если этот параметр указывается без параметра imageWidth, изображение масштабируется в соответствии с указанной высотой. Если указаны и ширина, и высота, то размер изображения изменяется соответствующим образом. Если не указана ни высота, ни ширина, используются исходные размер и пропорции изображения. Это значение представлено в точках.</li><li>asyncContext (<b>object \| value</b>). Определяемый пользователем объект, доступный в свойстве asyncContext объекта <a href="540c114f-0398-425c-baf3-7363f2f6bc47.htm">AsyncResult</a>. С его помощью можно указать объект или значение <b>AsyncResult</b>, если параметр callback — именованная функция.</li></ul>|Параметры _tableOptions_ и _cellFormat_ появились в версии 1.1 и поддерживаются в Excel 2013 и Excel Online.<br/><br/>Параметры _imageLeft_ и _ImageTop_ поддерживаются в Excel и PowerPoint.|
| _callback_|**object**|Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.||

## Значение обратного вызова

Когда выполняется функция, переданная в параметр _callback_, она получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью единственного параметра функции обратного вызова.

Если функция обратного вызова передана методу **setSelectedDataAsync**, свойство [AsyncResult.value](../../reference/shared/asyncresult.value.md) всегда возвращает значение **undefined**, так как объект или данные, которые следует получить, отсутствуют.


## Замечания

Значение, передаваемое для параметра _data_, содержит данные для записи в текущий выделенный фрагмент. Отличия в использовании значений:


-  **Строка.** Записывается обычный текст или другие данные, которые могут быть приведены к типу **string**.
    
    
    
    В Excel можно также указать параметр _data_ в виде допустимой формулы, чтобы добавить ее в выделенную ячейку. Например, если задать для параметра _data_ значение `"=SUM(A1:A5)"`, значения в указанном диапазоне будут суммироваться. Тем не менее, если задать формулу в связанной ячейке, добавленную (или существующую) формулу будет невозможно считать. При вызове метода [Document.getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) в выделенной ячейке для считывания ее данных этот метод может возвращать только данные, отображаемые в ячейке (результат формулы).
    
-  **Массив массивов ("matrix").** Будут вставлены табличные данные без заголовков. Например, чтобы записать данные в три строки двух столбцов, вы можете передать массив следующим образом: `[["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]`. Чтобы записать один столбец из трех строк, передайте массив следующим образом: `[["R1C1"], ["R2C1"], ["R3C1"]]`.
    
    
    
    В Excel вы также можете указать параметр _data_ как массив массивов, содержащий допустимые формулы, чтобы добавить их в выделенные ячейки. Например, если никакие другие данные не будут перезаписаны, а параметру _data_ будет присвоено значение `[["=SUM(A1:A5)","=AVERAGE(A1:A5)"]]`, эти две формулы будут добавлены в выделенный фрагмент. Как и при указании формулы в одной ячейке в текстовом виде, добавленные (или существующие) формулы невозможно считывать после того, как они заданы. Вы можете считывать только результаты формул.
    
-  **Объект [TableData](../../reference/shared/tabledata.md)** Вставляются табличные данные с заголовками.
    
    
    
     **Примечание.** Если вы укажете в Excel формулы в объекте **TableData**, который вы передаете для параметра _data_, вы можете не получить ожидаемые результаты, потому что функция "вычисляемые столбцы" Excel автоматически копирует формулы в столбце. Чтобы обойти это, когда вы хотите записать _data_ с формулами в выбранную таблицу, попробуйте указать данные как массив массивов (вместо объекта **TableData**) и для _coercionType_ указать **Microsoft.Office.Matrix** или "матрица".
    
 **Поведение конкретного приложения**

Кроме того, при записи данных в выделенный фрагмент применяются следующие действия, соответствующие конкретному приложению.

 **Для Word**


- Если фрагмент не выбран и точка вставки находится в допустимом расположении, указанные данные _data_ вставляются в точку вставки, как описано ниже.
    
      - If  _data_ is a string, the specified text is inserted.
    
  - Если параметр _data_ относится к массиву массивов ("matrix") или объекту **TableData**, будет вставлена новая таблица Word.
    
  - Если параметр _data_ относится к коду HTML, то вставляется этот код.
    
     >**Важно!**  Если какой-либо вставляемый код HTML недопустим, Word не сообщит об ошибке. Word вставит как можно больше кода HTML, не включая недопустимые данные.
  - Если параметр _data_ относится к Office Open XML, вставляются указанные данные XML.
    
  - Если параметр _data_ относится к потоку образа с кодировкой Base64, вставляется указанный образ.
    
- Если выделен фрагмент, он будет заменен на указанное значение параметра _data_ в соответствии с правилами, описанными выше.
    
-  **Вставка изображений**: вставляемые изображения помещаются в тексте. Параметры **imageLeft** и **imageTop** игнорируются. Пропорции изображения всегда блокируются. Если задан только параметр **imageWidth** или **imageHeight**, второе значение будет выбрано автоматически с учетом пропорций.
    
 **Для Excel**


- Если выделена одна ячейка:
    
      - If  _data_ is a string, the specified text is inserted as the value of the current cell.
    
  - если параметр _data_ содержит массив массивов ("матрицу"), вставляется заданный набор строк и столбцов, если при этом не будут перезаписаны другие данные в соседних ячейках;
    
  - если параметр _data_ содержит объект **TableData**, вставляется новая таблица Excel с заданным набором строк и заголовков, если при этом не будут перезаписаны другие данные в соседних ячейках.
    
- Если выделено несколько ячеек и их форма не соответствует форме _data_, возвращается ошибка.
    
- Если выделено несколько ячеек и их форма точно соответствует форме _data_, то значения выделенных ячеек заменяются на значения параметра _data_.
    
-  **Вставляются перемещаемые** изображения. Параметры положения **imageLeft** и **imageTop** указываются относительно выделенных ячеек. Отрицательные значения **imageLeft** и **imageTop** допустимы и могут быть откорректированы Excel для помещения изображения в пределах листа. Пропорции изображения блокируются, если не указаны параметры **imageWidth** и **imageHeight**. Если задан только параметр **imageWidth** или **imageHeight**, второе значение будет подобрано автоматически с учетом исходных пропорций.
    
Во всех остальных случаях возвращается сообщение об ошибке.

 **Для Excel Online**

В дополнение к поведению, описанному для Excel выше, при записи данных в Excel Online применяются следующие ограничения: 


- Общее количество ячеек, записанных на лист с помощью параметра _data_ в одном вызове этого метода, не может превышать 20 000.
    
- Количество _групп форматирования_, переданных параметру _cellFormat_, не может превышать 100. Одна группа форматирования состоит из набора форматов, примененного к указанному диапазону ячеек. Например, приведенный ниже вызов передает параметру _cellFormat_ две группы форматирования.
    

```js
  Office.context.document.setSelectedDataAsync(
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});
```

 **Для PowerPoint**

Вставляются перемещаемые изображения. Указывать параметры положения **imageLeft** и **imageTop** необязательно, но если указан один, должен быть указан и второй. Если указано одно значение, оно игнорируется. Отрицательные значения **imageLeft** и **imageTop** допустимы и позволяют поместить изображение за пределами слайда. Если не указано ни одного необязательного параметра, а слайд содержит заполнитель, изображение заменит заполнитель в слайде. Пропорции изображения блокируются, если не указаны параметры **imageWidth** и **imageHeight**. Если задан только параметр **imageWidth** или **imageHeight**, второе значение будет подобрано автоматически с учетом исходных пропорций.


## Пример

В следующем примере в выделенный фрагмент или ячейку записывается текст "Hello World!". В случае ошибки отображается значение свойства [error.message](../../reference/shared/error.message.md).


```js
function writeText() {
    Office.context.document.setSelectedDataAsync("Hello World!",
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                 write(error.name + ": " + error.message);
            }
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



Указав необязательный параметр _coercionType_, вы можете задать тип данных, которые требуется записать в выделенный фрагмент. В следующем примере записываются данные в виде массива, содержащего три строки и два столбца. Параметру _coercionType_ для этой структуры данных присваивается значение `"matrix"`. В случае ошибки отображается значение свойства [error.message](../../reference/shared/error.message.md).




```js
function writeMatrix() {
    Office.context.document.setSelectedDataAsync([["Red", "Rojo"], ["Green", "Verde"], ["Blue", "Azul"]], {coercionType: Office.CoercionType.Matrix}
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                write(error.name + ": " + error.message);
            }
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



В следующем примере записываются данные в виде таблицы с одним столбцом, заголовком и четырьмя строками. Параметру _coercionType_ для этой структуры данных присваивается значение `"table"`. В случае ошибки отображается значение свойства [error.message](../../reference/shared/error.message.md).




```js
function writeTable() {
    // Build table.
    var myTable = new Office.TableData();
    myTable.headers = [["Cities"]];
    myTable.rows = [['Berlin'], ['Roma'], ['Tokyo'], ['Seattle']];

    // Write table.
    Office.context.document.setSelectedDataAsync(myTable, {coercionType: Office.CoercionType.Table},
        function (result) {
            var error = result.error
            if (result.status === Office.AsyncResultStatus.Failed) {
                write(error.name + ": " + error.message);
            }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



 Если в выделенный фрагмент текста в Word записывается HTML-код, можно указать параметр _coercionType_ со значением `"html"`, как показано в следующем примере, в котором используются теги HTML `<b>`, выделяющие слово "Hello" полужирным шрифтом.




```js
function writeHtmlData() {
    Office.context.document.setSelectedDataAsync("<b>Hello</b> World!", {coercionType: Office.CoercionType.Html}, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            write('Error: ' + asyncResult.error.message);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Если в выделенный фрагмент текста в Word, PowerPoint или Excel записывается изображение, можно указать параметр _coercionType_ со значением `"image"`, как показано в приведенном ниже примере. Обратите внимание, что параметры imageLeft и imageTop игнорируются в Word.




```js
function insertPictureAtSelection(base64EncodedImageStr) {

    Office.context.document.setSelectedDataAsync(base64EncodedImageStr, {
       coercionType: Office.CoercionType.Image,
       imageLeft: 50,
       imageTop: 50,
       imageWidth: 100,
       imageHeight: 100
       },
       function (asyncResult) {
           if (asyncResult.status === Office.AsyncResultStatus.Failed) {
               console.log("Action failed with error: " + asyncResult.error.message);
           }
       });
}
```


## Сведения о поддержке


Флажок (![галочка](../../images/mod_off15_checkmark.png)) в приведенной ниже таблице указывает, что данный метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Поддерживаемые ведущие приложения по платформе**

||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Access**|![Галочка](../../images/mod_off15_checkmark.png)|||
|**Excel**|![Галочка](../../images/mod_off15_checkmark.png)|![Галочка](../../images/mod_off15_checkmark.png)|![Галочка](../../images/mod_off15_checkmark.png)|
|**PowerPoint**|![Галочка](../../images/mod_off15_checkmark.png)|![Галочка](../../images/mod_off15_checkmark.png)|![Галочка](../../images/mod_off15_checkmark.png)|
|**Word**|![Галочка](../../images/mod_off15_checkmark.png)|![Галочка](../../images/mod_off15_checkmark.png)|![Галочка](../../images/mod_off15_checkmark.png)|


|||
|:-----|:-----|
|**Доступен в наборах требований**|Выделение|
|**Минимальный уровень разрешений**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки




|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|В Word и Word Online добавлена поддержка записи данных в виде потока образ с кодировкой Base64.|
|1.1|В Word Online добавлена поддержка записи значений _data_ в виде **массива массивов** (матрицы) и **TableData** (таблицы).|
|1.1|В Excel, PowerPoint и Word из набора Office для iPad достигнут тот же уровень поддержки, что и в классических приложениях Excel, PowerPoint и Word для Windows.|
|1.1|В Word Online добавлена поддержка записи значений _data_ типа **string** (текст).|
|1.1|Добавлена поддержка [параметров форматирования при вставке таблиц](../../docs/excel/format-tables-in-add-ins-for-excel.md) в надстройках Excel с использованием необязательных параметров _tableOptions_ и _cellFormat_.|
|1.1|Добавлена поддержка записи табличных данных в надстройках Access.|
|1.0|Представлено|
