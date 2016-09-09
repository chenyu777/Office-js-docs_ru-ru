
# Метод Document.getFileAsync
Возвращает полный файл документа фрагментами размером до 4194304 байт (4 МБ). Надстройки для iOS поддерживают фрагменты до 65536 байт (64 КБ). Обратите внимание, что если указать размер фрагмента выше допустимого, возникнет сбой "Внутренняя ошибка". 

|||
|:-----|:-----|
|**Ведущие приложения:**|Excel, PowerPoint, Word|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Файл|
|**Последнее изменение в File**|1.1|

```js
Office.context.document.getFileAsync(fileType [, options], callback);
```


## Параметры



|**Имя**|**Тип**|**Описание**|**Примечания по вопросам поддержки**|
|:-----|:-----|:-----|:-----|
| _fileType_|[FileType](../../reference/shared/filetype-enumeration.md)|Указывает формат возвращаемого файла. Обязательный.<br/><table><tr><th>Ведущее приложение</th><th>Поддерживаемый тип fileType</th></tr><tr><td>Excel Online</td><td>Office.FileType.Compressed</td></tr><tr><td>PowerPoint в Windows для настольных компьютеров</td><td>Office.FileType.Compressed, Office.FileType.Pdf</td></tr><tr><td>Word для настольных компьютеров с Windows, MAC и iPad</td><td>Office.FileType.Compressed, Office.FileType.Pdf, Office.FileType.Text</td></tr><tr><td>Word Online</td><td>Office.FileType.Compressed, Office.FileType.Pdf, Office.FileType.Text</td></tr><tr><td>PowerPoint Online</td><td>Office.FileType.Compressed, Office.FileType.Pdf</td></tr></table>|**Изменен в** версии 1.1, см. раздел [История поддержки](#История-поддержки)|
| _options_|**object**|Задает следующие [необязательные параметры](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _sliceSize_|**number**|Задает требуемый размер фрагмента (в байтах) до 4194304 байт (4 МБ). Если не значение не задано, будет использоваться размер фрагмента по умолчанию: 4194304 байт (4 МБ). ||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** или **undefined**|Определяемый пользователем элемент любого типа, который возвращается в объекте **AsyncResult** без изменения.||
| _callback_|**object**|Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.||

## Значение обратного вызова

Когда выполняется функция, переданная в параметр _callback_, она получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью единственного параметра функции обратного вызова.

Если функция обратного вызова передана методу **getFileAsync**, можно использовать свойства объекта **AsyncResult** для возврата следующей информации:



|**Свойство**|**Применение**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Получает доступ к объекту [File](../../reference/shared/file.md).|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Определяет, удалось ли выполнить операцию.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Получает доступ к объекту [Error](../../reference/shared/error.md), который содержит сведения об ошибке, если операция завершилась неудачно.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Получает доступ к определенному пользователем объекту **object** или значению, если они передаются как параметр _asyncContext_.|

## Заметки

Для надстроек, работающих в ведущих приложениях Office (кроме Office для iOS), метод **getFileAsync** поддерживает получение файлов по фрагментам размером до 4194304 байт (4 МБ). Для надстроек, работающих в приложениях Office для iOS, метод **getFileAsync** поддерживает получение файлов по фрагментам размером до 65536 байтов (64 КБ).

Параметр _fileType_ можно задать, используя указанные ниже перечисления или текстовые значения.


**Перечисление FileType**


|**Перечисление**|**Значение**|**Описание**|
|:-----|:-----|:-----|
|Office.FileType.Compressed|"compressed"|Возвращает весь документ (DOCX, PPTX или XSLX) в формате Office Open XML (OOXML) в виде массива байтов.|
|Office.FileType.Pdf|"pdf"|Возвращает весь документ в формате PDF в виде массива байтов.|
|Office.FileType.Text|"text"|Возвращает только текст документа в виде **string**. |
В памяти может храниться не более двух документов. В противном случае произойдет сбой операции **getFileAsync**. Используйте метод [File.closeAsync](../../reference/shared/file.closeasync.md), чтобы закрыть файл после завершения работы с ним.


## Пример — получение документа в формате Office Open XML ("сжатый")

В примере ниже показано, как получить документ в формате Office Open XML ("сжатый") в срезах размером 65536 байтов (64 КБ). Примечание. В данном примере реализация `app.showNotification` взята из шаблона Visual Studio для надстроек Office.


```js
function getDocumentAsCompressed() {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 /*64 KB*/ }, 
        function (result) {
            if (result.status == "succeeded") {
            // If the getFileAsync call succeeded, then
            // result.value will return a valid File Object.
            var myFile = result.value;
            var sliceCount = myFile.sliceCount;
            var slicesReceived = 0, gotAllSlices = true, docdataSlices = [];
            app.showNotification("File size:" + myFile.size + " #Slices: " + sliceCount);

            // Get the file slices.
            getSliceAsync(myFile, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
            }
            else {
            app.showNotification("Error:", result.error.message);
            }
    });
}

function getSliceAsync(file, nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived) {
    file.getSliceAsync(nextSlice, function (sliceResult) {
        if (sliceResult.status == "succeeded") {
            if (!gotAllSlices) { // Failed to get all slices, no need to continue.
                return;
            }

            // Got one slice, store it in a temporary array.
            // (Or you can do something else, such as
            // send it to a third-party server.)
            docdataSlices[sliceResult.value.index] = sliceResult.value.data;
            if (++slicesReceived == sliceCount) {
               // All slices have been received.
               file.closeAsync();
               onGotAllSlices(docdataSlices);
            }
            else {
                getSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
            }
        }
            else {
                gotAllSlices = false;
                file.closeAsync();
                app.showNotification("getSliceAsync Error:", sliceResult.error.message);
            }
    });
}

function onGotAllSlices(docdataSlices) {
    var docdata = [];
    for (var i = 0; i < docdataSlices.length; i++) {
        docdata = docdata.concat(docdataSlices[i]);
    }

    var fileContent = new String();
    for (var j = 0; j < docdata.length; j++) {
        fileContent += String.fromCharCode(docdata[j]);
    }

    // Now all the file content is stored in 'fileContent' variable,
    // you can do something with it, such as print, fax...
}

```


## Пример — получение документа в формате PDF

В приведенном ниже примере возвращается документ в формате PDF.


```js
Office.context.document.getFileAsync(Office.FileType.Pdf,
    function(result) {
        if (result.status == "succeeded") {
            var myFile = result.value;
            var sliceCount = myFile.sliceCount;
            app.showNotification("File size:" + myFile.size + " #Slices: " + sliceCount);
            // Now, you can call getSliceAsync to download the files, as described in the previous code segment (compressed format).
            
            myFile.closeAsync();
        }
        else {
            app.showNotification("Error:", result.error.message);
        }
}
);


```


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Поддерживаемые ведущие приложения по платформе**


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||Y||
|**PowerPoint**|Y|Да|Y|
|**Word**|Y|Да|Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|Файл|
|**Минимальный уровень разрешений**|[ReadAllDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки


|**Версия**|**Изменения**|
|:-----|:-----|
|1.1| В PowerPoint Online добавлена поддержка типа **Office.FileType.Pdf** для параметра _fileType_.|
|1.1| В PowerPoint Online добавлена поддержка типа **Office.FileType.Compressed** для параметра _fileType_.|
|1.1| В Word Online добавлена поддержка типа **Office.FileType.Text** для параметра _fileType_.|
|1.1| В Excel Online добавлена поддержка типа **Office.FileType.Compressed** для параметра _fileType_.|
|1.1| В Word Online добавлена поддержка типов **Office.FileType.Compressed** и **Office.FileType.Pdf** для параметра _fileType_.|
|1.1|В PowerPoint и Word из набора Office для iPad добавлена поддержка всех значений **FileType** для параметра _fileType_.|
|1.1|В классических приложениях Word и PowerPoint для Windows добавлена поддержка типа **Office.FileType.Pdf** для параметра _fileType_.|
|1.0|Представлено|
