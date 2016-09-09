
# Метод Document.getSelectedDataAsync
Читает данные, содержащиеся в выбранном фрагменте документа.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, PowerPoint, Project, Word|
|**Доступен в наборах требований**|Выделение|
|**Последнее изменение в Selection**|1.1|

```js
Office.context.document.getSelectedDataAsync(coercionType [, options], callback); 
```


## Параметры



|**Имя**|**Тип**|**Описание**|**Примечания по вопросам поддержки**|
|:-----|:-----|:-----|:-----|
| _coercionType_|[CoercionType](../../reference/shared/coerciontype-enumeration.md)<br/><table><tr><td></td><td><b>Поддерживаемые ведущие приложения</b></td></tr><tr><td><b>Office.CoercionType.Text</b> (string)</td><td>Только Excel, Excel Online, PowerPoint, PowerPoint Online, Word и Word Online</td></tr><tr><td><b>Office.CoercionType.Matrix</b> (массив массивов)</td><td>Только Excel, Word и Word Online</td></tr><tr><td><b>Office.CoercionType.Table</b> (объект [TableData](../../reference/shared/tabledata.md))</td><td>Только Access, Excel, Word и Word Online</td></tr><tr><td><b>Office.CoercionType.Html</b></td><td>Только Word.</td></tr><tr><td><b>Office.CoercionType.Ooxml</b> (Office Open XML)</td><td>Только Word и Word Online</td></tr><tr><td><b>Office.CoercionType.SlideRange</b></td><td>Только PowerPoint и PowerPoint Online</td></tr></table>|Тип возвращаемой структуры данных. Обязательный.||
| _options_|**object**<br/><table><tr><td><i>valueFormat</i></td><td><b>[ValueFormat](../../reference/shared/valueformat-enumeration.md)</b></td><td>Указывает, форматируются ли значения номера и даты для возвращаемого результата.</td><td></td></tr><tr><td><i>filterType</i></td><td>[FilterType](../../reference/shared/filtertype-enumeration.md)</td><td>Указывает, применять ли фильтрацию при получении данных. Необязательный.</td><td>Этот параметр игнорируется в документах Word.</td></tr><tr><td><i>asyncContext</i></td><td><b>array</b>, <b>boolean</b>, <b>null</b>, <b>number</b>, <b>object</b>, <b>string</b>, или <b>undefined</b></td><td>Пользовательский элемент любого типа, возвращаемого в объекте <b>AsyncResult</b> без изменений.</td><td></td></tr></table>|Задает следующие [необязательные параметры](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _callback_|**object**|Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.||

## Значение обратного вызова

Когда выполняется функция, переданная в параметр _callback_, она получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью единственного параметра функции обратного вызова.

В функции обратного вызова, переданной методу **getSelectedDataAsync**, вы можете использовать свойства объекта **AsyncResult**, чтобы получить следующие сведения:



|**Свойство**|**Применение**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Доступ к значениям в выбранном фрагменте, которые возвращаются в структуре или формате, указанных в параметре _coercionType_. (Дополнительные сведения о приведении данных см. в разделе **Заметки**.)|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Определяет, удалось ли выполнить операцию.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Получает доступ к объекту [Error](../../reference/shared/error.md), который содержит сведения об ошибке, если операция завершилась неудачно.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Получает доступ к определенному пользователем объекту **object** или значению, если они передаются как параметр _asyncContext_.|

## Заметки

В надстройке области задач или контентной надстройке используйте метод **getSelectedDataAsync**, чтобы написать скрипт, считывающий данные из выделенного пользователем фрагмента в документе, электронной таблице, презентации или проекте. Например, когда пользователь выбирает содержимое в документе Word, вы можете использовать метод **getSelectedDataAsync**, чтобы считать этот фрагмент и отправить его веб-службе в качестве запроса или выполнить другую операцию.

После чтения выделения вы также можете использовать методы [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) и [addHandlerAsync](../../reference/shared/document.addhandlerasync.md) объекта **Document**, чтобы [выполнить обратную запись выделения или добавить обработчик события](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md) для обнаружения изменения выделения пользователем.

Метод **getSelectedDataAsync** может читать данные из выделенного фрагмента, только пока оно активно. Если вам необходимо установить постоянную связь для чтения и записи выделенного пользователем фрагмента в надстройках Word и Excel, используйте метод [Bindings.addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md), чтобы [установить привязку к этому выделению](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md).

С помощью параметра _coercionType_ метода **getSelectedDataAsync** можно указать структуру или формат считываемых данных.



|**Указанный параметр _coercionType_**|**Возвращаемые данные**|**Поддерживаемые ведущие приложения Office**|
|:-----|:-----|:-----|
|**Office.CoercionType.Text** или `"text"`|Строка.|Word, Excel, PowerPoint и Project.<br/><br/> **Примечание**. Если в Excel выбрано подмножество ячейки, возвращается все ее содержимое.|
|**Office.CoercionType.Matrix** или `"matrix"`|Массив массивов. Например, ` [['a','b'], ['c','d']]` для выбора двух строк в двух столбцах.|Word и Excel.|
|**Office.CoercionType.Table** или `"table"`|Объект [TableData](../../reference/shared/tabledata.md) для считывания таблицы с заголовками.|Word и Excel.|
|**Office.CoercionType.Html** или `"html"`|В формате HTML.|Только Word.|
|**Office.CoercionType.Ooxml** или `"ooxml"`|В формате Open Office XML (OpenXML).|Только Word.<br/><br/> **Совет**. При создании кода надстройки вы можете использовать `"ooxml"`_coercionType_ метода **getSelectedDataAsync**, чтобы узнать, как выбранное в документе Word содержимое определяется в тегах OpenXML. Затем используйте эти теги в параметре data метода [Document.setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md), чтобы записать в документ содержимое с этим форматом или структурой. Например, вы можете [вставить в документ изображение](http://blogs.msdn.com/b/officeapps/archive/2012/10/26/inserting-images-with-apps-for-office.aspx) в виде OpenXML.|
|**Office.CoercionType.SlideRange** или slideRange|Объект JSON, содержащий массив slides, который включает идентификаторы, названия и индексы выбранных слайдов. **Примечание.** Чтобы выбрать несколько слайдов, пользователь должен редактировать презентацию в представлении **обычном представлении**, **режиме структуры** или **режиме сортировщика слайдов**. Кроме того, этот метод не поддерживается в **режимах образца**. Например, `{"slides":[{"id":257,"title":"Slide 2","index":2},{"id":256,"title":"Slide 1","index":1}]}` позволяет выбрать два слайда.|Только PowerPoint.|
Если структура данных выделенного фрагмента не соответствует указанному типу _coercionType_, метод **getSelectedDataAsync** попытается привести данные к этому типу или структуре. Если выделенный фрагмент невозможно преобразовать в заданный тип **Office.CoercionType**, свойство **AsyncResult.status** возвращает значение `"failed"`.


## Пример

Для чтения значения текущего выделения необходимо написать функцию обратного вызова, которая читает выделение. В следующем примере показано, как:


-  **передать функцию обратного вызова**, которая считывает значение выделенного фрагмента в параметр _callback_ метода **getSelectedDataAsync**;
    
-  **считать выделение** как неформатированный текст без фильтров;
    
-  **показать значение** на странице надстройки.
    

```js
function getText() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
        { valueFormat: "unformatted", filterType: "all" },
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                write(error.name + ": " + error.message);
            } 
            else {
                // Get selected data.
                var dataValue = asyncResult.value; 
                write('Selected data is ' + dataValue);
            }            
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
|**PowerPoint**|Y|Да|Y|
|**Project**|Y|||
|**Word**|Y|Да|Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|Выделение|
|**Минимальный уровень разрешений**|[ReadDocument (ReadAllDocument требуется для получения Office Open XML)](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка PowerPoint Online.|
|1.1| В Word Online добавлена поддержка типов **Office.CoercionType.Matrix** и **Office.CoercionType.Table** для параметра _coercionType_.|
|1.1|В Excel, PowerPoint и Word из набора Office для iPad достигнут тот же уровень поддержки, что и в классических приложениях Excel, PowerPoint и Word для Windows.|
|1.1| В Word Online добавлена поддержка типа **Office.CoercionType.Text** для параметра _coercionType_.|
|1.1|В контентных надстройках для PowerPoint вы можете получить идентификаторы, заголовки и индексы для выбранного диапазона слайдов, передав **Office.CoercionType.SlideRange** как параметр _coercionType_ метода **getSelectedDataAsync**. В статье, посвященной методу [Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md), приводится наглядный пример того, как использовать это значение, чтобы перейти к выбранному слайду.|
|1.0|Представлено|
