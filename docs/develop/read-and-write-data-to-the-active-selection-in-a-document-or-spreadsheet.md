
# Считывание и запись данных в активное выделение документа или таблицы

Объект [Document](../../reference/shared/document.md) предоставляет методы, с помощью которых можно выполнять операции чтения и записи данных над текущим фрагментом, выделенным пользователем, в документе или электронной таблице. Для этого в объекте **Document** имеются методы **getSelectedDataAsync** и **setSelectedDataAsync**. Кроме того, в данной статье рассказывается, как считывать и записывать данные, а также создавать обработчики событий для обнаружения изменений в выделенном пользователем фрагменте.

Метод **getSelectedDataAsync** работает только для текущего фрагмента, выделенного пользователем. Если вам необходимо сохранить выделенный фрагмент в документе, чтобы он был доступен для чтения и записи во время последующих сеансов работы надстройки, необходимо добавить привязку с помощью метода [Bindings.addFromSelectionAsync](http://msdn.microsoft.com/en-us/library/edc99214-e63e-43f2-9392-97ead42fc155.aspx) (или создать привязку с помощью любого метода addFrom объекта [Bindings](http://msdn.microsoft.com/en-us/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1.aspx)). Дополнительные сведения о том, как создать привязку к области в документе, а также о чтении и записи данных через привязку см. в разделе [Привязка к областям в документе или электронной таблице](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md).


### Чтение выбранных данных


В примере ниже показано, как получить данные из выделенного фрагмента в документе с помощью метода [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md).


```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    }
    else {
        write('Selected data: ' + asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

В этом примере первый параметр _coercionType_ имеет значение **Office.CoercionType.Text** (вы также можете указать этот параметр, используя строку литерала `"text"`). Это означает, что свойство [value](../../reference/shared/asyncresult.status.md) объекта [AsyncResult](../../reference/shared/asyncresult.md), доступного из параметра _asyncResult_ в функции обратного вызова, возвратит **строку**, содержащую выделенный текст в документе. Если вы укажете какой-либо другой тип приведения, то получите другие значения. [Office.CoercionType](../../reference/shared/coerciontype-enumeration.md) — это перечисление значений доступных типов приведений. **Office.CoercionType.Text** имеет значение text.


 >**Совет**   **Информация о выборе матричного или табличного coercionType для доступа к данным.** Если предполагается динамический рост данных таблицы с добавлением строк и столбцов, при этом требуется обрабатывать заголовки таблицы, следует использовать табличные данные (указав параметр _coercionType_ метода **getSelectedDataAsync** в виде `"table"` или **Office.CoercionType.Table**). Добавление строк и столбцов в структуре данных поддерживается как табличными, так и матричными данными, но присоединение строк и столбцов поддерживается только табличными данными. В отсутствие необходимости добавления строк и столбцов, если при этом не требуется использовать заголовки, следует выбрать матричные данные (указав параметр  _coercionType_ метода **getSelecteDataAsync** в виде `"matrix"` или **Office.CoercionType.Matrix**), что позволяет использовать упрощенный способ взаимодействия с данными.

Анонимная функция, которая передается в функцию в качестве второго параметра _callback_, выполняется после завершения операции **getSelectedDataAsync**. При вызове функции передается один параметр _asyncResult_, который содержит результат и сведения о состоянии вызова. Если вызов завершается с ошибкой, свойство [error](../../reference/shared/asyncresult.context.md) объекта **AsyncResult** предоставляет доступ к объекту [Error](../../reference/shared/error.md). Вы можете проверить значение свойств [Error.name](../../reference/shared/error.name.md) и [Error.message](../../reference/shared/error.message.md), чтобы определить, почему операция завершилась с ошибкой. В противном случае будет отображен выделенный в документе текст.

Свойство [AsyncResult.status](../../reference/shared/asyncresult.error.md) используется в выражении **if** для проверки того, успешно ли выполнен вызов. [Office.AsyncResultStatus](../../reference/shared/asyncresultstatus-enumeration.md) — это перечисление доступных значений свойства **AsyncResult.status**. **Office.AsyncResultStatus.Failed** имеет значение failed (и его можно указать в виде строки литералов).


### Запись данных в выделение


В следующем примере показано, как записать в выделение строку "Hello World!".


```js
Office.context.document.setSelectedDataAsync("Hello World!", function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Передача в параметре _data_ других типов объектов может привести к разным результатам. Результат зависит от текущего выделения в документе, от ведущего приложения, а также от возможности приведения переданных данных применительно к текущему выделению.

Анонимная функция, которая передается в метод [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) в качестве параметра _callback_, выполняется после завершения асинхронного вызова. При записи данных в выделенный фрагмент с помощью метода **setSelectedDataAsync** параметр _asyncResult_ обратного вызова предоставляет доступ только к сведениям о состоянии вызова и к объекту [Error](../../reference/shared/error.md) в случае сбоя вызова.

 **Примечание.** Начиная с выпуска Excel 2013 с пакетом обновления 1 (SP1) и соответствующей сборки Excel Online вы можете [задать форматирование при записи таблицы в текущий выделенный фрагмент](../../docs/excel/format-tables-in-add-ins-for-excel.md).


### Обнаружение изменений в выделении


В примере ниже показано, как определять изменения в выделенном фрагменте, используя метод [Document.addHandlerAsync](../../reference/shared/document.addhandlerasync.md) для добавления обработчика события [SelectionChanged](../../reference/shared/document.selectionchanged.event.md) в документе.


```
Office.context.document.addHandlerAsync("documentSelectionChanged", myHandler, function(result){} 
);

// Event handler function.
function myHandler(eventArgs){
write('Document Selection Changed');
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Первый параметр _eventType_ задает имя события для подписки. Передача строки `"documentSelectionChanged"` для этого параметра эквивалентна передаче типа события **Office.EventType.DocumentSelectionChanged** перечисления [Office.EventType](../../reference/shared/eventtype-enumeration.md).

Анонимная функция `myHander()`, передаваемая в эту функцию в качестве второго параметра _handler_, представляет собой обработчик событий, который выполняется при изменении выделенного фрагмента в документе. При вызове этой функции передается единственный параметр _eventArgs_, который после завершения асинхронной операции будет содержать ссылку на объект [DocumentSelectionChangedEventArgs](../../reference/shared/document.selectionchangedeventargs.md). Вы можете использовать свойство [DocumentSelectionChangedEventArgs.document](../../reference/shared/document.selectionchangedeventargs.document.md) для доступа к документу, создавшему событие.


 >**Примечание.** Вы можете добавить несколько обработчиков событий для данного события, снова вызвав метод **addHandlerAsync** и передав дополнительную функцию обработчика события для параметра _handler_. Это будет работать правильно, поскольку имя каждой функции обработчика событий уникально.


### Отключение обнаружения изменений в выделении


В примере ниже показано, как остановить прослушивание события [Document.SelectionChanged](../../reference/shared/document.selectionchanged.event.md), вызвав метод [document.removeHandlerAsync](../../reference/shared/document.removehandlerasync.md).


```
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

Имя функции `myHandler`, передаваемое в качестве второго параметра _handler_, задает обработчик событий, который будет удален из события **SelectionChanged**.


 >**Важно!** Если при вызове метода _removeHandlerAsync_ вы не укажете необязательный параметр **handler**, то все обработчики событий для указанного объекта _eventType_ будут удалены.

