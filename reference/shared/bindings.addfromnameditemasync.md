
# Метод Bindings.addFromNamedItemAsync
Добавляет привязку к именованному элементу в документе.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Word|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, TableBindings, TextBindings|
|**Последнее изменение**|1.1|

```
Office.context.document.bindings.addFromNamedItemAsync(itemName, bindingType [, options], callback);
```


## Параметры



|**Имя**|**Тип**|**Описание**|**Примечания по вопросам поддержки**|
|:-----|:-----|:-----|:-----|
| _itemName_|**string**|Имя именованного элемента. Обязательный параметр.||
| _bindingType_|[BindingType](../../reference/shared/bindingtype-enumeration.md)|Указывает тип объекта привязки для создания. Обязательный параметр. Возвращает **NULL**, если выбранный объект невозможно привести к указанному типу.||
| _options_|**object**|Задает следующие [необязательные параметры](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _id_|**string**|Задает уникальное имя, которое будет использоваться для определения нового объекта привязки. Если для параметра _id_ не передан аргумент, автоматически создается свойство [Binding.id](../../reference/shared/binding.id.md).||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** или **undefined**|Определяемый пользователем элемент любого типа, который возвращается в объекте **AsyncResult** без изменения.||
| _callback_|**object**|Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.||

## Значение обратного вызова

Когда выполняется функция, переданная в параметр _callback_, она получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ исключительно с помощью параметра функции обратного вызова.

В функции обратного вызова, переданной методу **addFromNamedItemAsync**, вы можете использовать свойства объекта **AsyncResult**, чтобы получить следующие сведения.



|**Свойство**|**Применение**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Получает доступ к объекту [Binding](../../reference/shared/binding.md), который представляет указанный именованный элемент.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Определяет, удалось ли выполнить операцию.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Получает доступ к объекту [Error](../../reference/shared/error.md), который содержит сведения об ошибке, если операция завершилась неудачно.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Получает доступ к определенному пользователем объекту **object** или значению, если они передаются как параметр _asyncContext_.|

## Замечания

 **Для Excel** параметр _itemName_ может ссылаться на именованный диапазон или на таблицу.

По умолчанию при добавлении таблиц в Excel имя "Table1" назначается первой добавленной таблице, "Table2" — второй таблице и так далее. Для назначения более осмысленного имени таблице в пользовательском интерфейсе Excel используйте свойство **Имя таблицы** на вкладке **Работа с таблицами | Макет** ленты.


 >**Примечание.** В Excel при задании таблицы в качестве именованного элемента необходимо указать ее имя полностью, включая имя листа, в таком формате: `"Sheet1!Table1"`.

 **Для Word** параметр _itemName_ ссылается на свойство **Title** элемента управления контентом **Форматированный текст**. (Привязку можно выполнять только к элементу управления контентом **Форматированный текст**.)

По умолчанию элементу управления контентом не назначено значение **Title**. Чтобы назначить понятное имя в пользовательском интерфейсе Word, после вставки элемента управления контентом  **Форматированный текст** из группы **Элементы управления** на вкладке **Разработчик** ленты выберите команду **Свойства** в группе **Элементы управления**, чтобы открыть диалоговое окно  **Свойства элемента управления контентом**. Затем задайте для свойства  **Title** элемента управления контентом имя, на которое вы будете ссылаться в коде.


 >**Примечание.** Если в Word имеется несколько элементов управления контентом **Форматированный текст** с одинаковым значением свойства **Title**, то при попытке выполнить привязку к одному из этих элементов управления с помощью данного метода (путем указания его имени в качестве параметра _itemName_) операция завершится сбоем.


## Пример

В следующем примере выполняется привязка типа "matrix" к именованному элементу `myRange` в Excel, а также и назначается [id](../../reference/shared/binding.id.md) привязки `myMatrix`.


```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("myRange", "matrix", {id:'myMatrix'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

В следующем примере выполняется привязка типа "table" к именованному элементу `Table1` в Excel, а также и назначается **id** привязки `myTable`.




```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("Table1", "table", {id:'myTable'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

В следующем примере показано, как в Word создать привязку текста к элементу управления контентом "Форматированный текст" с именем `"FirstName"`, назначить  **id**`"firstName"`, а затем отобразить эти сведения.




```js
function bindContentControl() {
    Office.context.document.bindings.addFromNamedItemAsync('FirstName', 
        Office.BindingType.Text, {id:'firstName'},
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                write('Control bound. Binding.id: '
                    + result.value.id + ' Binding.type: ' + result.value.type);
            } else {
                write('Error:', result.error.message);
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
|1.1|В надстройках для Excel можно создавать привязки таблиц (передав _bindingType_ как **Office.BindingType.Table**) для диапазона ячеек, который содержит табличные данные, даже если эти данные не были внесены в электронную таблицу в виде таблицы (с помощью команд **Вставка**  >  **Таблицы**  >  **Таблица** или **Главная**  >  **Стили**  >  **Форматировать как таблицу**).|
|1.1|Добавлена поддержка табличной привязки в контентных надстройках для Access. |
|1.0|Представлено|

## См. также



#### Другие ресурсы


[Привязка к областям в документе или электронной таблице](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md#add-a-binding-to-a-named-item)
