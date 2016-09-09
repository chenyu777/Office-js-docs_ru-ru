
# Метод Bindings.addFromSelectionAsync
Добавляет привязку для текущего выделения в документе.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Word|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, TableBindings, TextBindings|
|**Последнее изменение**|1.1|

```
bindingsObj.addFromSelectionAsync(bindingType [, options], callback);
```


## Параметры



|**Имя**|**Тип**|**Описание**|**Примечания по вопросам поддержки**|
|:-----|:-----|:-----|:-----|
| _bindingType_|[BindingType](../../reference/shared/bindingtype-enumeration.md)|Указывает тип объекта привязки для создания. Обязательный параметр. Возвращает **NULL**, если выбранный объект невозможно привести к указанному типу.||
| _options_|**object**|Задает следующие [необязательные параметры](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _id_|**string**|Задает уникальное имя, которое будет использоваться для определения нового объекта привязки. Если для параметра _id_ не передан аргумент, автоматически создается свойство [Binding.id](../../reference/shared/binding.id.md).||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** или **undefined**|Определяемый пользователем элемент любого типа, который возвращается в объекте **AsyncResult** без изменения.||
| _callback_|**object**|Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.||

## Значение обратного вызова

Когда выполняется функция, переданная в параметр _callback_, она получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ исключительно с помощью параметра функции обратного вызова.

В функции обратного вызова, переданной методу **addFromSelectionAsync**, можно использовать свойства объекта **AsyncResult**, чтобы получить следующие сведения.



|**Свойство**|**Применение**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Получает доступ к объекту [Binding](../../reference/shared/binding.md), который представляет выделение, заданное пользователем.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Определяет, удалось ли выполнить операцию.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Получает доступ к объекту [Error](../../reference/shared/error.md), который содержит сведения об ошибке, если операция завершилась неудачно.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Получает доступ к определенному пользователем объекту **object** или значению, если они передаются как параметр _asyncContext_.|

## Замечания

Добавляет указанный тип объекта привязки в коллекцию **Привязки**, которая будет идентифицироваться по заданному _id_.


 >**Примечание.** В приложении Excel, если вызвать метод **addFromSelectionAsync** и передать в него идентификатор **Binding.id** существующей привязки, используется тип [Binding.type](../../reference/shared/binding.type.md) этой привязки, который не удастся изменить, указав другое значение для параметра _bindingType_. Если требуется использовать существующий _id_ и изменить значение параметра _bindingType_, сначала вызовите метод [Bindings.releaseByIdAsync](../../reference/shared/bindings.releasebyidasync.md), чтобы удалить привязку, а затем — метод **addFromSelectionAsync**, чтобы повторно установить привязку нового типа.


## Пример

Добавляет [TextBinding](../../reference/shared/binding.textbinding.md) для текущего выделения с **Binding.id**, имеющим значение "MyBinding".


```js
function addBindingFromSelection() {
    Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: 'MyBinding' }, 
        function (asyncResult) {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
        }
    );
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


|**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|
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
