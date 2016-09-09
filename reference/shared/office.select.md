

# Метод Office.select
Создает резервирование для возврата привязки на основе передаваемой строки выбора.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Word|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, PartialTableBindings, TableBindings, TextBindings|
|**Последнее изменение в**|1.1|

```js
Office.select(str, onError);
```


## Параметры


_str_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Тип: **string**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Строка выбора для анализа и создания резервирования.

_onError_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Тип: **function**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Функция, вызываемая после получения результатов обратного вызова, единственный параметр которой имеет тип **AsyncResult**. Необязательный.
    

## Значение обратного вызова

При выполнении функции, переданной в параметре _onError_, она получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью единственного параметра функции обратного вызова. Если операция завершилась ошибкой, используйте свойство [AsyncResult.error](../../reference/shared/asyncresult.error.md) для доступа к объекту [Error](../../reference/shared/error.md), который предоставляет сведения об ошибке.


## Заметки

Метод **Office.select** предоставляет доступ к резервированию объекта [Binding](../../reference/shared/binding.md), который пытается вернуть указанную привязку при вызове любого своего асинхронного метода.

Поддерживаемые форматы: "bindings# _bindingId_", который возвращает объект **Binding** для привязки с помощью [идентификатора](../../reference/shared/binding.id.md) `bindingId`. Дополнительные сведения см. в статьях [Асинхронное программирование в надстройках для Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md#asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings) и [Привязка к областям в документе или электронной таблице](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md).


 >**Примечание**. Если резервирование метода **select** успешно возвращает объект **Binding**, то этот объект предоставляет только следующие четыре метода объекта [Binding](../../reference/shared/binding.md): [getDataAsync](../../reference/shared/binding.getdataasync.md), [setDataAsync](../../reference/shared/binding.setdataasync.md), [addHandlerAsync](../../reference/shared/binding.addhandlerasync.md) и [removeHandlerAsync](../../reference/shared/binding.removehandlerasync.md). Если резервированию не удается возвратить объект **Binding**, то для получения дополнительной информации можно обратиться к объекту [asyncResult.error](../../reference/shared/asyncresult.error.md) с помощью метода обратного вызова _onError_. Если вам необходимо вызвать элемент объекта **Binding**, которого нет среди четырех методов, предоставленных резервированием объекта **Binding**, который возвращен методом **select**, примените метод [getByIdAsync](../../reference/shared/bindings.getbyidasync.md). Для этого с помощью свойства [Document.bindings](../../reference/shared/document.bindings.md) и метода [Bindings.getByIdAsync](../../reference/shared/bindings.getbyidasync.md) получите объект **Binding**.


## Пример

В следующем примере кода используется метод **select**, чтобы получить привязку с **id** " `cities`" из коллекции **Bindings**, а затем вызвать метод [addHandlerAsync](../../reference/shared/binding.addhandlerasync.md) для добавления обработчика событий для события [dataChanged](../../reference/shared/binding.bindingdatachangedevent.md) привязки.


```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
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
|**Доступен в наборах требований**|MatrixBindings, PartialTableBindings, TableBindings, TextBindings|
|**Минимальный уровень разрешений**|[ReadDocument (ReadAllDocument для Open Office XML)](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel и Word в Office для iPad.|
|1.1|Добавлена возможность использовать метод **select** для возврата привязок таблиц, созданных в контентных надстройках для Access.|
|1.0|Представлено|
