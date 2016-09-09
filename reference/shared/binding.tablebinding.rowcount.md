
# Свойство TableBinding.rowCount
Получает количество строк в таблице. Целочисленное значение.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Word|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Последнее изменение в Selection**|1.1|

```
var rowCount = bindingObj.rowCount;
```


## Возвращаемое значение

Количество строк в указанном объекте [TableBinding](../../reference/shared/binding.tablebinding.md).


## Заметки

При вставке пустой таблицы, выбрав одну строку в Excel 2013 и Excel Online (используя пункт **Таблица** на вкладке **Вставка**), оба ведущих приложения Office создают одну строку заголовков, за которой следует пустая строка. Однако, если сценарий приложения создает привязку к вставленной таблице (например, с использованием метода [addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md)), а затем проверяет значение свойства **rowCount**, возвращенное значение будет отличаться в зависимости от того, где открыт лист: в Excel 2013 или Excel Online.


- В классическом приложении Excel **rowCount** возвращает значение 0 (не учитывается пустая строка, следующая после заголовков).
    
- В Excel Online **rowCount** возвращает значение 1 (учитывается пустая строка, следующая после заголовков).
    
Чтобы избежать разницы в сценарии, проверьте значение `rowCount == 1`, если да, проверьте все ли строки являются незаполненными.

В контентных надстройках для Access из соображений производительности свойство **rowCount** всегда возвращает значение -1.


## Пример




```js
function showBindingRowCount() {
    Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
        write("Rows: " + asyncResult.value.rowCount);
    });
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
|**Доступен в наборах требований**|TableBindings|
|**Минимальный уровень разрешений**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
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
