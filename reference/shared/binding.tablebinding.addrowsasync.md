
# Метод TableBinding.addRowsAsync
Добавляет строки и значения в таблицу.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Word|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Последнее изменение в **|1.1|

```js
bindingObj.addRowsAsync(rows, [,options], callback);
```


## Параметры

_rows_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Тип: **array**.

&nbsp;&nbsp;&nbsp;&nbsp;Массив массивов, который содержит одну или несколько строк данных для добавления в таблицу. Обязательный.
    
_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Тип: **object**.

&nbsp;&nbsp;&nbsp;&nbsp;Задает приведенные ниже [необязательные параметры](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)
    
&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Тип: **array, boolean, null, number, object, string или undefined**.<br/><br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Пользовательский элемент любого типа, который возвращается в объекте [AsyncResult](../../reference/shared/asyncresult.md) без изменений. Необязательный.<br/><br/>

_callback_<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Тип: **object**.
    
&nbsp;&nbsp;&nbsp;&nbsp;Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип [AsyncResult](../../reference/shared/asyncresult.md). Необязательный.



|**Имя**|**Тип**|**Описание**|**Примечания по вопросам поддержки**|
|:-----|:-----|:-----|:-----|
| _rows_|**array**|Массив массивов, который содержит одну или несколько строк данных для добавления в таблицу. Обязательный.||
| _options_|**object**|Задает следующие [необязательные параметры](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** или **undefined**|Определяемый пользователем элемент любого типа, который возвращается в объекте **AsyncResult** без изменения.||
| _callback_|**object**|Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.||

## Значение обратного вызова

Когда выполняется функция, переданная в параметр _callback_, она получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью единственного параметра функции обратного вызова.

В функции обратного вызова, переданной в метод **addRowsAsync**, можно использовать свойства объекта **AsyncResult**, чтобы возвратить приведенные ниже данные.



|**Свойство**|**Применение**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Всегда возвращает значение **undefined**, так как объекты и данные не извлекаются.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Определяет, удалось ли выполнить операцию.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Получает доступ к объекту [Error](../../reference/shared/error.md), который содержит сведения об ошибке, если операция завершилась неудачно.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Получает доступ к определенному пользователем объекту **object** или значению, если они передаются как параметр _asyncContext_.|

## Заметки

Успешный или неудачный результат выполнения операции **addRowsAsync** является атомарным. То есть либо вся операция добавления строк выполняется успешно, либо происходит полный откат (и свойство **AsyncResult.status**, возвращенное в обратный вызов, будет содержать сведения об ошибке).


- Каждая строка массива, передаваемого в качестве аргумента _data_, должна содержать такое же число столбцов, как и обновляемая таблица. В противном случае вся операция завершится ошибкой.
    
- Все строки и ячейки массива должны быть успешно добавлены в новые строки таблицы. Если какая-либо строка или ячейка по какой-то причине не добавляется, вся операция завершается ошибкой.
    
 **Дополнительные заметки для приложения Excel Online**

Общее количество ячеек в значении, переданном параметру _rows_ при одном вызове этого метода, не может превышать 20 000.


## Пример




```js
function addRowsToTable() {
    Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
        var binding = asyncResult.value;
        binding.addRowsAsync([["6", "k"], ["7", "j"]]);
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
|**Доступен в наборах требований**|TableBindings|
|**Минимальный уровень разрешений**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки




|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel и Word в Office для iPad.|
|1.1|Добавлена поддержка записи табличных данных в надстройках Access.|
|1.0|Представлено|
