
# Метод TableBinding.setTableOptionsAsync
Обновляет параметры форматирования привязанной таблицы.

|||
|:-----|:-----|
|**Ведущие приложения:**|Excel|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Не в наборе|
|**Добавлено в версии**|1.1|

```
bindingObj.setTableOptionsAsync(tableOptions [,options] , callback);
```


## Параметры



|**Имя**|**Тип**|**Описание**|**Примечания по вопросам поддержки**|
|:-----|:-----|:-----|:-----|
| _tableOptions_|**object**|Литерал объекта, содержащий список пар "имя-значение" для свойств, определяющих применяемые параметры таблицы. Обязательный параметр.||
| _options_|**object**|Задает следующие [необязательные параметры](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** или **undefined**|Определяемый пользователем элемент любого типа, который возвращается в объекте **AsyncResult** без изменения.||
| _callback_|**object**|Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.||

## Значение обратного вызова

Когда выполняется функция, переданная в параметр _callback_, она получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью единственного параметра функции обратного вызова.

В функции обратного вызова, переданной методу **goToByIdAsync**, вы можете использовать свойства объекта **AsyncResult**, чтобы получить следующие сведения:



|**Свойство**|**Применение**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Всегда возвращает значение **undefined**, так как данные или объекты, которые можно вернуть при задании параметров таблицы, отсутствуют.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Определяет, удалось ли выполнить операцию.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Получает доступ к объекту [Error](../../reference/shared/error.md), который содержит сведения об ошибке, если операция завершилась неудачно.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Получает доступ к определенному пользователем **object** или значению, если они передаются как параметр _asyncContext_.|

## Пример

В следующем примере показано, как:


-  **создать литерал объекта**, который указывает [параметры форматирования таблицы](../../docs/excel/format-tables-in-add-ins-for-excel.md) для обновления в связанной таблице;
    
-  **вызвать setTableOptions** в ранее связанной таблице (с **id**`myBinding`) с передачей объекта с параметром форматирования в качестве параметра _tableOptions_.
    

```js
function updateTableFormatting(){
    var tableOptions = {bandedRows: true, filterButton: false, style: "TableStyleMedium3"}; 

    Office.select("bindings#myBinding").setTableOptionsAsync(tableOptions, function(asyncResult){});
}
```




## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Поддерживаемые ведущие приложения по платформе**


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Да|Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|Не в наборе.|
|**Минимальный уровень разрешений**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки




|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel в Office для iPad.|
|1.1|Представлено|
