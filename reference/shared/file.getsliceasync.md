
# Метод File.getSliceAsync
Возвращает заданный фрагмент.

|||
|:-----|:-----|
|**Ведущие приложения:**|PowerPoint, Word|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Файл|
|**Добавлен в версии**|1.0|

```js
File.getSliceAsync(sliceIndex, callback);
```


## Параметры


_sliceIndex_ <br/>
&nbsp;&nbsp;&nbsp;&nbsp;Тип: **число**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Задает индекс извлекаемого фрагмента с отсчетом от нуля. Обязательный.<br/><br/>
    
_callback_ <br/>
&nbsp;&nbsp;&nbsp;&nbsp;Тип: **object**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Функция, вызываемая после получения результатов обратного вызова, единственный параметр которой имеет тип [AsyncResult](../../reference/shared/asyncresult.md). Необязательный.
    

## Значение обратного вызова

Когда выполняется функция, переданная в параметр _callback_, она получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью единственного параметра функции обратного вызова.

В функции обратного вызова, передаваемой в метод **getSliceAsync**, можно использовать свойства объекта **AsyncResult**, чтобы получить указанные ниже сведения.



|**Свойство**|**Применение**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Получает доступ к объекту [Slice](../../reference/shared/slice.md).|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Определяет, удалось ли выполнить операцию.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Получает доступ к объекту [Error](../../reference/shared/error.md), который содержит сведения об ошибке, если операция завершилась неудачно.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Получает доступ к определенному пользователем объекту **object** или значению, если они передаются как параметр _asyncContext_.|

## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).

||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Да|Y|
|**Word**|Y|Да|Y|

|||
|:-----|:-----|
|**Доступен в наборе требований**|Файл|
|**Минимальный уровень разрешений**|[ReadDocument (ReadAllDocument требуется для получения Office Open XML)](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка PowerPoint и Word в Office для iPad.|
|1.0|Представлено|
