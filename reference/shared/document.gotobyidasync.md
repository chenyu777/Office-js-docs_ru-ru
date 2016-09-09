
# Метод Document.goToByIdAsync
Переходит к указанному объекту или месту в документе.

|||
|:-----|:-----|
|**Ведущие приложения:**|Excel, PowerPoint, Word|
|**Доступен в наборах требований**|Не в наборе|
|**Добавлено в версии**|1.1|

```js
Office.context.document.goToByIdAsync(id, goToType, [,options], callback);
```


## Параметры



|**Имя**|**Тип**|**Описание**|**Примечания по вопросам поддержки**|
|:-----|:-----|:-----|:-----|
| _id_|**string** или **number**|Идентификатор объекта или расположения для перехода. Обязательный параметр.||
| _goToType_|[GoToType](../../reference/shared/gototype-enumeration.md)|Тип расположения, к которому выполняется переход. Обязательный параметр.||
| _options_|**object**|Задает следующие [необязательные параметры](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _selectionMode_|[SelectionMode](../../reference/shared/selectionmode-enumeration.md)|Указывает, выбрано (выделено) ли расположение, заданное параметром _id_.|**В Excel:**<br/> **Office.SelectionMode.Selected** выбирает все содержимое привязки ли именованного элемента. <br/>**Office.SelectionMode.None** для привязок к тексту выбирает ячейку; для привязок к матрицам, таблицам и именованных элементов — выбирает первую ячейку с данными (а не первую ячейку в столбце заготовка таблицы).<br/><br/> **В PowerPoint:**<br/> **Office.SelectionMode.Selected** выбирает заголовок слайда или первое текстовое поле на слайде.<br/> **Office.SelectionMode.None** не выбирает ничего.<br/><br/> **В Word:**<br/> **Office.SelectionMode.Selected** выбирает все содержимое привязки. <br/>**Office.SelectionMode.None** для текстовых привязок перемещает указатель в начало текста, а для привязок к матрицам и таблицам выбирает первую ячейку с данными (а не первую ячейку строки заголовка таблицы).|
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** или **undefined**|Определяемый пользователем элемент любого типа, который возвращается в объекте **AsyncResult** без изменения.||
| _callback_|**object**|Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.||

## Значение обратного вызова

Когда выполняется функция, переданная в параметр _callback_, она получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью единственного параметра функции обратного вызова.

В функции обратного вызова, переданной методу **goToByIdAsync**, вы можете использовать свойства объекта **AsyncResult**, чтобы получить следующие сведения:



|**Свойство**|**Применение**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Возврат к текущему представлению.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Определяет, удалось ли выполнить операцию.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Получает доступ к объекту [Error](../../reference/shared/error.md), который содержит сведения об ошибке, если операция завершилась неудачно.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Получает доступ к определенному пользователем объекту **object** или значению, если они передаются как параметр _asyncContext_.|

## Заметки

PowerPoint не поддерживает метод **goToByIdAsync** в **режимах образцов**.


## Пример

 **Переход к привязке по идентификатору (Word и Excel)**

В следующем примере показано, как:


-  **создать привязку таблицы** с помощью метода [addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) как пример привязки для работы;
    
-  **указать привязку** для перехода;
    
-  **передать анонимную функцию обратного вызова**, которая возвращает состояние операции в параметр _callback_ метода **goToByIdAsync**;
    
-  **показать значение** на странице надстройки.
    



```js
function gotoBinding() {
    //Create a new table binding for the selected table.
    Office.context.document.bindings.addFromSelectionAsync("table",{ id: "MyTableBinding" }, function (asyncResult) {
    if (asyncResult.status == "failed") {
              showMessage("Action failed with error: " + asyncResult.error.message);
           }
           else {
              showMessage("Added new binding with type: " + asyncResult.value.type +" and id: " + asyncResult.value.id);
           }
    });

    //Go to binding by id.
    Office.context.document.goToByIdAsync("MyTableBinding", Office.GoToType.Binding, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}
```



 **Переход к таблице в электронной книге (Excel)**

В следующем примере показано, как:


-  **указать по имени таблицу** для перехода;
    
-  **передать анонимную функцию обратного вызова**, которая возвращает состояние операции в параметр _callback_ метода **goToByIdAsync**;
    
-  **показать значение** на странице надстройки.
    



```js
function goToTable() {
    Office.context.document.goToByIdAsync("Table1", Office.GoToType.NamedItem, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}
```



 **Переход к выбранному слайду по идентификатору (PowerPoint)**

В следующем примере показано, как:


-  **получить идентификатор** текущего выбранного слайда с помощью метода [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md);
    
-  **указать полученный идентификатор** как слайд для перехода;
    
-  **передать анонимную функцию обратного вызова**, которая возвращает состояние операции в параметр _callback_ метода **goToByIdAsync**;
    
-  **показать значение** преобразованного в строку объекта JSON, который был возвращен методом `asyncResult.value`, со сведениями о выбранных слайдах на странице надстройки.
    



```js
var firstSlideId = 0;
function gotoSelectedSlide() {
    //Get currently selected slide's id
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
    //Go to slide by id.
    Office.context.document.goToByIdAsync(firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```



 **Переход к слайду по индексу (PowerPoint)**

В следующем примере показано, как:


-  **указать индекс** первого, последнего или следующего слайда для перехода;
    
-  **передать анонимную функцию обратного вызова**, которая возвращает состояние операции в параметр _callback_ метода **goToByIdAsync**;
    
-  **показать значение** на странице надстройки.
    



```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}
```




## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Поддерживаемые ведущие приложения по платформе**


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Да|Y|
|**PowerPoint**|Y|Да|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|Не в наборе|
|**Минимальный уровень разрешений**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка PowerPoint Online.|
|1.1|Добавлена поддержка Excel, PowerPoint и Word в Office для iPad.|
|1.1|Представлено|
