
# Метод Settings.saveAsync
Хранится в содержащейся в памяти копии контейнера свойств параметров документа.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, PowerPoint, Word|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Параметры|
|**Последнее изменение в **|1.1|

```js
Office.context.document.settings.saveAsync(callback);
```


## Параметры



_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Тип: **object**

&nbsp;&nbsp;&nbsp;&nbsp;Функция, вызываемая после получения результатов обратного вызова, единственный параметр которой имеет тип **AsyncResult**. Необязательный.

    



## Значение обратного вызова

Когда выполняется функция, переданная в параметр _callback_, она получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью единственного параметра функции обратного вызова.

В функции обратного вызова, переданной методу **saveAsync**, вы можете использовать свойства объекта **AsyncResult**, чтобы получить следующие сведения.



|**Свойство**|**Применение**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Всегда возвращает значение **undefined**, так как объекты и данные не извлекаются.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Определяет, удалось ли выполнить операцию.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Получает доступ к объекту [Error](../../reference/shared/error.md), который содержит сведения об ошибке, если операция завершилась неудачно.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Получает доступ к определенному пользователем объекту **object** или значению, если они передаются как параметр _asyncContext_.|

## Заметки

Все параметры, ранее сохраненные надстройкой, загружаются при ее инициализации, поэтому на протяжении всего сеанса можно использовать только методы [set](../../reference/shared/settings.set.md) и [get](../../reference/shared/settings.get.md) для работы с копией контейнера свойств в памяти. Если требуется сохранить параметры, чтобы они были доступны при следующем использовании надстройки, воспользуйтесь методом **saveAsync**.


 >**Примечание.** Метод **saveAsync** сохраняет содержащуюся в памяти копию контейнера свойств параметров в файле документа; но изменения, внесенные непосредственно в файл документа, сохраняются только при выполнении пользователем сохранения документа в файловой системе (или параметра **AutoRecover**).

Метод [refreshAsync](../../reference/shared/settings.refreshasync.md) полезен только в сценариях совместной работы над документами (которые поддерживаются только в Word), когда каждый экземпляр надстройки может изменить параметры и эти изменения необходимо сделать доступными для всех экземпляров.


## Пример




```js
function persistSettings() {
    Office.context.document.settings.saveAsync(function (asyncResult) {
        write('Settings saved with status: ' + asyncResult.status);
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
|**PowerPoint**|Y|Да|Y|
|**Word**|Y|Да|Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|Параметры|
|**Минимальный уровень разрешений**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки




|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка PowerPoint Online.|
|1.1|Добавлена поддержка Excel, PowerPoint и Word в Office для iPad.|
|1.1|Добавлена поддержка настраиваемых параметров в контекстных надстройках для Access.|
|1.0|Представлено|
