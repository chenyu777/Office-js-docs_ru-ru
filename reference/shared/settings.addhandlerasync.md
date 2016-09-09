

# Метод Settings.addHandlerAsync
Добавляет обработчик события **settingsChanged**.

|||
|:-----|:-----|
|**Ведущие приложения:**|Excel|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Параметры|
|**Последнее изменение в **|1.0|

```js
Office.context.document.settings.addHandlerAsync(eventType, handler [, options], callback);
```


## Параметры



|**Имя**|**Тип**|**Описание**|**Примечания по вопросам поддержки**|
|:-----|:-----|:-----|:-----|
| _eventType_|[EventType](../../reference/shared/eventtype-enumeration.md)|Указывает тип добавляемого события. Обязательный.||
| _handler_|**object**|Добавляемая функция обработчика событий. Обязательный.||
| _options_|**object**|Задает следующие [необязательные параметры](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods):||
| _asyncContext_|**array**, **boolean**, **null**, **number**, **object**, **string** или **undefined**|Определяемый пользователем элемент любого типа, который возвращается в объекте **AsyncResult** без изменения.||
| _callback_|**object**|Функция, вызываемая при возвращении обратного вызова, единственный параметр которой имеет тип **AsyncResult**.||

## Значение обратного вызова

Когда выполняется функция, переданная в параметр _callback_, она получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ с помощью единственного параметра функции обратного вызова.

В функции обратного вызова, переданной методу **addHandlerAsync**, вы можете использовать свойства объекта **AsyncResult**, чтобы получить следующие сведения:



|**Свойство**|**Применение**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Всегда возвращает значение **undefined**, так как при добавлении обработчика события нет данных или объектов, которые можно вернуть.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Определяет, удалось ли выполнить операцию.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Получает доступ к объекту [Error](../../reference/shared/error.md), который содержит сведения об ошибке, если операция завершилась неудачно.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Получает доступ к определенному пользователем объекту **object** или значению, если они передаются как параметр _asyncContext_.|

## Заметки

Вы можете добавить несколько обработчиков событий для указанного _eventType_, используя уникальное имя для каждой функции обработчика событий.


 >**Важно!** Обработчик события **settingsChanged** можно зарегистрировать с помощью кода вашей надстройки, когда эта надстройка работает с клиентом Excel. Но событие будет возникать, только если электронная таблица загружаемой надстройки открывается в Excel Online _и_ с ней работают несколько пользователей (совместное редактирование). Поэтому фактически событие **settingsChanged** поддерживается только в Excel Online со сценарием совместного редактирования.


## Пример




```js
function addSelectionChangedEventHandler() {
    Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged, MyHandler);
}

function MyHandler(eventArgs) {
    write('Event raised: ' + eventArgs.type);
    doSomethingWithSettings(eventArgs.settings);
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
|**Excel**||Y||

|||
|:-----|:-----|
|**Доступен в наборах требований**|Параметры|
|**Минимальный уровень разрешений**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки




|**Версия**|**Изменения**|
|:-----|:-----|
|1.0|Представлено|

