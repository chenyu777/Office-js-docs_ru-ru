
# Labs.Components.ActivityComponentAttempt

 _**Область применения**: приложения для Office | надстройки Office | Office Mix | PowerPoint_

Представляет попытку завершения компонента действия.

```
class Permissions
```


## Методы




### constructor

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

Создает новый экземпляр класса **ActivityComponentAttempt**.

 **Параметры**


|**Имя**|**Описание**|
|:-----|:-----|
| _labs_|Экземпляры лаборатории ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx)), связанные с компонентом.|
| _componentId_|Идентификатор компонента, связанного с попыткой.|
| _attemptId_|Идентификатор попытки.|
| _values_|Значения, связанные с компонентом, если они есть.|

### complete

 `public function complete(callback: Labs.Core.ILabCallback<void>): void`

Индикатор завершения действия.

 **Параметры**


|**Имя**|**Описание**|
|:-----|:-----|
| _callback_|Функция обратного вызова, которая вызывается после завершения действия.|

### processAction

 `public function processAction(action: Labs.Core.IAction): void`

Функция, которая выполняет действия, полученные для заданной попытки, а затем заполняет состояние лаборатории.

 **Параметры**


|**Имя**|**Описание**|
|:-----|:-----|
| _action_|Экземпляр действия ([Labs.Core.IAction](../../reference/office-mix/labs.core.iaction.md)).|
