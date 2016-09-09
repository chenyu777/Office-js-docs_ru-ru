
# Labs.Components.InputComponentAttempt

 _**Область применения**: приложения для Office | надстройки Office | Office Mix | PowerPoint_

Представляет попытку взаимодействия с компонентом ввода.

```
class InputComponentAttempt extends Components.ComponentAttempt
```


## Методы




### constructor

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

Создает новый экземпляр класса **InputComponentAttempt**.

 **Параметры**


|Параметр|Описание|
|:-----|:-----|
| _labs_|Лаборатории ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx)), связанные с попыткой.|
| _componentID_|Идентификатор компонента, связанного с попыткой.|
| _attemptId_|Идентификатор определенной попытки.|
| _values_|Массив, содержащий экземпляры значения ([Labs.Core.IValueInstance](../../reference/office-mix/labs.core.ivalueinstance.md)).|

### processAction

 `public function processAction(action: Labs.Core.IAction): void`

Выполняет итерацию полученных действий для указанной попытки и заполняет состояние лаборатории.

 **Параметры**


|Параметр|Описание|
|:-----|:-----|
| _action_|Действие, связанное с состоянием лаборатории.|

### getSubmissions

 `public function getSubmissions(): Components.InputComponentSubmission[]`

Получает все ранее отправленные ответы для заданной попытки.


### submit

 `public function submit(answer: Components.InputComponentAnswer, result: Components.InputComponentResult, callback: Labs.Core.ILabCallback<Components.InputComponentSubmission>): void`

Отправляет новый ответ, который был рассчитан лабораторией, и для расчета которого не потребуется использование основного приложения.

 **Параметры**


|Параметр|Описание|
|:-----|:-----|
| _answer_|Ответ, связанный с попыткой.|
| _result_|Результат, связанный с отправкой.|
| _callback_|Функция обратного вызова, которая срабатывает после получения ответа.|
