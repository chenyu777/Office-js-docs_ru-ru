
# Labs.Components.ChoiceComponentAttempt

 _**Область применения**: приложения для Office | надстройки Office | Office Mix | PowerPoint_

Представляет попытку компонента выбора.

```
class ChoiceComponentAttempt extends Components.ComponentAttempt
```


## Методы




### constructor

 `function constructor(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

Создает новый экземпляр класса **ChoiceComponentAttempt**.

 **Параметры**


|**Имя**|**Описание**|
|:-----|:-----|
| _labs_|Экземпляр [Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx) для использования с попыткой.|
| _attemptId_|Идентификатор, связанный с попыткой.|
| _values_|Значения, связанные с попыткой.|

### timeout

 `public function timeout(callback: Labs.Core.ILabCallback<void>): void`

Указывает на то, что время ожидания лаборатории истекло.

 **Параметры**


|**Имя**|**Описание**|
|:-----|:-----|
| _callback_|Функции обратного вызова, которые срабатывают, когда сервер получает сообщение об истечении времени ожидания.|

### getSubmissions

 `public function getSubmissions(): Components.ChoiceComponentSubmission[]`

Получает все ранее отправленные ответы для заданной попытки.


### submit

 `public function submit(answer: Components.ChoiceComponentAnswer, result: Components.ChoiceComponentResult, callback: Labs.Core.ILabCallback<Components.ChoiceComponentSubmission>): void`

Отправляет новый ответ, который был рассчитан лабораторией, и для расчета которого не потребуется использование основного приложения.

 **Параметры**


|**Имя**|**Описание**|
|:-----|:-----|
| _answer_|Ответ для попытки.|
| _result_|Результат отправки.|
| _callback_|Функция обратного вызова, которая срабатывает после получения ответа.|

### processAction

 `public function processAction(action: Labs.Core.IAction): void`

Запускает обработку действия [Labs.Core.IAction](../../reference/office-mix/labs.core.iaction.md).

