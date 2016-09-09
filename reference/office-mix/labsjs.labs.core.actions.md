
# LabsJS.Labs.Core.Actions
Предоставляет общий обзор интерфейса API JavaScript для LabJS.Labs.Core.Actions.

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Эти интерфейсы API представляют операции лаборатории, указывая текущее ее поведение. Такие API удобны при создании новых компонентов или подключений к новому драйверу (кроме Office Mix).

## Модуль API LabsJS.Labs.Core.Actions

Модуль Actions содержит элементы указанных ниже типов.


### Интерфейсы


|||
|:-----|:-----|
|[Labs.Core.Actions.ICloseComponentOptions](../../reference/office-mix/labs.core.actions.iclosecomponentoptions.md)|Компонент, который требуется закрыть.|
|[Labs.Core.Actions.ICreateAttemptOptions](../../reference/office-mix/labs.core.actions.icreateattemptoptions.md)|Компонент, связанный с попыткой.|
|[Labs.Core.Actions.ICreateAttemptResult](../../reference/office-mix/labs.core.actions.icreateattemptresult.md)|Результат создания попытки для заданного компонента.|
|[Labs.Core.Actions.ICreateComponentOptions](../../reference/office-mix/labs.core.actions.icreatecomponentoptions.md)|Создает новый компонент.|
|[Labs.Core.Actions.ICreateComponentResult](../../reference/office-mix/labs.core.actions.icreatecomponentresult.md)|Результат [Labs.Core.IActionResult](../../reference/office-mix/labs.core.iactionresult.md) при создании нового компонента.|
|[Labs.Core.Actions.IGetValueResult](../../reference/office-mix/labs.core.actions.igetvalueresult.md)|Результат действия получения значения.|
|[Labs.Core.Actions.ISubmitAnswerResult](../../reference/office-mix/labs.core.actions.isubmitanswerresult.md)|Результат отправки ответа для попытки.|
|[Labs.Core.Actions.IAttemptTimeoutOptions](../../reference/office-mix/labs.core.actions.iattempttimeoutoptions.md)|Параметры, доступные для действия времени ожидания текущей попытки.|
|[Labs.Core.Actions.IGetValueOptions](../../reference/office-mix/labs.core.actions.igetvalueoptions.md)|Доступные параметры для операции получения значения.|
|[Labs.Core.Actions.IResumeAttemptOptions](../../reference/office-mix/labs.core.actions.iresumeattemptoptions.md)|Параметры, связанные с попыткой возобновления.|
|[Labs.Core.Actions.ISubmitAnswerOptions](../../reference/office-mix/labs.core.actions.isubmitansweroptions.md)|Параметры, доступные для действия отправки ответа.|

### Переменные


|||
|:-----|:-----|
| `var CloseComponentAction: string`|Закрывает компонент и указывает, что больше действий для него не будет.|
| `var CreateAttemptAction: string`|Действие создания новой попытки.|
| `var CreateComponentAction: string`|Действие создания нового компонента.|
| `var AttemptTimeoutAction: string`|Попытка выполнить действие времени ожидания.|
| `var GetValueAction: string`|Действие получения значения, связанного с попыткой.|
| `var ResumeAttemptAction: string`|Возобновление действия попытки. Позволяет указать, что пользователь возобновляет заданную попытку.|
| `var SubmitAnswerAction: string`|Действие отправки ответа для заданной попытки.|
