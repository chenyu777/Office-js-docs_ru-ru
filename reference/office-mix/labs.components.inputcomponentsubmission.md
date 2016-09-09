
# Labs.Components.InputComponentSubmission

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Представляет отправку в компонента ввода.

```
class InputComponentSubmission
```


## Свойства


|Свойство|Описание|
|:-----|:-----|
| `public var answer: Components.InputComponentAnswer`|Ответ ([Labs.Components.InputComponentAnswer](../../reference/office-mix/labs.components.inputcomponentanswer.md)), связанный с отправкой.|
| `public var result: Components.InputComponentResult`|Результат ([Labs.Components.InputComponentResult](../../reference/office-mix/labs.components.inputcomponentresult.md)) отправки.|
| `public var time: number`|Время получения отправленного элемента.|

## Методы




### constructor

 `function constructor(answer: Components.InputComponentAnswer, result: Components.InputComponentResult, time: number)`

Создает новый экземпляр класса **InputComponentSubmission**.

 **Параметры**


|Параметр|Описание|
|:-----|:-----|
| _answer_|Ответ, связанный с отправкой.|
| _result_|Результат отправки.|
| _время_|Время получения отправленного элемента.|
