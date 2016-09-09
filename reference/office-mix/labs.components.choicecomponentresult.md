
# Labs.Components.ChoiceComponentResult

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Результат отправки компонента выбора.

```
class ChoiceComponentResult
```


## Свойства


|Свойство|Описание|
|:-----|:-----|
| `public var score: any`|Показатель, связанный с отправкой.|
| `public var complete: boolean`|Показывает, завершил ли результат попытку.  Возвращает значение **True**, если результат завершил попытку.|

## Методы




### constructor

 `function constructor(score: any, complete: boolean)`

Создает новый экземпляр класса **ChoiceComponentResult**.

 **Параметры**


|Параметр|Описание|
|:-----|:-----|
| _score_|Показатель результата.|
| _complete_|Показывает, завершил ли результат попытку.|
