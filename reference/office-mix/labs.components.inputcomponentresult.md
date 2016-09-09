
# Labs.Components.InputComponentResult

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Результат отправки компонента ввода.

```
class InputComponentResult
```


## Свойства


|Свойство|Описание|
|:-----|:-----|
| `public var score: any`|Показатель, связанный с отправкой.|
| `public var complete: boolean`|Показывает, привел ли отправленный результат к завершению попытки.  Возвращает значение **True**, если попытка завершена.|

## Методы




### constructor

 `function constructor(score: any, complete: boolean)`

Создает новый экземпляр класса **InputComponentResult**.

 **Параметры**


|Параметр|Описание|
|:-----|:-----|
| _score_|Показатель, связанный с результатом.|
| _complete_|Возвращает логическое значение **true**, если результат завершил попытку.|
