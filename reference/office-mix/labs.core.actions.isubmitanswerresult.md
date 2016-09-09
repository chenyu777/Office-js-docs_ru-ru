
# Labs.Core.Actions.ISubmitAnswerResult

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Результат отправки ответа для попытки.

```
interface ISubmitAnswerResult extends Core.IActionResult
```


## Свойства


|||
|:-----|:-----|
| `submissionId: string`|Идентификатор, связанный с отправкой. Предоставлен сервером.|
| `complete: boolean`|Возвращает значение **true**, если попытка завершается из-за текущей отправки.|
| `score: any`|Сведения о показателе, связанные с отправкой.|
