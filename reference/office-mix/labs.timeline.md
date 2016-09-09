
# Labs.Timeline

 _**Область применения**: приложения для Office | надстройки Office | Office Mix | PowerPoint_

Предоставляет доступ к функции временной шкалы labs.js.

```
class Timeline
```


## Методы




### method

 `function constructor(labsInternal: Labs.LabsInternal)`

Создает новый экземпляр класса **Timeline**.


### next

 `public function next(completionStatus: Labs.Core.ICompletionStatus, callback: Labs.Core.ILabCallback<void>): void`

Указывает на то, что для временной шкалы следует выполнить переход к следующему слайду.

 **Параметры**


|||
|:-----|:-----|
| _completionStatus_|Показывает текущее состояние лаборатории.|
| _callback_|Функция обратного вызова, которая срабатывает, когда лаборатория переходит к следующему слайду.|
