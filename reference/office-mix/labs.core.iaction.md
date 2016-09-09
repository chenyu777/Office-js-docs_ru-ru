
# Labs.Core.IAction

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Представляет действие лаборатории, то есть взаимодействие между пользователем и указанной лабораторией.

```
interface IAction
```


## Свойства


|||
|:-----|:-----|
| `type: string`|Тип действия, выполняемого пользователем.|
| `options: Core.IActionOptions`|Параметры [Labs.Core.IActionOptions](../../reference/office-mix/labs.core.iactionoptions.md), отправленные с помощью действия, предпринятого пользователем.|
| `result: Core.IActionResult`|Результат действия [Labs.Core.IActionResult](../../reference/office-mix/labs.core.iactionresult.md).|
| `time: number`|Время выполнения действия, представленное в виде миллисекунд, прошедших с 00:00:00 1 января 1970 года по времени в формате UTC.|
