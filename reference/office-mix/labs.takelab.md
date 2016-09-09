
# Labs.takeLab

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Запускает указанную лабораторию и позволяет отправлять результаты работы в лаборатории на сервер. Обратите внимание, что невозможно запустить лабораторию во время ее изменения.

```
function takeLab(callback: Core.ILabCallback<LabInstance>): void
```


## Параметры


|**Имя**|**Описание**|
|:-----|:-----|
| _callback_|Метод обратного вызова, который срабатывает при создании объекта [Labs.LabInstance](../../reference/office-mix/labs.labinstance.md).|
