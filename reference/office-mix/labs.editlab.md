
# Labs.editLab

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Открывает указанную лабораторию для изменения. В режиме редактирования можно указать данные конфигурации лаборатории. Но изменить лабораторию, которая выполняется (то есть запущена), невозможно.

```
function editLab(callback: Core.ILabCallback<LabEditor>): void
```


## Параметры


|**Имя**|**Описание**|
|:-----|:-----|
| _callback_|Метод обратного вызова, который срабатывает при создании объекта [Labs.LabInstance](../../reference/office-mix/labs.labinstance.md).|
