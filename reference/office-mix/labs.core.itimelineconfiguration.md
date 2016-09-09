
# Labs.Core.ITimelineConfiguration

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Параметры конфигурации для [Labs.Timeline](../../reference/office-mix/labs.timeline.md). Позволяет задать ряд параметров конфигурации временной шкалы.

```
interface ITimelineConfiguration
```


## Свойства


|||
|:-----|:-----|
| `duration: number`|Продолжительность работы лаборатории (в секундах).|
| `capabilities: string[]`|Список массива возможностей временной шкалы, которые поддерживаются лабораторией (например, воспроизведение, приостанавливание, поиск).|
