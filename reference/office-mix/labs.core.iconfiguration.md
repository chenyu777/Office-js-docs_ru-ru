
# Labs.Core.IConfiguration

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Структура данных конфигурации лаборатории.

```
interface IConfiguration extends Core.IUserData
```


## Свойства


|||
|:-----|:-----|
| `appVersion: Core.IVersion`|Версия приложения, связанного с этой конфигурацией.|
| `components: Core.IComponent[]`|Компоненты, входящие в состав лаборатории.|
| `name: string`|Имя лаборатории.|
| `timeline: Core.ITimelineConfiguration`|Конфигурация временной шкалы для лаборатории.|
| `analytics: Core.IAnalyticsConfiguration`|Конфигурация аналитики лаборатории.|
