
# Labs.Core.Actions.ICreateComponentOptions

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Создает новый компонент.

```
interface ICreateComponentOptions extends Core.IActionOptions
```


## Свойства


|||
|:-----|:-----|
| `componentId: string`|Компонент, вызывающий действие создания компонента.|
| `component: Core.IComponent`|Компонент [Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md), который требуется создать.|
| `correlationId?: string`|Необязательное поле для согласования этого компонента во всех экземплярах лаборатории. Позволяет ведущему приложению определить различные попытки, связанные с одним и тем же компонентом.|
