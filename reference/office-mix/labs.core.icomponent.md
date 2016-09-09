
# Labs.Core.IComponent

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Базовый класс для представления компонентов лаборатории.

```
interface IComponent extends Core.ILabObject, Core.IUserData
```


## Свойства


|||
|:-----|:-----|
| `name: string`|Имя компонента.|
| `values: {[type:string]: Core.IValue[]}`|Схема свойств значений, связанная с компонентом.|
