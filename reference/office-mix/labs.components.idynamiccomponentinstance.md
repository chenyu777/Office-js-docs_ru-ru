
# Labs.Components.IDynamicComponentInstance

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Экземпляр динамического компонента.

```
interface IDynamicComponentInstance extends Labs.Core.IComponentInstance
```


## Свойства


|Имя|Описание|
|:-----|:-----|
| `generatedComponentTypes: string[]`|Массив, содержащий типы компонентов, которые может создавать этот динамический компонент.|
| `maxComponents: number`|Максимальное количество компонентов, которые создаст этот динамический компонент. Или **Labs.Components.Infinite**, если ограничения нет.|
