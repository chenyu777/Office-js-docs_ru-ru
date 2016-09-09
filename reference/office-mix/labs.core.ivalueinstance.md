
# Labs.Core.IValueInstance

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Экземпляр объекта [Labs.Core.IValue](../../reference/office-mix/labs.core.ivalue.md), который содержит данные значения, если они есть.

```
interface IValueInstance
```


## Свойства


|||
|:-----|:-----|
| `valueId: string`|Идентификатор представленного этим экземпляром значения.|
| `isHint: boolean`|Возвращает логическое значение **true**, если данное значение считается подсказкой.|
| `hasValue: boolean`|Возвращает логическое значение **true**, если данные экземпляра содержат значение.|
| `value?: any`|Значение. Этот параметр может быть или не быть задан в зависимости от того, скрытый ли он.|
