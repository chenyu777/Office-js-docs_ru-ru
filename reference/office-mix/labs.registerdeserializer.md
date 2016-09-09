
# Labs.registerDeserializer

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Выполняет десериализацию указанного объекта JSON, преобразуя его в обычный объект. Эту функцию должны использовать только авторы компонентов.

```
function registerDeserializer(type: string, deserialize: (json: Core.ILabObject): any): void
```


## Параметры


|**Имя**|**Описание**|
|:-----|:-----|
|json|[Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md) для десериализации.|

## Возвращаемое значение

Возвращает экземпляр [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md).

