
# Labs.IMessageHandler

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Интерфейс, позволяющий определять обработчики событий.

```
interface IMessageHandler(origin: Window, data: any, callback: Labs.Core.ILabCallback<any>): void
```


## 

 **Параметры**


|||
|:-----|:-----|
| `origin`|Окно лаборатории, из которого исходит сообщение.|
| `data`|Содержание сообщения.|
| `callback`|Функция обратного вызова, которая срабатывает после получения сообщения.|
