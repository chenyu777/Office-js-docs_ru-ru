
# Labs.connect (overload)

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Инициализирует соединение с ведущим приложением.

```
function connect(labHost: Core.ILabHost, callback: Core.ILabCallback<Core.IConnectionResponse>)
```


## Параметры


|||
|:-----|:-----|
| _labHost_|Необязательный. Экземпляр [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md), к которому нужно выполнить подключение. Если ведущее приложение не указано, оно будет создано с помощью [Labs.DefaultHostBuilder](../../reference/office-mix/labs.defaulthostbuilder.md).|
| _callback_|Параметр обратного вызова, который срабатывает после установления соединения.|

## Возвращаемое значение

Возвращает подключение к ведущему приложению.

