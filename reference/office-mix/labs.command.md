
# Labs.Command

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Стандартная команда, используемая для передачи сообщений между клиентом и ведущим приложением.

```
class Command
```


## Свойства


|**Имя**|**Описание**|
|:-----|:-----|
| `public var type: string`|Тип команды.|
| `public var commandData: any`|Необязательные данные, связанные с этой командой.|

## Методы




### конструктор

 `function constructor(type: string, commandData?: any)`

Описание

 **Параметры**


|||
|:-----|:-----|
| `type`|Тип команды.|
| `commandData`|Необязательные данные, связанные с этой командой.|
