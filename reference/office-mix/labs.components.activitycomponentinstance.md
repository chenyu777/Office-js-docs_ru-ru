
# Labs.Components.ActivityComponentInstance

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Представляет текущий экземпляр компонента действия.

```
class ActivityComponentInstance extends Labs.ComponentInstance<Components.ActivityComponentAttempt>
```


## Свойства


|**Имя**|**Описание**|
|:-----|:-----|
| `public var component: Components.IActivityComponentInstance`|Базовый объект [Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md), представляемый этим классом.|

## Методы




### constructor

 `function constructor(component: Components.IActivityComponentInstance)`

Создает новый экземпляр класса [Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md).

 **Параметры**


|**Имя**|**Описание**|
|:-----|:-----|
| _компонент_|**IActivityComponentInstance** для создания класса из класса.|

### buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.ActivityComponentAttempt`

Создает новый экземпляр **ActivityComponentAttempt** и реализует абстрактный метод, определенный для базового класса.

 **Параметры**


|**Имя**|**Описание**|
|:-----|:-----|
| _createAttemptResult_|Результат действия создания попытки.|
