
# Labs.Components.InputComponentInstance

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Представляет экземпляр компонента ввода.

```
class InputComponentInstance extends Labs.ComponentInstance<Components.InputComponentAttempt>
```


## Свойства


|Свойство|Описание|
|:-----|:-----|
| `public var component: Components.IInputComponentInstance`|Базовый объект [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md), представляемый этим классом.|

## Методы




### constructor

 `function constructor(component: Components.IInputComponentInstance)`

Создает новый экземпляр [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md).

 **Параметры**


|Параметр|Описание|
|:-----|:-----|
| _компонент_|[Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md) для создания этого класса.|

### buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.InputComponentAttempt`

Создает новый [Labs.Components.InputComponentAttempt](../../reference/office-mix/labs.components.inputcomponentattempt.md). Реализует абстрактный метод, определенный для базового класса.

 **Параметры**


|Параметр|Описание|
|:-----|:-----|
| _createAttemptResult_|Результат действия создания попытки.|
