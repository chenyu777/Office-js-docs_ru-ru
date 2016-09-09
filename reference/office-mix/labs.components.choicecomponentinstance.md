
# Labs.Components.ChoiceComponentInstance

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Представляет экземпляр компонента выбора.

```
class ChoiceComponentInstance extends Labs.ComponentInstance<Components.ChoiceComponentAttempt>
```


## Свойства


|Свойство|Описание|
|:-----|:-----|
| `public var component: Components.IChoiceComponentInstance`|Базовый объект [Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md), представляемый этим классом.|

## Методы




### constructor

 `function constructor(component: Components.IChoiceComponentInstance)`

Создает новый экземпляр класса **ChoiceComponentInstance**.

 **Параметры**


|Параметр|Описание|
|:-----|:-----|
| _компонент_|Объект [Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md) для создания этого класса.|

### buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.ChoiceComponentAttempt`

Создает новый экземпляр **ChoiceComponentAttempt** и реализует абстрактный метод, определенный для базового класса.

 **Параметры**


|Параметр|Описание|
|:-----|:-----|
| _createAttemptResult_|Результат действия создания попытки.|
