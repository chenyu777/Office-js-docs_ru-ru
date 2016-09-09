
# Labs.Components.DynamicComponentInstance

 _**Область применения**: приложения для Office | надстройки Office | Office Mix | PowerPoint_

Представляет экземпляр динамического компонента.

```
class DynamicComponentInstance extends Labs.ComponentInstanceBase
```


## Свойства


|Свойство|Описание|
|:-----|:-----|
| `public var component: Components.IDynamicComponentInstance`|Определение экземпляра компонента.|

## Методы




### constructor

 `function constructor(component: Components.IDynamicComponentInstance)`

Создает новый экземпляр динамического компонента с помощью определения [Labs.Components.IDynamicComponentInstance](../../reference/office-mix/labs.components.idynamiccomponentinstance.md).


### getComponents

 `public function getComponents(callback: Labs.Core.ILabCallback<Labs.ComponentInstanceBase[]>): void`

Получает все компоненты, созданные этим динамическим компонентом.

 **Параметры**


|Параметр|Описание|
|:-----|:-----|
| _callback_|Функция обратного вызова, которая срабатывает после получения всех компонентов.|

### createComponent

 `public function createComponent(component: Labs.Core.IComponent, callback: Labs.Core.ILabCallback<Labs.ComponentInstanceBase>): void`

Создает новый компонент, используя динамический компонент в качестве базы компонентов.

 **Параметры**


|Параметр|Описание|
|:-----|:-----|
| _component_|Компонент ([Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md)) для создания экземпляра.|
| _callback_|Функция обратного вызова, которая срабатывает после создания компонента.|

### close

 `public function close(callback: Labs.Core.ILabCallback<void>): void`

Указывает, что дополнительных отправок, связанных с этим экземпляром компонента, не будет.

 **Параметры**


|Параметр|Описание|
|:-----|:-----|
| _callback_|Функция обратного вызова, которая срабатывает после закрытия экземпляра.|

### isClosed

 `public function isClosed(callback: Labs.Core.ILabCallback<boolean>): void`

Указывает, закрыт ли динамический компонент. Возвращает значение **true**, если он закрыт.

