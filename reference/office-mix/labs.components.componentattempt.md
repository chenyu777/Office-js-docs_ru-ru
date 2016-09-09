
# Labs.Components.ComponentAttempt

 _**Область применения**: приложения для Office | надстройки Office | Office Mix | PowerPoint_

Базовый класс для попыток компонентов.

```
class ComponentAttempt
```


## Свойства


|**Имя**|**Описание**|
|:-----|:-----|
| `public var _componentId: string`|Идентификатор указанного компонента.|
| `public var _id: string`|Идентификатор связанной лаборатории.|
| `public var _labs: Labs.LabsInternal`|Объект лаборатории ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx)), который используется для взаимодействия с базовым интерфейсом [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md).|
| `public var _resumed: boolean`|Возвращает значение **True**, если лаборатория возобновила обработку заданной попытки.|
| `public var _state: Labs.ProblemState`|Текущее состояние попытки, представленное перечислением [Labs.ProblemState](../../reference/office-mix/labs.problemstate.md).|
| `public var _values: { [type:string]: Labs.ValueHolder<any>[]}`|Значения, связанные с попыткой и содержащиеся в объекте [Labs.ValueHolder](../../reference/office-mix/labs.valueholder.md) (если таковые есть).|

## Методы




### constructor

 `(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

Создает новый экземпляр класса ComponentAttempt и предоставляет входные значения параметров.

 **Параметры**


|**Имя**|**Описание**|
|:-----|:-----|
| _labs_|Экземпляр [Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx) для использования с попыткой.|
| _attemptId_|Идентификатор, связанный с попыткой.|
| _values_|Массив значений ([Labs.Core.IValueInstance](../../reference/office-mix/labs.core.ivalueinstance.md)), связанный с попыткой.|

### isResumed

 `public function isResumed(): boolean`

Логическая функция, которая показывает, возобновлена ли работа лаборатории.  Возвращает значение **True**, если она возобновлена.

 **Параметры**

Нет.


### resume

 `public function resume(callback: Labs.Core.ILabCallback<void>): void`

Указывает, возобновила ли лаборатория обработку заданной попытки, и загружает существующие данные в рамках этого процесса. Попытку необходимо возобновить, прежде чем ее можно будет использовать.

 **Параметры**


|**Имя**|**Описание**|
|:-----|:-----|
| _callback_|Функция обратного вызова, которая срабатывает при возобновлении попытки.|

### getState

 `public function getState(): Labs.ProblemState`

Получает состояние лаборатории.

 **Параметры**

Нет.


### processAction

 `public function processAction(action: Labs.Core.IAction): void`

Выполняет действие, связанное с попыткой.

 **Параметры**

Нет.


### getValues

 `public function getValues(key: string): Labs.ValueHolder<any>[]`

Получает значения, связанные с попыткой.

 **Параметры**


|**Имя**|**Описание**|
|:-----|:-----|
| _key_|Ключ, связанный со значением в схеме значений.|
