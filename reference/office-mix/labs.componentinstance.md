
# Labs.ComponentInstance

 _**Область применения**: приложения для Office | надстройки Office | Office Mix | PowerPoint_

Представляет экземпляр определенного компонента, созданный для пользователя в среде выполнения. Объект содержит преобразованное представление компонента для определенного запуска лаборатории.

```
class ComponentInstance<T> extends Labs.ComponentInstanceBase
```


## Свойства

Нет.


## Методы




### constructor

 `function constructor()`

Инициализирует новый экземпляр класса **ComponentInstance**.


### createAttempt

 `public function createAttempt(callback: Labs.Core.ILabCallback<T>): void`

Создает новую попытку в контексте компонента.

 **Параметры**


|**Имя**|**Описание**|
|:-----|:-----|
| _callback_|Обратный вызов срабатывает, когда создается попытка.|

### getAttempts

 `public function getAttempts(callback: Labs.Core.ILabCallback<T[]>): void`

Получает все попытки, связанные с данным компонентом.

 **Параметры**


|**Имя**|**Описание**|
|:-----|:-----|
| _callback_|Обратный вызов срабатывает после получения попыток.|

### getCreateAttemptOptions

 `public function getCreateAttemptOptions(): Labs.Core.Actions.ICreateAttemptOptions`

Получает параметры создания попытки, заданные по умолчанию. Может быть переопределен производными классами.


### buildAttempt

 `public function buildAttempt(createAttemptResult: Labs.Core.IAction): T`

Создает попытку на основе заданного действия. Должен реализовываться с помощью производных классов.

 **Параметры**


|**Имя**|**Описание**|
|:-----|:-----|
| _createAttemptResult_|Действие создания попытки для указанной попытки.|
