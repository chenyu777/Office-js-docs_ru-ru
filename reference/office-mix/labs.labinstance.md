
# Labs.LabInstance

 _**Область применения**: приложения для Office | надстройки Office | Office Mix | PowerPoint_

Экземпляр лаборатории, настроенный для текущего пользователя. Используйте этот объект, чтобы записывать и извлекать данные лаборатории для пользователя.

```
class LabInstance
```


## Переменные


|||
|:-----|:-----|
| `public var data: any`|Переменная контейнера для хранения данных пользователя.|
| `public var components: Labs.ComponentInstanceBase[]`|Компоненты, которые составляют экземпляр лаборатории.|

## Методы




### getState

 `public function getState(callback: Labs.Core.ILabCallback<any>): void`

Получает текущее состояние лаборатории для заданного пользователя.

 **Параметры**


|||
|:-----|:-----|
| _callback_|Функция обратного вызова, которая срабатывает при получении состояния лаборатории.|

### setState

 `public function setState(state: any, callback: Labs.Core.ILabCallback<void>): void`

Задает состояние лаборатории для указанного пользователя.

 **Параметры**


|||
|:-----|:-----|
| _state_|Состояние, которое требуется задать.|
| _callback_|Функция обратного вызова, которая срабатывает, когда состояние задано.|

### done

 `public function done(callback: Labs.Core.ILabCallback<void>): void`

Функция индикатора, указывающая на то, что пользователь завершил выполнение лаборатории.

 **Параметры**


|||
|:-----|:-----|
| _callback_|Функция обратного вызова, которая срабатывает после завершения работы лаборатории.|