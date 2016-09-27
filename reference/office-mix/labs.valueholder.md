
# Labs.ValueHolder

 _**Область применения**: приложения для Office | надстройки Office | Office Mix | PowerPoint_

Объект-контейнер, который содержит и отслеживает значения для определенной лаборатории. Значение может храниться как локально, так и на сервере.

```
class ValueHolder<T>
```


## Переменные


|||
|:-----|:-----|
| `public var isHint: boolean`|Возвращает значение **True**, если значение является подсказкой.|
| `public var hasBeenRequested: boolean`|Возвращает значение **True**, если лаборатория запросила значение.|
| `public var hasValue: boolean`|Возвращает значение **True**, если необходимое значение уже содержится в контейнере значений.|
| `public var value: T`|Значение, которое хранится в контейнере.|
| `public var id: string`|Идентификатор значения.|

## Методы




### getValue

 `public function getValue(callback: Labs.Core.ILabCallback<T>): void`

Получает указанное значение.

 **Параметры**


|||
|:-----|:-----|
| _callback_|Функция обратного вызова, которая возвращает указанное значение.|

### provideValue

 `public function provideValue(value: T): void`

Внутренний метод, который предоставляет значение контейнеру значений.

 **Параметры**


|||
|:-----|:-----|
| _value_|Значение, которое нужно предоставить контейнеру значений.|