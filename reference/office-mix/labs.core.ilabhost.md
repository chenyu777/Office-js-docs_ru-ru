
# Labs.Core.ILabHost

 _**Область применения**: приложения для Office | надстройки Office | Office Mix | PowerPoint_

Предоставляет уровень абстракции для подключения Labs.js к основному приложению.

```
interface ILabHost
```


## Методы


### getSupportedVersions

 `getSupportedVersions(): Core.ILabHostVersionInfo[]`

Получает версии, поддерживаемые узлом лаборатории.

 **Параметры**

Нет.


### connect

 `connect(versions: Core.ILabHostVersionInfo[], callback: Core.ILabCallback<Core.IConnectionResponse>)`

Инициализирует соединение с основным приложением.

 **Параметры**


|||
|:-----|:-----|
| _версии_|Список версий основного приложения, которые может использовать клиент.|
| _callback_|Функция обратного вызова, которая срабатывает при соединении.|

### disconnect

 `disconnect(callback: Core.ILabCallback<void>)`

Разрывает соединение с основным приложением.

 **Параметры**


|||
|:-----|:-----|
| _completionStatus_|Состояние лаборатории на момент отключения.|
| _callback_|Функция обратного вызова, которая срабатывает при отключении.|

### on

 `on(handler: (string: any, any: any): void)`

Добавляет обработчик событий для обработки сообщений, поступающих от основного приложения. Разрешенное обещание будет возвращено основному приложению.

 **Параметры**


|||
|:-----|:-----|
| _handler_|Обработчик событий.|

### sendMessage

 `sendMessage(type: string, options: Core.IMessage, callback: Core.ILabCallback<Core.IMessageResponse>)`

Отправляет сообщение основному приложению.

 **Параметры**


|||
|:-----|:-----|
| _type_|Тип отправляемого сообщения.|
| _options_|Параметры сообщения.|
| _callback_|Функция обратного вызова, которая срабатывает после получения сообщения.|

### create

 `create(options: Core.ILabCreationOptions, callback: Core.ILabCallback<void>)`

Создает лабораторию. Хранит данные основного приложения и выделяет место для хранения конфигурации и прочих элементов.

 **Параметры**


|||
|:-----|:-----|
| _options_|Параметры, которые передаются при операции create.|
| _callback_|Функция обратного вызова, которая срабатывает после создания лаборатории.|

### getConfiguration

 `getConfiguration(callback: Core.ILabCallback<Core.IConfiguration>)`

Получает текущую конфигурацию лаборатории от основного приложения.

 **Параметры**


|||
|:-----|:-----|
| _callback_|Функция обратного вызова для получения сведений о конфигурации.|

### setConfiguration

 `setConfiguration(configuration: Core.IConfiguration, callback: Core.ILabCallback<void>)`

Задает новую конфигурацию лаборатории для основного приложения.

 **Параметры**


|||
|:-----|:-----|
| _configuration_|Заданная конфигурация лаборатории.|
| _callback_|Функция обратного вызова, которая срабатывает, когда конфигурация задана.|

### getConfigurationInstance

 `getConfigurationInstance(callback: Core.ILabCallback<Core.IConfigurationInstance>)`

Получает конфигурацию экземпляра для лаборатории.

 **Параметры**


|||
|:-----|:-----|
| _callback_|Функция обратного вызова, которая срабатывает после получения экземпляра конфигурации.|

### getState

 `getState(callback: Core.ILabCallback<any>)`

Получает текущее состояние лаборатории для заданного пользователя.

 **Параметры**


|||
|:-----|:-----|
| _completionStatus_|Функция обратного вызова, которая возвращает текущее состояние лаборатории.|

### setState

 `setState(state: any, callback: Core.ILabCallback<void>)`

Задает состояние лаборатории для указанного пользователя.

 **Параметры**


|||
|:-----|:-----|
| _state_|Состояние лаборатории.|
| _callback_|Функция обратного вызова, которая срабатывает, когда состояние задано.|

### takeAction

 `takeAction(type: string, options: Core.IActionOptions, callback: Core.ILabCallback<Core.IAction>)`

Предпринимает попытку совершить действие.

 **Параметры**


|||
|:-----|:-----|
| _type_|Тип действия.|
| _options_|Параметры, предоставленные с помощью действия.|
| _callback_|Функция обратного вызова, которая возвращает последнее выполненное действие.|

### takeAction

 `takeAction(type: string, options: Core.IActionOptions, result: Core.IActionResult, callback: Core.ILabCallback<Core.IAction>)`

Предпринимает действие, которое уже выполнено.

 **Параметры**


|||
|:-----|:-----|
| _type_|Тип действия.|
| _options_|Параметры, предоставленные с помощью действия.|
| _result_|Результат действия.|
| _callback_|Функция обратного вызова, которая возвращает последнее выполненное действие.|

### getActions

 `getActions(type: string, options: Core.IGetActionOptions, callback: Core.ILabCallback<Core.IAction[]>)`

Предпринимает попытку совершить действие.

 **Параметры**


|||
|:-----|:-----|
| _type_|Тип действия получения.|
| _options_|Параметры, предоставленные с помощью действия получения.|
| _callback_|Функция обратного вызова, которая возвращает список выполненных действий.|
