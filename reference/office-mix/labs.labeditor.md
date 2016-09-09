
# Labs.LabEditor

 _**Область применения**: приложения для Office | надстройки Office | Office Mix | PowerPoint_

С помощью объекта **LabEditor** можно изменять определенную лабораторию, а также получать и задавать данные конфигурации, связанные с этой лабораторией.

```
class LabEditor
```


## Методы


### getConfiguration

 `public function getConfiguration(callback: Labs.Core.ILabCallback<Labs.Core.IConfiguration>): void`

Получает текущую конфигурацию лаборатории.

 **Параметры**


|**Имя**|**Описание**|
|:-----|:-----|
| _callback_|Функция обратного вызова, которая срабатывает после получения конфигурации.|

### setConfiguration

 `public function getConfiguration(callback: Labs.Core.ILabCallback<Labs.Core.IConfiguration>): void`

Задает новую конфигурацию лаборатории.

 **Параметры**


|**Имя**|**Описание**|
|:-----|:-----|
| _configuration_|Конфигурация, которую требуется задать.|
| _callback_|Функция обратного вызова, которая срабатывает, когда конфигурация задана.|

### done

 `public function done(callback: Labs.Core.ILabCallback<void>): void`

Указывает на то, что пользователь завершил редактирование лаборатории.

 **Параметры**


|**Имя**|**Описание**|
|:-----|:-----|
| _callback_|Функция обратного вызова, которая срабатывает, когда работа редактора лаборатории завершена.|
