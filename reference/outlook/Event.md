

# Событие

Объект `event` передается как параметр для функций надстройки, вызванных кнопками в режиме без пользовательского интерфейса. Этот объект позволяет надстройке определить нажатую кнопку и уведомить узел о завершении обработки.

Например, рассмотрим кнопку, определенную в манифесте надстройки следующим образом:

```
<Control xsi:type="Button" id="eventTestButton">
  <Label resid="eventButtonLabel" />
  <Tooltip resid="eventButtonTooltip" />
  <Supertip>
    <Title resid="eventSuperTipTitle" />
    <Description resid="eventSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>testEventObject</FunctionName>
  </Action>
</Control>
```

Для атрибута `id` кнопки задано значение `eventTestButton`, что приводит к вызову функции `testEventObject`, определенной в надстройке. Эта функция выглядит следующим образом:

```
function testEventObject(event) {
  // The event object implements the Event interface

  // This value will be "eventTestButton"
  var buttonId = event.source.id;

  // Signal to the host app that processing is complete.
  event.completed();
}
```

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.3|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Применимый режим Outlook| Создание или чтение|

### Элементы

####  source :Object

Получает идентификатор кнопки надстройки, с помощью которой вызван метод.

Свойство `source` возвращает объект с указанными ниже свойствами.

| Свойство | Описание |
| --- | --- |
| `id` | Значение атрибута `id` элемента `Control`, который определяет кнопку надстройки в манифесте надстройки. |

Это значение можно использовать, если для вызова одной функции используется несколько кнопок, однако в зависимости от нажатой кнопки вам потребуется выполнить различные действия.

##### Тип:

*   Объект

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.3|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```
// Function is used by two buttons:
// button1 and button2
function multiButton (event) {
  // Check which button was clicked
  var buttonId = event.source.id;

  if (buttonId === 'button1') {
    doButton1Action();
  else {
    doButton2Action();
  }

  event.completed();
}
```

### Методы

####  completed()

Указывает, что надстройка завершила обработку, активированную с помощью кнопки надстройки.

Этот метод необходимо вызывать в конце функции, которую вызвали с помощью команды надстройки, определенной с использованием элемента `Action`, для атрибута `xsi:type` которого задано значение `ExecuteFunction`. При вызове этого метода клиент узла узнает, что функция завершена и что он может очистить любое состояние, связанное с вызовом функции. Например, если пользователь закроет Outlook до вызова этого метода, Outlook предупредит его, что функция по-прежнему выполняется.

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.3|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Применимый режим Outlook| Создание или чтение|

##### Пример

```
function processItem (event) {
  // Do some processing

  event.completed();
}
```