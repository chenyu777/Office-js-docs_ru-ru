
# LabsJS.Labs

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Модуль LabsJS.Labs содержит набор ключевых API JavaScript, позволяющих создавать надстройки Office (лаборатории). Эти API обеспечивают точку входа для разработки лабораторий.

## Модуль API LabsJS.Labs

Модуль Labs лабораторий содержит элементы следующих типов.


### Переменные


|||
|:-----|:-----|
|[Labs.DefaultHostBuilder](../../reference/office-mix/labs.defaulthostbuilder.md)|Используйте этот объект для создания экземпляра [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md) по умолчанию.|

### Функции


|||
|:-----|:-----|
|[Labs.Connect](../../reference/office-mix/labs.connect.md)|Инициализирует соединение с ведущим приложением.|
|[Labs.connect (overload)](../../reference/office-mix/labs.connect-overload.md)|Инициализирует соединение с ведущим приложением и предоставляет входные параметры.|
|[Labs.isConnected](../../reference/office-mix/labs.isconnected.md)|Инициализирует соединение с ведущим приложением.|
|[Labs.getConnectionInfo](../../reference/office-mix/labs.getconnectioninfo.md)|Получает сведения о конфигурации, связанные с указанным соединением.|
|[Labs.disconnect](../../reference/office-mix/labs.disconnect.md)|Разрывает соединение между лабораторией и ведущим приложением и предоставляет состояние завершения действий с лабораторией.|
|[Labs.editLab](../../reference/office-mix/labs.editlab.md)|Открывает указанную лабораторию для изменения. В режиме редактирования можно указать данные конфигурации лаборатории. Но изменить лабораторию, которая выполняется (то есть запущена), невозможно.|
|[Labs.takeLab](../../reference/office-mix/labs.takelab.md)|Запускает указанную лабораторию и позволяет отправлять результаты работы в лаборатории на сервер. Обратите внимание, что невозможно запустить лабораторию во время ее изменения.|
|[Labs.on](../../reference/office-mix/labs.on.md)|Добавляет новый обработчик для указанного события.|
|[Labs.off](../../reference/office-mix/labs.off.md)|Удаляет обработчик для указанного события.|
|[Labs.getTimeline](../../reference/office-mix/labs.gettimeline.md)|Получает экземпляр объекта [Labs.Timeline](../../reference/office-mix/labs.timeline.md), который позволяет контролировать элемент управления проигрывателем ведущего приложения.|
|[Labs.registerDeserializer](../../reference/office-mix/labs.registerdeserializer.md)|Выполняет десериализацию указанного объекта JSON, преобразуя его в обычный объект. Эту функцию должны использовать только авторы компонентов.|

### Классы


|||
|:-----|:-----|
|[Labs.ComponentInstanceBase](../../reference/office-mix/labs.componentinstancebase.md)|Базовый класс для инициализации экземпляров компонентов.|
|[Labs.ComponentInstance](../../reference/office-mix/labs.componentinstance.md)|Представляет экземпляр определенного компонента, созданный для пользователя в среде выполнения. Объект содержит преобразованное представление компонента для определенного запуска лаборатории.|
|[Labs.Command](../../reference/office-mix/labs.command.md)|Стандартная команда, используемая для передачи сообщений между клиентом и ведущим приложением.|
|[Labs.LabEditor](../../reference/office-mix/labs.labeditor.md)|Объект **LabEditor** позволяет изменять определенную лабораторию, а также получать и задавать данные конфигурации, связанные с этой лабораторией.|
|[Labs.LabInstance](../../reference/office-mix/labs.labinstance.md)|Экземпляр лаборатории, настроенный для текущего пользователя. Используйте этот объект, чтобы записывать и извлекать данные лаборатории для пользователя.|
|[Labs.Timeline](../../reference/office-mix/labs.timeline.md)|Предоставляет доступ к функции временной шкалы labs.js.|
|[Labs.ValueHolder](../../reference/office-mix/labs.valueholder.md)|Объект-контейнер, который содержит и отслеживает значения для определенной лаборатории. Значение может храниться локально или на сервере.|

### Интерфейсы


|||
|:-----|:-----|
|[Labs.GetActionsCommandData](../../reference/office-mix/labs.getactionscommanddata.md)|Позволяет получить данные, связанные с командой [LabsJS.Labs.Core.GetActions](../../reference/office-mix/labsjs.labs.core.getactions.md).|
|[Labs.IMessageHandler](../../reference/office-mix/labs.imessagehandler.md)|Интерфейс, позволяющий определять обработчики событий.|
|[Labs.ITimelineNextMessage](../../reference/office-mix/labs.itimelinenextmessage.md)|Предоставляет средства для взаимодействия с объектом [Labs.Core.IMessage](https://msdn.microsoft.com/library/office/mt599680.aspx).|
|[Labs.SendMessageCommandData](../../reference/office-mix/labs.sendmessagecommanddata.md)|Данные, связанные с командой [Labs.CommandType.TakeAction](https://msdn.microsoft.com/library/office/mt599680.aspx).|
|[Labs.TakeActionCommandData](../../reference/office-mix/labs.takeactioncommanddata.md)|Данные, связанные с командой выполнения действия.|

### Перечисления


|||
|:-----|:-----|
|[Labs.ConnectionState](../../reference/office-mix/labs.connectionstate.md)|Перечисляет возможные состояния соединения лаборатории с ведущим приложением.|
|[Labs.ProblemState](../../reference/office-mix/labs.problemstate.md)|Значения состояния для заданной лаборатории.|
