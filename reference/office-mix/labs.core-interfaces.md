
# Интерфейсы Labs.Core
Интерфейсы в модуле **LabsJS.Labs.Core**.

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Модуль **LabsJS.Labs.Core** содержит указанные ниже интерфейсы.

## 


|||
|:-----|:-----|
|[Labs.Core.IAction](../../reference/office-mix/labs.core.iaction.md)|Представляет действие лаборатории, то есть взаимодействие между пользователем и указанной лабораторией.|
|[Labs.Core.IActionResult](../../reference/office-mix/labs.core.iactionresult.md)|Результаты выполнения действия. В зависимости от типа действия результаты либо задаются сервером, либо предоставляются клиентом при выполнении действия.|
|[Labs.Core.IComponentInstance](../../reference/office-mix/labs.core.icomponentinstance.md)|Базовый класс для экземпляров компонентов лаборатории.|
|[Labs.Core.IConfigurationInfo](../../reference/office-mix/labs.core.iconfigurationinfo.md)|Сведения о конфигурации лаборатории.|
|[Labs.Core.IConnectionResponse](../../reference/office-mix/labs.core.iconnectionresponse.md)|Сведения об ответе, возвращаемые при вызове подключения.|
|[Labs.Core.IGetActionOptions](../../reference/office-mix/labs.core.igetactionoptions.md)|Параметры, передаваемые в рамках действия **get**.|
|[Labs.Core.ILabCreationOptions](../../reference/office-mix/labs.core.ilabcreationoptions.md)|Параметры, передаваемые в рамках операции создания лаборатории.|
|[Labs.Core.ILabHostVersionInfo](../../reference/office-mix/labs.core.ilabhostversioninfo.md)|Сведения о версии узла лаборатории.|
|[Labs.Core.IActionOptions](../../reference/office-mix/labs.core.iactionoptions.md)|Определение параметров действий лаборатории. Параметры, которые передаются при выполнении данного действия.|
|[Labs.Core.IUserInfo](../../reference/office-mix/labs.core.iuserinfo.md)|Предоставляет сведения пользователя, относящиеся к лаборатории.|
|[Labs.Core.IValueInstance](../../reference/office-mix/labs.core.ivalueinstance.md)|Экземпляр объекта [Labs.Core.IValue](../../reference/office-mix/labs.core.ivalue.md), который содержит данные значения, если они есть.|
|[Labs.Core.IVersion](../../reference/office-mix/labs.core.iversion.md)|Предоставляет сведения о версии лаборатории.|
|[Labs.Core.IAnalyticsConfiguration](../../reference/office-mix/labs.core.ianalyticsconfiguration.md)|Сведения о конфигурации пользовательской аналитики. Позволяет указать, какой именно IFrame требуется загрузить для отображения пользовательской аналитики при запуске лаборатории пользователем.|
|[Labs.Core.ICompletionStatus](../../reference/office-mix/labs.core.icompletionstatus.md)|Состояние завершения для лаборатории. Состояние передается при завершении работы лаборатории для отображения результата взаимодействия.|
|[Labs.Core.ILabCallback](../../reference/office-mix/labs.core.ilabcallback.md)|Интерфейс для обработки методов обратного вызова Labs.js.|
|[Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md)|Объект, связанный с лабораторией. Объект содержит поле type, указывающее, каков тип объекта.|
|[Labs.Core.ITimelineConfiguration](../../reference/office-mix/labs.core.itimelineconfiguration.md)|Параметры конфигурации для [Labs.Timeline](../../reference/office-mix/labs.timeline.md). Позволяет задать ряд параметров конфигурации временной шкалы.|
|[Labs.Core.IUserData](../../reference/office-mix/labs.core.iuserdata.md)|Базовый интерфейс для представления пользовательских данных, хранящихся в объекте.|
|[Labs.Core.IValue](../../reference/office-mix/labs.core.ivalue.md)|Базовый класс для значений, хранящихся в лаборатории.|
|[Labs.Core.IConfiguration](../../reference/office-mix/labs.core.iconfiguration.md)|Структура данных конфигурации лаборатории.|
|[Labs.Core.IConfigurationInstance](../../reference/office-mix/labs.core.iconfigurationinstance.md)|Базовый класс для экземпляров конфигурации лаборатории.|
|[Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md)|Базовый класс для представления компонентов лаборатории.|
|[Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md)|Предоставляет уровень абстракции для подключения Labs.js к ведущему приложению.|
|[Labs.Core.ModeChangedEventData](../../reference/office-mix/labs.core.modechangedeventdata.md)|Данные, связанные с событием смены режима.|
|[Labs.Core.IEventCallback](../../reference/office-mix/labs.core.ieventcallback.md)|Интерфейс для обработки обратных вызовов EventManager.|
