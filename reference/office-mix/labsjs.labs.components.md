
# LabsJS.Labs.Components
Предоставляет общий обзор интерфейса API JavaScript для Labs.Components Labs.JS.

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

API в модуле Labs.Components представляют четыре типа компонентов по умолчанию, которые в настоящее время доступны для разработки в лабораториях (компоненты действия (Activity), компоненты выбора (Choice), компоненты ввода (Input) и динамические компоненты (Dynamic)).

## Модуль Labs.Components

В таблице ниже приведены типы компонентов Labs.Components.


### Классы


|||
|:-----|:-----|
|[Labs.Components.ComponentAttempt](../../reference/office-mix/labs.components.componentattempt.md)|Базовый класс для попыток компонентов.|
|[Labs.Components.ActivityComponentAttempt](../../reference/office-mix/labs.components.activitycomponentattempt.md)|Представляет попытку завершения компонента действия.|
|[Labs.Components.ActivityComponentInstance](../../reference/office-mix/labs.components.activitycomponentinstance.md)|Представляет текущий экземпляр компонента действия.|
|[Labs.Components.ChoiceComponentAnswer](../../reference/office-mix/labs.components.choicecomponentanswer.md)|Ответ для проблемы, представленной в компоненте выбора.|
|[Labs.Components.ChoiceComponentAttempt](../../reference/office-mix/labs.components.choicecomponentattempt.md)|Представляет попытку компонента выбора.|
|[Labs.Components.ChoiceComponentInstance](../../reference/office-mix/labs.components.choicecomponentinstance.md)|Представляет экземпляр компонента выбора.|
|[Labs.Components.ChoiceComponentResult](../../reference/office-mix/labs.components.choicecomponentresult.md)|Результат отправки компонента выбора.|
|[Labs.Components.ChoiceComponentSubmission](../../reference/office-mix/labs.components.choicecomponentsubmission.md)|Представляет отправку, связанную с компонентом выбора.|
|[Labs.Components.DynamicComponentInstance](../../reference/office-mix/labs.components.dynamiccomponentinstance.md)|Представляет экземпляр динамического компонента.|
|[Labs.Components.InputComponentAnswer](../../reference/office-mix/labs.components.inputcomponentanswer.md)|Представляет ответ для проблемы компонента ввода.|
|[Labs.Components.InputComponentAttempt](../../reference/office-mix/labs.components.inputcomponentattempt.md)|Представляет попытку взаимодействия с компонентом ввода.|
|[Labs.Components.InputComponentInstance](../../reference/office-mix/labs.components.inputcomponentinstance.md)|Представляет экземпляр компонента ввода.|
|[Labs.Components.InputComponentResult](../../reference/office-mix/labs.components.inputcomponentresult.md)|Результат отправки компонента ввода.|
|[Labs.Components.InputComponentSubmission](../../reference/office-mix/labs.components.inputcomponentsubmission.md)|Представляет отправку в компонента ввода.|

### Интерфейсы


|||
|:-----|:-----|
|[Labs.Components.IActivityComponent](../../reference/office-mix/labs.components.iactivitycomponent.md)|Представляет компонент действия. Расширяет [Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md).|
|[Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md)|Представляет определенный экземпляр компонента действия. Расширяет [Labs.Core.IComponentInstance](../../reference/office-mix/labs.core.icomponentinstance.md).|
|[Labs.Components.IChoice](../../reference/office-mix/labs.components.ichoice.md)|Доступные варианты для определенной проблемы.|
|[Labs.Components.IChoiceComponent](../../reference/office-mix/labs.components.ichoicecomponent.md)|Позволяет взаимодействовать с компонентом выбора.|
|[Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md)|Экземпляр компонента выбора.|
|[Labs.Components.IDynamicComponent](../../reference/office-mix/labs.components.idynamiccomponent.md)|Позволяет взаимодействовать с динамическим компонентом.|
|[Labs.Components.IDynamicComponentInstance](../../reference/office-mix/labs.components.idynamiccomponentinstance.md)|Экземпляр динамического компонента.|
|[Labs.Components.IHint](../../reference/office-mix/labs.components.ihint.md)|Подсказка для проблемы лаборатории.|
|[Labs.Components.IInputComponent](../../reference/office-mix/labs.components.iinputcomponent.md)|Позволяет взаимодействовать с компонентом ввода.|
|[Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md)|Экземпляр компонента ввода.|
