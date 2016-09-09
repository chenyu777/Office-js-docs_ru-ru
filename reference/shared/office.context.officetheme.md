
# Свойство Context.officeTheme
Предоставляет доступ к свойствам цветов темы Office.

 **Важно!** В настоящее время этот API работает только в Excel, Outlook, PowerPoint и Word из набора [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview) для настольных компьютеров с Windows.


|||
|:-----|:-----|
|**Ведущие приложения:**|Excel, Outlook, PowerPoint, Word|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Не в наборе|
|**Добавлено в версии**|1.3|



```js
Office.context.officeTheme
```


## Элементы


**Свойства**

|||
|:-----|:-----|
|Имя|Описание|
|[bodyBackgroundColor ](../../reference/shared/office.context.bodybackgroundcolor.md)|Получает цвет фона текста сообщения для темы Office.|
|[bodyForegroundColor](../../reference/shared/office.context.bodyforegroundcolor.md)|Получает цвет переднего плана для темы Office.|
|[controlBackgroundColor](../../reference/shared/office.context.controlbackgroundcolor.md)|Получает цвет фона элемента управления для темы Office.|
|[controlForegroundColor](../../reference/shared/office.context.controlforegroundcolor.md)|Получает цвет переднего плана элемента управления для темы Office.|

## Замечания

Цвета тем Office позволяют согласовать цветовую схему надстройки с текущей темой Office, которую пользователь выбрал с помощью команды **Файл**  >  **Учетная запись Office**  >  **Office Theme UI** и которая применяется во всех ведущих приложениях Office. Цвета тем Office можно использовать для всех надстроек Outlook и области задач.


## Пример


```js
function applyOfficeTheme(){
    // Get office theme colors.
    var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
    var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
    var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
    var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

    // Apply body background color to a CSS class.
    $('.body').css('background-color', bodyBackgroundColor);
}
```


## Сведения о поддержке



|||
|:-----|:-----|
|**Минимальный уровень разрешений**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Надстройки области задач, надстройки Outlook, контентные надстройки|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки


|**Версия**|**Изменения**|
|:-----|:-----|
|1.3|Представлено|
