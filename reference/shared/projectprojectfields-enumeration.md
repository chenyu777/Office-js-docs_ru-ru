
# Перечисление ProjectProjectFields
Указывает поля проекта, доступные в качестве параметров для метода **[getProjectFieldAsync](../../reference/shared/projectdocument.getprojectfieldasync.md)**.

|||
|:-----|:-----|
|**Ведущие приложения:**|Project|
|**Добавлено в версии**|1.0|

```
ProjectProjectFields={
    CurrencyDigits: 0, 
    CurrencySymbol: 1, 
    CurrencySymbolPosition: 2, 
    DurationUnits: 3,
    GUID: 4, 
    Finish: 5, 
    Start: 6, 
    ReadOnly: 7, 
    VERSION: 8, 
    WorkUnits: 9, 
    ProjectServerUrl: 10, 
    WSSUrl: 11, 
    WSSList: 12
}
```


## Элементы


****


|**Элемент	**|**Описание**|
|:-----|:-----|
|**CurrencyDigits**|Количество десятичных знаков для валюты.|
|**CurrencySymbol**|Символ валюты.|
|**CurrencySymbolPosition**|Расположение знака валюты: "не указано" = -1; "перед значением без пробела ($0)" = 0; "после значения без пробела (0$)" = 1; "перед значением с пробелом ($ 0)" = 2; "после значения с пробелом (0 $)" = 3.|
|**GUID**|GUID проекта.|
|**Finish**|Дата завершения проекта.|
|**Начало**|Дата начала проекта.|
|**ReadOnly**|Указывает, доступен ли проект только для чтения.|
|**ВЕРСИЯ**|Версия проекта.|
|**WorkUnits**|Единицы трудозатрат проекта, например, дни или часы.|
|**ProjectServerUrl**|URL-адрес Project Web App для проектов, хранимых на сервере Project Server.|
|**WSSUrl**|URL-адрес SharePoint для проектов, синхронизированных со списком SharePoint.|
|**WSSList**|Имя списка SharePoint для проектов, синхронизированных со списком задач.|

## Заметки

Константу **ProjectProjectFields** можно использовать в качестве параметра метода **[getProjectFieldAsync](../../reference/shared/projectdocument.getprojectfieldasync.md)**.


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что данное перечисление поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает это перечисление.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Поддерживаемые ведущие приложения по платформе**


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**Типы надстроек**|Область задач|
|**Библиотека**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



****


|**Версия**|**Изменения**|
|:-----|:-----|
|1.0|Представлено|

## См. также



#### Другие ресурсы


[Метод getProjectFieldAsync](../../reference/shared/projectdocument.getprojectfieldasync.md)
