
# Перечисление ProjectViewTypes
Указывает типы представлений, которые может распознать метод **[getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)**.

|||
|:-----|:-----|
|**Ведущие приложения:**|Project|
|**Добавлено в версии**|1.0|

```
ProjectViewTypes={
    Gantt           : 1, 
    NetworkDiagram  : 2, 
    TaskDiagram     : 3, 
    TaskForm        : 4, 
    TaskSheet       : 5, 
    ResourceForm    : 6, 
    ResourceSheet   : 7, 
    ResourceGraph   : 8, 
    TeamPlanner     : 9, 
    TaskDetails     : 10, 
    TaskNameForm    : 11, 
    ResourceNames   : 12, 
    Calendar        : 13, 
    TaskUsage       : 14, 
    ResourceUsage   : 15, 
    Timeline        : 16
}
```


## Элементы


****


|**Элемент	**|**Описание**|
|:-----|:-----|
|**Gantt**|Представление "Диаграмма Ганта".|
|**NetworkDiagram**|Представление "Сетевой график".|
|**TaskDiagram**|Представление "Схема задач".|
|**TaskForm**|Представление "Форма задач".|
|**TaskSheet**|Представление "Лист задач".|
|**ResourceForm**|Представление "Форма ресурсов".|
|**ResourceSheet**|Представление "Лист ресурсов".|
|**ResourceForm**|Представление "Форма ресурсов".|
|**ResourceGraph**|Представление "График ресурсов".|
|**TeamPlanner**|Представление "Планировщик работы группы".|
|**TaskDetails**|Представление "Сведения о задаче".|
|**TaskNameForm**|Представление "Форма названий задач".|
|**ResourceNames**|Представление "Имена ресурсов".|
|**Календарь**|Представление "Календарь".|
|**TaskUsage**|Представление "Использование задач".|
|**ResourceUsage**|Представление "Использование ресурсов".|
|**Временная шкала**|Представление "Временная шкала".|

## Заметки

Метод **[getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)** возвращает значение константы **ProjectViewTypes** и имя, соответствующее активному представлению.


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


[Метод getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)
