
# Перечисление InitializationReason
Указывает, была ли надстройка вставлена только что или уже была частью документа. 

|||
|:-----|:-----|
|**Ведущие приложения:**|Excel, Project, Word|
|**Добавлено в версии**|1.0|

```
Office.InitializationReason
```


## Элементы


**Значения**


|**Перечисление**|**Значение**|**Описание**|
|:-----|:-----|:-----|
|Office.InitializationReason.Inserted|"inserted"|Надстройка была вставлена в документ.|
|Office.InitializationReason.DocumentOpened|"documentOpened"|Надстройка уже содержалась в документе, который был открыт.|

## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что данное перечисление поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает это перечисление.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Поддерживаемые ведущие приложения по платформе**


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Да|Y|
|**Project**|Y|||
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки




|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel и Word в Office для iPad.|
|1.0|Представлено|
