
# Свойство Slice.data
Получает необработанные данные фрагмента файла.

|||
|:-----|:-----|
|**Ведущие приложения:**|PowerPoint, Word|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Файл|
|**Последнее изменение в **|1.1|

```
var sliceData = slice.data;
```


## Возвращаемое значение

Необработанные данные фрагмента файла в формате **Office.FileType.Text** ("text") или **Office.FileType.Compressed** ("compressed"), как указано параметром _fileType_ вызова метода [Document.getFileAsync](../../reference/shared/document.getfileasync.md).


## Заметки

Файлы в "сжатом" формате будут возвращать массив байтов, который может быть преобразован в строку в кодировке Base 64 (при необходимости).


## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что это свойство поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает это свойство.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Да|Y|
|**Word**|Y|Да|Y|


|||
|:-----|:-----|
|**Доступен в наборах требований**|Файл|
|**Минимальный уровень разрешений**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



****


|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка PowerPoint и Word в Office для iPad.|
|1.0|Представлено|