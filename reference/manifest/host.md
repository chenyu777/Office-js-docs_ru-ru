
# Элемент Host
Указывает тип ведущего приложения Office, поддерживаемый надстройкой Office.

 **Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.


## Синтаксис:


```XML
<Host Name= ["Document" | "Database" | "Mailbox" | "Presentation" | "Project" | "Workbook"] />
```


## Атрибуты



|**Атрибут**|**Тип**|**Обязательный**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя|string|Обязательный|Имя типа ведущего приложения Office.|

## Заметки

В атрибуте **Name** элемента **Host** вы можете указать приведенные ниже значения. Каждое значение сопоставлено с набором из одного или нескольких ведущих приложений Office, поддерживаемых вашей надстройкой.



|**Имя**|**Ведущие приложения Office**|
|:-----|:-----|
| `"Document"`|Word, Word Online, Word на iPad|
| `"Database"`|Веб-приложения Access|
| `"Mailbox"`|Outlook, Outlook Web App, Outlook Web App для устройств|
| `"Notebook"`|OneNote Online|
| `"Presentation"`|PowerPoint, PowerPoint Online, PowerPoint на iPad|
| `"Project"`|Project|
| `"Workbook"`|Excel, Excel Online, Excel на iPad|

## Заметки

Дополнительные сведения об указании поддерживаемых ведущих приложений см. в статье [Указание ведущих приложений Office и требований к API](../../docs/overview/specify-office-hosts-and-api-requirements.md).

