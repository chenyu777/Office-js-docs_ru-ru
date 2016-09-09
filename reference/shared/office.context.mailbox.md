
# Свойство Context.mailbox
Получает объект **mailbox**, который предоставляет доступ к элементам API специально для надстроек Outlook.

|||
|:-----|:-----|
|**Ведущие приложения:**|Outlook|
|**Доступно в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Почтовый ящик|
|**Последнее изменение в **|1.0|

```js
var outlookOm = Office.context.mailbox;
```


## Возвращаемое значение

Объект [mailbox](http://msdn.microsoft.com/library/a3880d3b-8a09-4cf9-9274-f2682cb3b769%28Office.15%29.aspx).


## Пример

В следующем примере кода показано обращение к объекту [item](http://msdn.microsoft.com/library/ad288df1-3ca2-474c-bea4-c51f46e6fc43%28Office.15%29.aspx) API JavaScript для Office.


```js
// Access the Item object.
var item = Office.context.mailbox.item;

```




## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Outlook для Mac**|
|:-----|:-----|:-----|:-----|
|**Outlook**|Y|Да|Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|Почтовый ящик|
|**Минимальный уровень разрешений**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Outlook|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки


|**Версия**|**Изменения**|
|:-----|:-----|
|1.0|Представлено|
