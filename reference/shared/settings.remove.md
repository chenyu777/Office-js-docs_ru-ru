

# Метод Settings.remove
Удаляет указанный параметр.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, PowerPoint, Word|
|**Доступен в [наборах требований](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Параметры|
|**Последнее изменение в **|1.1|

```js
Office.context.document.settings.remove(name);
```


## Параметры



_name_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Тип: **string**

&nbsp;&nbsp;&nbsp;&nbsp;Имя удаляемого параметра с учетом регистра.

    



## Замечания

 **NULL** является допустимым значением параметра, поэтому присвоение значения **NULL** не удалит параметр из контейнера свойств параметров.


 >**Важно!** Учтите, что метод **Settings.remove** влияет только на копию контейнера свойств параметров, содержащуюся в памяти. Чтобы предотвратить удаление указанного параметра в документе, в какой-либо точке после вызова метода **Settings.remove** и до закрытия приложения необходимо вызвать метод [Settings.saveAsync](../../reference/shared/settings.saveasync.md).


## Пример




```js
function removeMySetting() {
    Office.context.document.settings.remove('mySetting');
}
```




## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).



||**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Да|Y|
|**PowerPoint**|Y|Да|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Доступен в наборах требований**|Параметры|
|**Минимальный уровень разрешений**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Типы надстроек**|Контентные надстройки и надстройки области задач|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки




|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка PowerPoint Online.|
|1.1|Добавлена поддержка Excel, PowerPoint и Word в Office для iPad.|
|1.1|Добавлена поддержка создания настраиваемых параметров в контентных надстройках для Access.|
|1.0|Представлено|