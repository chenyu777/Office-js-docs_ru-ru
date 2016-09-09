
# Объект AsyncResult
Объект, который инкапсулирует результат асинхронного запроса, включая сведения о состоянии и ошибке, если запрос завершился ошибкой.

|||
|:-----|:-----|
|**Ведущие приложения:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Последнее изменение в **|1.1|

```
AsyncResult
```


## Элементы


**Свойства**


|**Имя**|**Описание**|
|:-----|:-----|
|**[asyncContext](../../reference/shared/asyncresult.asynccontext.md)**|Получает определяемый пользователем элемент, передаваемый необязательному параметру _asyncContext_ вызванного метода в том состоянии, в каком был передан.|
|**[error](../../reference/shared/asyncresult.error.md)**|Получает объект **Error** с описанием ошибки, если таковая возникает.|
|**[status](../../reference/shared/asyncresult.status.md)**|Получает состояние асинхронной операции.|
|**[value](../../reference/shared/asyncresult.value.md)**|Получает полезные данные или содержимое асинхронной операции (если имеется).|

## Заметки

Когда выполняется функция, переданная в параметр _callback_ в метод "Async", она получает объект [AsyncResult](../../reference/shared/asyncresult.md), к которому можно получить доступ исключительно с помощью параметра функции обратного вызова.

Ниже приведен пример, применимый к контентным надстройкам и надстройкам области задач. В примере показан вызов метода [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) для объекта **Document**.




```js
Office.context.document.getSelectedDataAsync("text", {valueFormat:"unformatted", filterType:"all"}, 
   function (result) {
      if (result.status === "success")      
         var dataValue = result.value; // Get selected data.
         write('Selected data is ' + dataValue);
      else {            
         var err = result.error; 
         write(err.name + ": " + err.message);
      }
   });
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

Анонимная функция, переданная в качестве аргумента _callback_ (`function (result){...}`), имеет один параметр с именем _result_, который предоставляет доступ к объекту **AsyncResult** при выполнении функции. После завершения вызова метода **getSelectedDataAsync** выполняется функция обратного вызова, и следующая строка кода обращается к свойству **value** объекта **AsyncResult**, чтобы вернуть данные, выбранные в документе.

 `var dataValue = result.value;`

Обратите внимание, что в других строках кода функции используется параметр _result_ функции обратного вызова для доступа к свойствам **status** и **error** объекта **AsyncResult**.

Объект **AsyncResult** доступен из функции, переданной в качестве аргумента в параметре _callback_ следующих методов.



|**Родительский объект**|**Способ**|
|:-----|:-----|
|**Document** (только для Excel, PowerPoint, Project и Word)|[getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md)|
||[setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md)|
|**Bindings** (только для Excel и Word)|[addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md)|
||[addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md)|
||[getAllAsync](../../reference/shared/bindings.getallasync.md)|
||[getByIdAsync](../../reference/shared/bindings.getbyidasync.md)|
||[releaseByIdAsync](../../reference/shared/bindings.releasebyidasync.md)|
|**Binding** (только для Excel и Word)|[getDataAsync](../../reference/shared/binding.getdataasync.md)|
||[setDataAsync](../../reference/shared/binding.setdataasync.md)|
||[removeHandlerAsync](../../reference/shared/binding.removehandlerasync.md)|
|**TableBinding** (только для Excel и Word)||
||[addRowsAsync](../../reference/shared/binding.tablebinding.addrowsasync.md)|
||[deleteAllDataValuesAsync](../../reference/shared/binding.tablebinding.deletealldatavaluesasync.md)|
|**Settings** (только для Excel, PowerPoint и Word)|[refreshAsync](../../reference/shared/settings.refreshasync.md)|
||[saveAsync](../../reference/shared/settings.saveasync.md)|
|**CustomXmlNode** (только для Word)|[getNodesAsync](../../reference/shared/customxmlnode.getnodesasync.md)|
||[getNodeValueAsync](../../reference/shared/customxmlnode.getnodevalueasync.md)|
||[getXmlAsync](../../reference/shared/customxmlnode.getxmlasync.md)|
||[setNodeValueAsync](../../reference/shared/customxmlnode.setnodevalueasync.md)|
||[setXmlAsync](../../reference/shared/customxmlnode.setxmlasync.md)|
|**CustomXmlPart** (только для Word)|[deleteAsync](../../reference/shared/customxmlpart.deleteasync.md)|
||[getNodesAsync](../../reference/shared/customxmlpart.getnodesasync.md)|
||[getXmlAsync](../../reference/shared/customxmlpart.getxmlasync.md)|
|**CustomXmlParts** (только для Word)|[addAsync](../../reference/shared/customxmlparts.addasync.md)|
||[getByIdAsync](../../reference/shared/customxmlparts.getbyidasync.md)|
||[getByNamespaceAsync](../../reference/shared/customxmlparts.getbynamespaceasync.md)|
|**CustomXmlPrefixMappings** (только для Word)|[addNamespaceAsync](../../reference/shared/customxmlprefixmappings.addnamespaceasync.md)|
||[getNamespaceAsync](../../reference/shared/customxmlprefixmappings.getnamespaceasync.md)|
||[getPrefixAsync](../../reference/shared/customxmlprefixmappings.getprefixasync.md)|
|**Mailbox** (только для Outlook)|[getUserIdentityTokenAsync](http://msdn.microsoft.com/library/c658518b-6867-41a0-99cf-810303e4c539%28Office.15%29.aspx)|
||[makeEwsRequestAsync](http://msdn.microsoft.com/library/2ec380e0-4a67-4146-92a6-6a39f65dc6f2%28Office.15%29.aspx)|
|**CustomProperties** (только для Outlook)|[saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56%28Office.15%29.aspx)|
|**Item** (только для Outlook)|[loadCustomPropertiesAsync](http://msdn.microsoft.com/library/dfbec151-8ea7-4915-b723-09ea1396a261%28Office.15%29.aspx)|
|**RoamingSettings** (только для Outlook)|[saveAsync](http://msdn.microsoft.com/library/a616f71c-a447-423f-a0d2-e9d6f1ac32f8%28Office.15%29.aspx)|

## Сведения о поддержке


Заглавная буква Y в следующей матрице указывает на то, что этот метод поддерживается в соответствующем ведущем приложении Office. Пустая ячейка означает, что ведущее приложение Office не поддерживает этот метод.

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).



| |**Office для рабочего стола Windows**|**Office Online (в браузере)**|**Office для iPad**|**OWA для устройств**|**Outlook для Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Да|Y|||
|**Outlook**|Y|Да||Да|Y|
|**PowerPoint**|Y|Да|Y|||
|**Project**|Y|||||
|**Word**|Y|Да|Y|||

|||
|:-----|:-----|
|**Типы надстроек**|Надстройки области задач, надстройки Outlook, контентные надстройки|
|**Library**|Office.js|
|**Пространство имен**|Office|

## Журнал поддержки



****


|**Версия**|**Изменения**|
|:-----|:-----|
|1.1|Добавлена поддержка Excel, PowerPoint и Word в Office для iPad.|
|1.1|Добавлена поддержка надстроек для Access.|
|1.0|Представлено|
