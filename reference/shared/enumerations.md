
# Перечисления

Перечислимые значения можно указывать с помощью либо полного имени перечисления (`Office.CoercionType.Text`), либо его соответствующего текстового значения (`"text"`). Например, в следующем вызове метода используются имена перечислений:


```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, {valueFormat:Office.ValueFormat.Unformatted, filterType:Office.FilterType.All},
   function (result) {
      if (result.status === Office.AsyncResultStatus.Success)
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


А здесь в том же самом вызове используются текстовые значения перечислений:




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
```


## Справочные материалы



|**Название**|**Определение**|
|:-----|:-----|
|[ActiveView](activeview-enumeration.md)|Указывает состояние активного представления документа, например возможность редактирования документа пользователем.|
|[AsyncResultStatus](asyncresultstatus-enumeration.md)|Указывает результат асинхронного вызова.|
|[AttachmentType](http://msdn.microsoft.com/library/83883a47-a937-4afb-a55e-e789057335c4%28Office.15%29.aspx)|Указывает тип вложения почтового сообщения или приглашения на собрание. Outlook 2013 не поддерживает это перечисление.|
|[BindingType](bindingtype-enumeration.md)|Указывает тип объекта привязки, который нужно вернуть.|
|[BodyType](http://msdn.microsoft.com/library/31350fe6-4c42-4cbb-a5b2-4fb2d360fa11%28Office.15%29.aspx)|Указывает тип основного текста встречи или сообщения.|
|[CoercionType](coerciontype-enumeration.md)|Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.|
|[CustomXMLNodeType](customxmlnodetype-enumeration.md)|Указывает тип узла.|
|[DocumentMode](documentmode-enumeration.md)|Указывает атрибут документа в соответствующем приложении: только чтение или чтение и запись. |
|[EntityType](http://msdn.microsoft.com/library/0035be38-8a65-4693-bcc4-0a8dd7b1495b%28Office.15%29.aspx)|Указывает тип сущности.|
|[EventType](eventtype-enumeration.md)|Указывает тип вызванного события.|
|[FileType](filetype-enumeration.md)|Указывает формат, в котором возвращается документ.|
|[GoToType](gototype-enumeration.md)|Указывает тип места или объекта для перехода.|
|[FilterType](filtertype-enumeration.md)|Указывает, применяется ли фильтрация из ведущего приложения при извлечении данных.|
|[InitializationReason](initializationreason-enumeration.md)|Указывает, была ли надстройка вставлена только что или уже была частью документа.|
|[ItemType](http://msdn.microsoft.com/library/e0bb23fd-f360-4b0f-b72c-1cf08d4cab3f%28Office.15%29.aspx)|Указывает тип элемента.|
|[notificationMessageType](http://msdn.microsoft.com/library/ff00c89d-0019-4545-a95b-7ed0db712ce9%28Office.15%29.aspx)|Указывает уведомление для встречи или сообщения.|
|[ProjectProjectFields](projectprojectfields-enumeration.md)|Указывает поля проекта, доступные в качестве параметров для метода [getProjectFieldAsync](projectdocument.getprojectfieldasync.md).|
|[ProjectResourceFields](projectresourcefields-enumeration.md)|Задает поля ресурса, доступные в качестве параметра для метода [getResourceFieldAsync](projectdocument.gettaskfieldasync.md).|
|[ProjectTaskFields](projecttaskfields-enumeration.md)|Задает поля задачи, доступные в качестве параметров для метода [getTaskFieldAsync](projectdocument.gettaskfieldasync.md).|
|[ProjectViewTypes](projectviewtypes-enumeration.md)|Указывает типы представлений, которые может распознать метод [getSelectedViewAsync](projectdocument.getselectedviewasync.md).|
|[RecipientType](http://msdn.microsoft.com/library/6e7c4029-6e52-47f6-98d2-4cd3ce7bd8b4%28Office.15%29.aspx)|Указывает тип получателя для встречи.|
|[ResponseType](http://msdn.microsoft.com/library/b3e723ca-4be0-4846-ad97-0eecab4355eb%28Office.15%29.aspx)|Указывает ответ на приглашение на собрание.|
|[SelectionMode](selectionmode-enumeration.md)|Указывает, следует ли выделять расположение для перехода (при использовании метода [Document.goToByIdAsync](document.gotobyidasync.md)).|
|[SourceProperty](http://msdn.microsoft.com/library/6a209a7f-57cd-4dc3-869e-07b0f5928b28%28Office.15%29.aspx)|Указывает источник данных, возвращаемых вызванным методом.|
|[Таблица](table-enumeration.md)|Указывает перечисляемые значения для свойства `cells:` в параметре _cellFormat_[методов форматирования таблиц](../../docs/excel/format-tables-in-add-ins-for-excel.md).|
|[ValueFormat](valueformat-enumeration.md)|Указывает, форматируются ли значения, такие как числа и даты, возвращаемые вызванным методом.|

## Сведения о поддержке


Поддержка каждого перечисления зависит от ведущего приложения Office. Информацию о поддержке перечисления в том или ином приложении см. в разделе "Сведения о поддержке".

Дополнительные сведения о требованиях к серверу и ведущему приложению Office см. в статье [Требования к запуску надстроек для Office](../../docs/overview/requirements-for-running-office-add-ins.md).


|||
|:-----|:-----|
|**Типы надстроек**|Надстройки области задач, надстройки Outlook, контентные надстройки|
|**Library**|Office.js|
|**Пространство имен**|Office|
