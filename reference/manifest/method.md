
# Элемент Method
Указывает отдельный метод из API JavaScript для Office, необходимый для активации надстройки Office.

 **Тип надстройки:** контентные надстройки и надстройки области задач.


## Синтаксис:


```XML
<Method Name="string "/>
```


## Элемент, в котором содержится:

 _ [Методы](../../reference/manifest/methods.md)_


## Атрибуты



|**Атрибут**|**Тип**|**Обязательный**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя|string|Обязательный|Указывает имя необходимого метода, соответствующее его родительскому объекту. Например, чтобы задать метод **getSelectedDataAsync**, необходимо указать `"Document.getSelectedDataAsync"`.|

## Замечания

Элементы **Methods** и **Method** не поддерживаются для почтовых надстроек. Дополнительные сведения о наборах требований см. в статье [Указание ведущих приложений Office и требований к API](../../docs/overview/specify-office-hosts-and-api-requirements.md#SpecifyRequirementSets_intro).


 >**Внимание!** Минимальную версию невозможно указать для отдельных методов. Чтобы убедиться, что метод доступен в среде выполнения, при вызове этого метода в сценарии надстройки следует также использовать оператор **if**. Дополнительные сведения о том, как это сделать, см. в статье [Общие сведения об API JavaScript для Office](../../docs/develop/understanding-the-javascript-api-for-office.md#HostAPISupport_UsingIfStatements).

