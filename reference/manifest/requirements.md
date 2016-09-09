
# Элемент Requirements
Указывает минимальный набор элементов API JavaScript для Office ([набор требований](../../docs/overview/specify-office-hosts-and-api-requirements.md#SpecifyRequirementSets_sets) и/или методов), необходимых для активации надстройки Office.

 **Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.


## Синтаксис:


```XML
<Requirements>
   ...
</Requirements>
```


## Элемент, в котором содержится:

[элемент OfficeApp](../../reference/manifest/officeapp.md)


## Может содержать:



|**Элемент**|**Содержимое**|**Почтовое**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Sets](../../reference/manifest/sets.md)|x|x|x|
|[Методы](../../reference/manifest/methods.md)|x||x|

## Заметки

Дополнительные сведения о наборах требований см. в статье [Указание ведущих приложений Office и требований к API](../../docs/overview/specify-office-hosts-and-api-requirements.md).

