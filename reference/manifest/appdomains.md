
# Элемент AppDomains
Указывает дополнительные домены, которые надстройка Office будет использовать для загрузки страниц.

 **Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.


## Синтаксис:


```XML
<AppDomains>
   ...
</AppDomains>
```


## Элемент, в котором содержится:

[элемент OfficeApp](../../reference/manifest/officeapp.md)


## Может содержать:

[AppDomain](../../reference/manifest/appdomain.md)


## Замечания

Элементы **AppDomains** и **AppDomain** используются для указания дополнительных доменов, отличных от указанного в элементе [SourceLocation](../../reference/manifest/sourcelocation.md). Дополнительные сведения см. в статье [XML-манифест надстроек Office](../../docs/overview/add-in-manifests.md).

