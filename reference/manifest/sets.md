
# Элемент Sets
Указывает минимальное подмножество API JavaScript для Office, необходимое для активации надстройки Office.

 **Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.


## Синтаксис:


```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```


## Элемент, в котором содержится:

[Requirements](../../reference/manifest/requirements.md)


## Может содержать:

[Set](../../reference/manifest/set.md)


## Атрибуты



|**Атрибут**|**Тип**|**Обязательный**|**Описание**|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|string|необязательный|Задает значение атрибута **MinVersion** по умолчанию для всех дочерних элементов [Set](../../reference/manifest/set.md). Значение по умолчанию: "1.1".|

## Замечания

Дополнительные сведения о наборах требований см. в статье [Указание ведущих приложений Office и требований к API](../../docs/overview/specify-office-hosts-and-api-requirements.md).

Дополнительные сведения об атрибуте **MinVersion** элемента **Set** и атрибуте **DefaultMinVersion** элемента **Sets** см. в статье [Указание элемента Requirements в манифесте](../../docs/overview/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).

