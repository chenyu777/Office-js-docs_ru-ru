
# Переопределение элемента
Предоставляет способ указать значение параметра для дополнительного языкового стандарта.

 **Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.


## Синтаксис:


```XML
<Override Locale="string " Value="string " />
```


## Элемент, в котором содержится:


||
|:-----|
|[CitationText](../../reference/manifest/citationtext.md)|
|[Описание](../../reference/manifest/description.md)|
|[DictionaryName](../../reference/manifest/dictionaryname.md)|
|[DictionaryHomePage](../../reference/manifest/dictionaryhomepage.md)|
|[DisplayName](../../reference/manifest/displayname.md)|
|[HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md)|
|[IconUrl](../../reference/manifest/iconurl.md)|
|[QueryUri](../../reference/manifest/queryuri.md)|
|[SourceLocation](../../reference/manifest/sourcelocation.md)|
|[SupportUrl](../../reference/manifest/supporturl.md)|

## Атрибуты



|**Атрибут**|**Тип**|**Обязательный**|**Описание**|
|:-----|:-----|:-----|:-----|
|Языковой стандарт|string|Обязательный|Задает имя языка и региональных параметров для языкового стандарта этого переопределения в формате языковых тегов BCP 47, например `"en-US"`.|
|Значение|string|Обязательный|Задает значение параметра, представленное для указанного языкового стандарта.|

## Дополнительные ресурсы



- [Локализация надстроек для Office](../../docs/develop/localization.md#off15wecon_LocalesManifest)
    
