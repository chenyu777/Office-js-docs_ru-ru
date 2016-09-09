
# Элемент Permissions
Указывает уровень доступа к API для надстройки Office. Запрашивая разрешения, руководствуйтесь принципом минимальных привилегий.

 **Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.


## Синтаксис:

Для надстроек области задач и контентных надстроек:


```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

Для почтовых надстроек:




```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```


## Элемент, в котором содержится:

 _[элемент OfficeApp](../../reference/manifest/officeapp.md)_


## Замечания

Подробные сведения см. в статьях [Запрашивание разрешений на использование API в надстройках области задач и контентных надстройках](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) и [Общие сведения о разрешениях для надстроек Outlook](../../docs/outlook/understanding-outlook-add-in-permissions.md).

