# Элемент CustomTab
На ленте можно указать вкладку и группу для команд надстройки. Это может быть вкладка по умолчанию (**Главная**, **Сообщение** или **Собрание**) либо специальная вкладка, которую определяет надстройка.

На специальных вкладках надстройка может создать до 10 групп. Каждая группа может включать не более 6 элементов управления, независимо от того, на какой вкладке она отображается. Надстройка может создать не более одной специальной вкладки.

Атрибут **id** должен быть уникальным для манифеста.

## Дочерние элементы
|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Group](./group.md)      | Да |  Определяет группу команд.  |
|  [Label](#label)      | Да |  Метка элемента CustomTab или Group.  |
|  [Control](#control)    | Да |  Коллекция из одного или нескольких объектов Control.  |

## Group
Обязательный. См. статью об [элементе Group](./group.md).

## Label (Tab)
Обязательный элемент. Метка настраиваемой вкладки. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **String** в элементе [ShortStrings](./resources.md#shortstrings), вложенном в элемент [Resources](./resources.md).


##  Пример элемента CustomTab
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```