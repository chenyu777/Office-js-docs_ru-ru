# Элемент OfficeMenu
Определяет коллекцию элементов управления, которые нужно добавить в контекстное меню Office. Применяется в надстройках Word, Excel, PowerPoint и OneNote.

## Атрибуты

| Атрибут            | Обязательный | Описание                          |
|:---------------------|:--------:|:-------------------------------------|
| [xsi:type](#xsitype) | Да      | Тип определяемого элемента OfficeMenu.|

## Дочерние элементы
|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Control](#control)    | Да |  Коллекция из одного или нескольких объектов Control.  |

## xsi:type
Указывает то встроенное меню клиентского приложения Office, в которое необходимо добавить название надстройки Office.

- `ContextMenuText`. Отображает элемент в контекстном меню, когда пользователь открывает это меню, щелкая правой кнопкой мыши по выделенному тексту. Применяется для Word, Excel, PowerPoint и OneNote.
- `ContextMenuCell`. Отображает элемент в контекстном меню, когда пользователь открывает это меню, щелкая правой кнопкой мыши ячейку электронной таблицы. Применяется для Excel. 

## Control

Для каждого элемента **OfficeMenu** требуется один или несколько элементов управления [меню](./menu.md#menu-control). 


## Пример

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="myMenuID">
      <Label resid="residLabel3" />
      <Supertip>
          <Title resid="residLabel" />
          <Description resid="residToolTip" />
      </Supertip>   
      <Icon>
        <bt:Image size="16" resid="icon1_16x16" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_80x80" />
      </Icon>    
      <Items>
        <Item id="myMenuItemID">
          <Label resid="residLabel3"/>
          <Supertip>
            <Title resid="residLabel" />
            <Description resid="residToolTip" />
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="icon1_16x16" />
            <bt:Image size="32" resid="icon1_32x32" />
            <bt:Image size="80" resid="icon1_80x80" />
          </Icon>    
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl2" />    
          </Action>    
        </Item>
      </Items>
    </Control>   
</OfficeMenu>
```
