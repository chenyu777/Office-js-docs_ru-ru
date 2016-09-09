# Элемент Control

Определяет функцию JavaScript, которая выполняет действие или открывает область задач. Элемент **Control** может быть кнопкой или пунктом меню. Элемент [Group](group.md) должен содержать по крайней мере один элемент **Control**.

## Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|**xsi:type**|Да|Тип определяемого элемента управления. Возможные типы: Button и Menu.|
|**id**|Нет|ИД элемента управления. Может содержать до 125 знаков.|

## Элемент управления Button

Кнопка выполняет одно действие, когда пользователь ее нажимает. Она может выполнять функцию или отображать область задач. Каждый элемент управления "Кнопка" должен иметь элемент `id`, уникальный для манифеста. 

### Дочерние элементы
|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Label**     | Да |  Текст для кнопки. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **String** элементе [ShortStrings](./resources.md#shortstrings), вложенном в элемент [Resources](./resources.md).        |
|  **ToolTip**  |Нет|Подсказка для кнопки. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String** — это дочерний элемент **LongStrings**, который является дочерним для элемента [Resources](resource.md).|     
|  [Supertip](./supertip.md)  | Да |  Суперподсказка для кнопки.    |
|  [Icon](./icon.md)      | Да |  Изображение для кнопки.         |
|  [Action](./action.md)    | Да |  Задает выполняемое действие.  |



```XML
        <!-- Define a control that calls a JavaScript function. -->

                 <Control xsi:type="Button" id="Button1Id1">
                  <Label resid="residLabel" />
                  <Tooltip resid="residToolTip" />
                  <Supertip>
                    <Title resid="residLabel" />
                    <Description resid="residToolTip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon1_32x32" />
                    <bt:Image size="32" resid="icon1_32x32" />
                    <bt:Image size="80" resid="icon1_32x32" />
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>getData</FunctionName>
                  </Action>
                </Control>


                <!-- Define a control that shows a task pane. -->

                <Control xsi:type="Button" id="Button2Id1">
                  <Label resid="residLabel2" />
                  <Tooltip resid="residToolTip" />
                  <Supertip>
                    <Title resid="residLabel" />
                    <Description resid="residToolTip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon2_32x32" />
                    <bt:Image size="32" resid="icon2_32x32" />
                    <bt:Image size="80" resid="icon2_32x32" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="residUnitConverterUrl" />
                  </Action>
                </Control>
```

### Пример кнопки ExecuteFunction

```xml
<Control xsi:type="Button" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### Пример кнопки ShowTaskpane

```xml
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```
## Элементы управления "Меню" (кнопка с раскрывающимся списком)

Меню определяет статический список вариантов. Каждый элемент меню либо выполняет функцию, либо отображает область задач. Вложенные меню не поддерживаются. 

При использовании с [точкой расширения](extensionpoint.md) **ContextMenu****PrimaryCommandSurface** элемент управления Menu определяет следующее:

- элемент меню корневого уровня;

- список элементов подменю.

При использовании с **PrimaryCommandSurface** корневой элемент меню отображает кнопку на ленте. По нажатию кнопки в подменю отображается раскрывающийся список. При использовании с **ContextMenu** в контекстное меню вставляется элемент меню с подменю. В обоих случаях отдельные элементы подменю могут либо вызывать функцию JavaScript, либо отображать область задач. В настоящее время поддерживается только один уровень подменю.

В приведенном ниже примере показано, как определить элемент меню с двумя элементами подменю. Первый элемент подменю отображает область задач, а второй запускает функцию JavaScript.

```xml
<Control xsi:type="Menu" id="TestMenu2">
              <Label resid="residLabel3" />
              <Tooltip resid="residToolTip" />
              <Supertip>
                <Title resid="residLabel" />
                <Description resid="residToolTip" />
              </Supertip>
              <Icon>
                <bt:Image size="16" resid="icon1_32x32" />
                <bt:Image size="32" resid="icon1_32x32" />
                <bt:Image size="80" resid="icon1_32x32" />
              </Icon>
              <Items>
                <Item id="showGallery2">
                  <Label resid="residLabel3"/>
                  <Supertip>
                    <Title resid="residLabel" />
                    <Description resid="residToolTip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon1_32x32" />
                    <bt:Image size="32" resid="icon1_32x32" />
                    <bt:Image size="80" resid="icon1_32x32" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>MyTaskPaneID1</TaskpaneId>
                    <SourceLocation resid="residUnitConverterUrl" />
                  </Action>
                </Item>
              <Item id="showGallery3">
                  <Label resid="residLabel5"/>
                  <Supertip>
                    <Title resid="residLabel" />
                    <Description resid="residToolTip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon4_32x32" />
                    <bt:Image size="32" resid="icon4_32x32" />
                    <bt:Image size="80" resid="icon4_32x32" />
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>getButton</FunctionName>
                  </Action>
                </Item>
              </Items>
            </Control>

```

### Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Label**     | Да |  Текст для кнопки. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **String** в элементе [ShortStrings](./resources.md#shortstrings), вложенном в элемент [Resources](./resources.md).      |
|  **ToolTip**  |Нет|Подсказка для кнопки. Для атрибута **resid** должно быть задано значение атрибута **id**, принадлежащего элементу **String**. **String** — это дочерний элемент **LongStrings**, который является дочерним для элемента [Resources](resource.md).|     
|  [Supertip](./supertip.md)  | Да |  Суперподсказка для кнопки.    |
|  [Icon](./icon.md)      | Да |  Изображение для кнопки.         |
|  [Items](#items)     | Да |  Коллекция кнопок, отображающихся в меню. Содержит элементы **Item** для каждого элемента подменю. Каждый элемент **Item** содержит дочерние элементы, вложенные в [элемент управления Button](#элемент-управления-button).|


### Примеры элементов управления Menu

```xml
<Control xsi:type="Menu" id="TestMenu2">
              <Label resid="residLabel3" />
              <Tooltip resid="residToolTip" />
              <Supertip>
                <Title resid="residLabel" />
                <Description resid="residToolTip" />
              </Supertip>
              <Icon>
                <bt:Image size="16" resid="icon1_32x32" />
                <bt:Image size="32" resid="icon1_32x32" />
                <bt:Image size="80" resid="icon1_32x32" />
              </Icon>
              <Items>
                <Item id="showGallery2">
                  <Label resid="residLabel3"/>
                  <Supertip>
                    <Title resid="residLabel" />
                    <Description resid="residToolTip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon1_32x32" />
                    <bt:Image size="32" resid="icon1_32x32" />
                    <bt:Image size="80" resid="icon1_32x32" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>MyTaskPaneID1</TaskpaneId>
                    <SourceLocation resid="residUnitConverterUrl" />
                  </Action>
                </Item>
              <Item id="showGallery3">
                  <Label resid="residLabel5"/>
                  <Supertip>
                    <Title resid="residLabel" />
                    <Description resid="residToolTip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon4_32x32" />
                    <bt:Image size="32" resid="icon4_32x32" />
                    <bt:Image size="80" resid="icon4_32x32" />
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>getButton</FunctionName>
                  </Action>
                </Item>
              </Items>
            </Control>

```


```xml
<Control xsi:type="Menu" id="msgReadMenuButton">
  <Label resid="menuReadButtonLabel" />
  <Supertip>
    <Title resid="menuReadSuperTipTitle" />
    <Description resid="menuReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="red-icon-16" />
    <bt:Image size="32" resid="red-icon-32" />
    <bt:Image size="80" resid="red-icon-80" />
  </Icon>
  <Items>
    <Item id="msgReadMenuItem1">
      <Label resid="menuItem1ReadLabel" />
      <Supertip>
        <Title resid="menuItem1ReadLabel" />
        <Description resid="menuItem1ReadTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getItemClass</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```
