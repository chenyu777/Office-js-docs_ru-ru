# Элемент ExtensionPoint

 Определяет, где надстройка предоставляет функции в интерфейсе Office. Элемент **ExtensionPoint** — дочерний элемент элемента [FormFactor](./formfactor.md). 

## Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Да  | Тип определяемой точки расширения.|


## Точки расширения для команд надстроек Word, Excel, PowerPoint и OneNote

- **PrimaryCommandSurface** — лента в Office.
- **ContextMenu** — контекстное меню, которое появляется при нажатии правой кнопкой мыши в интерфейсе Office.

В следующих примерах показано, как использовать элемент **ExtensionPoint** со значениями атрибута **PrimaryCommandSurface** и **ContextMenu**, и какие дочерние элементы использовать с каждым из них.


 >**Важно!** Для элементов, содержащих атрибут идентификатора, необходимо предоставить уникальный идентификатор. Рекомендуем использовать название компании с идентификатором. Например, используйте следующий формат.<CustomTab id="mycompanyname.mygroupname">


```XML
 <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <CustomTab id="Contoso Tab">
            <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
             <!-- <OfficeTab id="TabData"> -->
              <Label resid="residLabel4" />
              <Group id="Group1Id12">
                <Label resid="residLabel4" />
                <Icon>
                  <bt:Image size="16" resid="icon1_32x32" />
                  <bt:Image size="32" resid="icon1_32x32" />
                  <bt:Image size="80" resid="icon1_32x32" />
                </Icon>
                <Tooltip resid="residToolTip" />
                <Control xsi:type="Button" id="Button1Id1">

                   <!-- information about the control -->
                </Control>
                <!-- other controls, as needed -->
              </Group>
            </CustomTab>
          </ExtensionPoint>

        <ExtensionPoint xsi:type="ContextMenu">
          <OfficeMenu id="ContextMenuCell">
            <Control xsi:type="Menu" id="ContextMenu2">
                   <!-- information about the control -->
            </Control>
           <!-- other controls, as needed -->
          </OfficeMenu>
         </ExtensionPoint>
```

**Дочерние элементы**
 
|**Элемент**|**Описание**|
|:-----|:-----|
|**CustomTab**|Обязательный, если требуется добавить на ленту настраиваемую вкладку (с помощью элемента **PrimaryCommandSurface**). Если используется элемент **CustomTab**, использовать элемент **OfficeTab** невозможно. Атрибут **id** является обязательным.|
|**OfficeTab**|Обязательный, если требуется расширить стандартную вкладку ленты Office (с помощью элемента **PrimaryCommandSurface**). Нельзя использовать элементы **OfficeTab** и **CustomTab** одновременно. Дополнительные сведения см. в статье [OfficeTab](officetab.md).|
|**OfficeMenu**|Обязательный при добавлении команд надстройки в контекстное меню по умолчанию (с помощью элемента **ContextMenu**). Для атрибута **id** необходимо задать следующее значение: <br/> - **ContextMenuText** для Excel или Word. Отображает элемент в контекстном меню, когда пользователь щелкает выделенный текст правой кнопкой мыши. <br/> - **ContextMenuCell** для Excel. Отображает элемент в контекстном меню, когда пользователь нажимает ячейку электронной таблицы правой кнопкой мыши.|
|**Group**|Группа точек расширения пользовательского интерфейса на вкладке. Группа может включать до шести элементов управления. Атрибут **id** является обязательным. Это строка длиной до 125 символов.|
|**Label**|Обязательный. Метка группы. Для атрибута **resid** необходимо задать значение атрибута **id** элемента **String**. Элемент **String** — это дочерний элемент элемента **ShortStrings**, который является дочерним для элемента **Resources**.|
|**Значок**|Обязательный. Задает значок группы, который будет использоваться на устройствах с малым форм-фактором либо при отображении слишком большого количества кнопок. Для атрибута **resid** необходимо задать значение атрибута **id** элемента **Image**. Элемент **Image** — это дочерний элемент элемента **Images**, который является дочерним для элемента **Resources**. Атрибут **size** указывает размер изображения в пикселях. Необходимо три размера изображения: 16, 32 и 80. Кроме того, поддерживается пять необязательных размеров: 20, 24, 40, 48 и 64.|
|**Tooltip**|Необязательный. Подсказка группы. Для атрибута **resid** необходимо задать значение атрибута **id** элемента **String**. Элемент **String** — это дочерний элемент элемента **LongStrings**, который является дочерним для элемента **Resources**.|
|**Control**|В каждой группе должен быть хотя бы один элемент управления. Элемент **Control** может иметь значение **Button** или **Menu**. Используйте элемент **Menu**, чтобы задать раскрывающийся список кнопок. В настоящий момент поддерживаются только кнопки и меню. Дополнительные сведения см. в разделах [Элементы управления "Кнопка"](#Элементы-управления-"Кнопка") и [Элементы управления меню](#Элементы-управления-меню).<br/>**Примечание.** Чтобы упростить устранение неполадок, рекомендуем добавлять элемент **Control** и соответствующий дочерний элемент **Resources** по одному.

## Точки расширения для команд надстроек Outlook

- [CustomPane](#custompane) 
- [MessageReadCommandSurface](#messagereadcommandsurface) 
- [MessageComposeCommandSurface](#messagecomposecommandsurface) 
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface) 
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) (можно использовать только в [DesktopFormFactor](./formfactor.md)).

### CustomPane

Точка расширения CustomPane определяет надстройку, которая активируется при выполнении определенных правил. Она предназначена только для формы чтения и отображается на горизонтальной области. 

**Дочерние элементы**

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **RequestedHeight** | Нет |  Запрашиваемая высота панели отображения при запуске на настольном компьютере (от 32 до 450 пикселей).  |
|  **SourceLocation**  | Да |  URL-адрес файла исходного кода надстройки. Речь идет об элементе **Url** в элементе [Resources](./resources.md).  |
|  **Rule**  | Да |  Правило или коллекция правил, которые определяют время активации надстройки. Дополнительные сведения см. в статье [Правила активации для надстроек Outlook](../../outlook/manifests/activation-rules.md). |
|  **DisableEntityHighlighting**  | Нет |  Указывает, следует ли отключить выделение сущностей. |


#### Пример элемента CustomPane
```xml
<ExtensionPoint xsi:type="CustomPane">
   <RequestedHeight>100< /RequestedHeight> 
   <SourceLocation resid="residReadTaskpaneUrl"/>
   <Rule xsi:type="RuleCollection" Mode="Or">
     <Rule xsi:type="ItemIs" ItemType="Message"/>
     <Rule xsi:type="ItemHasAttachment"/>
     <Rule xsi:type="ItemHasKnownEntity" EntityType="Address"/>
   </Rule>
</ExtensionPoint>
```

### MessageReadCommandSurface
Эта точка расширения помещает кнопки на панель команд для чтения почты. В классической версии Outlook эта панель отображается на ленте.

**Дочерние элементы**

|  Элемент |  Описание  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Добавляет команды на вкладку ленты по умолчанию.  |
|  [CustomTab](./customtab.md) |  Добавляет команды на специальную вкладку ленты.  |

#### Пример элемента OfficeTab
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### Пример элемента CustomTab
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```
### MessageComposeCommandSurface
Эта точка расширения добавляет кнопки на ленту для надстроек, использующих форму создания сообщения. 

**Дочерние элементы**

|  Элемент |  Описание  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Добавляет команды на вкладку ленты по умолчанию.  |
|  [CustomTab](./customtab.md) |  Добавляет команды на специальную вкладку ленты.  |

#### Пример элемента OfficeTab
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### Пример элемента CustomTab

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```
### AppointmentOrganizerCommandSurface

Эта точка расширения добавляет кнопки на ленту для формы, предназначенной для организатора собрания. 

**Дочерние элементы**

|  Элемент |  Описание  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Добавляет команды на вкладку ленты по умолчанию.  |
|  [CustomTab](./customtab.md) |  Добавляет команды на специальную вкладку ленты.  |

#### Пример элемента OfficeTab
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### Пример элемента CustomTab
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### AppointmentAttendeeCommandSurface

Эта точка расширения добавляет кнопки на ленту для формы, предназначенной для участника собрания. 

**Дочерние элементы**

|  Элемент |  Описание  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Добавляет команды на вкладку ленты по умолчанию.  |
|  [CustomTab](./customtab.md) |  Добавляет команды на специальную вкладку ленты.  |

#### Пример элемента OfficeTab
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### Пример элемента CustomTab
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### Module

Эта точка расширения добавляет кнопки на ленту для расширения модуля. 

**Дочерние элементы**

|  Элемент |  Описание  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  Добавляет команды на вкладку ленты по умолчанию.  |
|  [CustomTab](./customtab.md) |  Добавляет команды на специальную вкладку ленты.  |

