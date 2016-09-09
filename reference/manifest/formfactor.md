# Элемент FormFactor

Указывает параметры для надстройки заданного форм-фактора. Например, если определить `Host` как `MailHost`, а форм-фактор как `DesktopFormFactor`, параметры будут применяться к классическому приложению Outlook, но _не_ к Outlook Web App или Outlook.com. Он содержит все сведения о надстройке для этого форм-фактора, кроме узла **Resources**.

Каждое определение FormFactor содержит элемент **FunctionFile** и один или несколько элементов **ExtensionPoint**. Дополнительные сведения см. в статьях [Элемент FunctionFile](./functionfile.md) и [Элемент ExtensionPoint](./extensionpoint.md). 

Поддерживаются следующие элементы FormFactor:

- `DesktopFormFactor` (клиенты Office для Windows или Mac).

## Дочерние элементы

| Элемент                               | Обязательный | Описание  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](./extensionpoint.md) | Да      | Определяет, где надстройка предоставляет функции. |
| [FunctionFile](./functionfile.md)     | Да      | URL-адрес файла, который содержит функции JavaScript.|
| [GetStarted](./getstarted.md)         | Нет       | Определяет выноску, которая отображается при установке надстройки в ведущих приложениях Word, Excel и PowerPoint. |

## Пример элемента FormFactor

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
