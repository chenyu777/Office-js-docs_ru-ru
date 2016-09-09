# Элемент VersionOverrides

Корневой элемент, который содержит сведения о командах надстройки. Элемент манифеста **VersionOverrides** является дочерним для элемента [OfficeApp](./officeapp.md). Этот элемент поддерживается в схеме манифестов версий 1.1 и выше, но определен в схеме VersionOverrides версии 1.0. 

## Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **xmlns**       |  Да  |  Расположение схемы. Необходимое значение — `http://schemas.microsoft.com/office/mailappversionoverrides`.|
|  **xsi:type**  |  Да  | Версия схемы. Пока допускается только значение `VersionOverridesV1_0`. |


## Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Описание**    |  НЕТ   |  Описывает надстройку. Переопределяет элемент `Description` в любой родительской части манифеста. Текст описания содержится в дочернем элементе **LongString**, включенном в элемент [Resources](./resources.md). Для атрибута `resid` элемента **Description** задано значение атрибута `id` элемента `String`, который содержит текст.|
|  **Requirements**  |  Нет   |  Задает минимальные набор требований и версию библиотеки Office.js, необходимые надстройке. Переопределяет элемент `Requirements` в родительской части манифеста.| 
|  [Hosts](./hosts.md)                |  Да  |  Задает набор узлов Office. Дочерний элемент Hosts переопределяет элемент Hosts в родительской части манифеста.  |
|  [Resources](./resources.md)    |  Да  | Определяет коллекцию ресурсов (строк, URL-адресов и изображений), на которые ссылаются другие элементы манифеста.|



### Пример элемента VersionOverrides
```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources> 
      <!-- add information on resources -->
   </Resources>
</VersionOverrides>
...
</OfficeApp>
```
