# Элемент Action
 Указывает действие, которое необходимо выполнить, когда пользователь выбирает элемент управления [Button](./button-control.md) или [Menu](./menu-control.md).
 
## Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Да  | Тип выполняемого действия|


## Дочерние элементы

|  Элемент |  Описание  |
|:-----|:-----|
|  [FunctionName](#functionname) |    Указывает имя выполняемой функции. |
|  [SourceLocation](#sourcelocation) |    Указывает расположение исходного файла для этого действия. |
  

## xsi:type
Этот атрибут указывает действие, которое выполняется, когда пользователь нажимает кнопку. Допустимые значения:
- ExecuteFunction;
- ShowTaskpane.

## FunctionName
Обязательный элемент, если атрибуту **xsi:type** присвоено значение ExecuteFunction. Указывает имя выполняемой функции. Функция содержится в файле, указанном в элементе [FunctionFile](./functionfile.md).

```xml
<Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
</Action>
```

## SourceLocation
Обязательный элемент, если атрибуту **xsi:type** присвоено значение ShowTaskpane. Указывает расположение исходного файла для этого действия. Атрибуту **resid** должно быть присвоено значение атрибута **id** элемента **Url** в элементе [Urls](./resources.md#urls), включенном в элемент [Resources](./resources.md).

```xml
 <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
```  
