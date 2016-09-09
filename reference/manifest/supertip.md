## Supertip
Определяет форматированную подсказку (элементы Title и Description). Используется элементами управления [Button](./button.md) и [Menu](./menu-control.md). 

## Дочерние элементы
|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Title](#title)        | Да |   Текст подсказки.         |
|  [Description](#description)  | Да |  Описание подсказки.    |

## Title
Обязательный элемент. Текст суперподсказки. Атрибуту  **resid** должно быть присвоено значение атрибута **id** элемента **String** в элементе [ShortStrings](./resources.md#shortstrings), вложенном в элемент  [Resources](./resources.md).

## Описание
Обязательный элемент. Описание суперподсказки. Атрибуту  **resid** должно быть присвоено значение атрибута **id** элемента **String** в элементе [LongStrings](./resources.md#longstrings), вложенном в элемент  [Resources](./resources.md).

```xml
 <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
```