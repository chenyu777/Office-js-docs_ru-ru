﻿# Элемент Icon
Определяет элементы **Image** для элементов управления [Button](./button.md) и [Menu](./menu-control.md).

## Дочерние элементы
|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Изображение](#Изображение)        | Да |   атрибут resid используемого изображения         |

## Изображение
Изображение кнопки. Атрибуту **resid** нужно присвоить значение атрибута **id** элемента **Image** в элементе **Images** в элементе [Resources](./resources.md). Атрибут **size** указывает размер изображения в пикселях. Обязательными являются три размера изображения (16, 32 и 80 пикселей), а поддерживаются еще пять (20, 24, 40, 48 и 64 пикселя).|


```xml
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
```  