﻿# Объект ChartPoint (API JavaScript для Excel)

Представляет точку из ряда в диаграмме.

## Properties

| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|value|object|Возвращает значение точки диаграммы. Только для чтения.|

## Связи
| Связь | Тип   |Описание|
|:---------------|:--------|:----------|
|format|[ChartPointFormat](chartpointformat.md)|Инкапсулирует свойства формата точки диаграммы. Только для чтения.|

## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе


### load(param: object)
Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.

#### Синтаксис
```js
object.load(param);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|param|object|Необязательный параметр. Принимает имена параметров и связей в виде строки с разделителями или массива. Либо укажите объект [loadOption](loadoption.md).|

#### Возвращаемое значение
void
