# Объект Icon (API JavaScript для Excel)

_Относится к: Excel 2016, Excel Online, Excel для iOS, Office 2016_

Представляет собой значок ячейки.

## Свойства

| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|index|int|Представляет собой индекс значка данного набора.|
|set|string|Представляет собой набор, в который входит значок. Возможные значения: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|

_Ознакомьтесь с [примерами](#примерами) доступа к свойствам._

## Связи
Нет


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
