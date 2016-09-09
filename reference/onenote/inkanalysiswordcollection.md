# Объект InkAnalysisWordCollection (API JavaScript для OneNote)

_Применяется для OneNote Online_  


Представляет коллекцию объектов InkAnalysisWord.

## Свойства

| Свойство     | Тип   |Описание|Отзывы|
|:---------------|:--------|:----------|:-------|
|count|int|Возвращает количество объектов InkAnalysisWord на странице. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWordCollection-count)|
|items|[InkAnalysisWord[]](inkanalysisword.md)|Коллекция объектов InkAnalysisWord. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWordCollection-items)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number или string)](#getitemindex-number-или-string)|[InkAnalysisWord](inkanalysisword.md)|Получает объект InkAnalysisWord по идентификатору или индексу в коллекции. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWordCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkAnalysisWord](inkanalysisword.md)|Получает объект InkAnalysisWord по его позиции в коллекции.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWordCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWordCollection-load)|

## Сведения о методе


### getItem(index: number или string)
Получает объект InkAnalysisWord по идентификатору или индексу в коллекции. Только для чтения.

#### Синтаксис
```js
inkAnalysisWordCollectionObject.getItem(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number или string|Идентификатор объекта InkAnalysisWord или расположение индекса объекта InkAnalysisWord в коллекции.|

#### Возвращаемое значение
[InkAnalysisWord](inkanalysisword.md)

### getItemAt(index: number)
Получает объект InkAnalysisWord по его позиции в коллекции.

#### Синтаксис
```js
inkAnalysisWordCollectionObject.getItemAt(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number|Значение индекса получаемого объекта. Используется нулевой индекс.|

#### Возвращаемое значение
[InkAnalysisWord](inkanalysisword.md)

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
