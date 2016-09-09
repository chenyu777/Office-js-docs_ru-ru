# Объект InkAnalysisLineCollection (API JavaScript для OneNote)

_Применяется для OneNote Online_  


Представляет коллекцию объектов InkAnalysisLine.

## Свойства

| Свойство     | Тип   |Описание|Отзывы|
|:---------------|:--------|:----------|:-------|
|count|int|Возвращает количество объектов InkAnalysisLine на странице. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-count)|
|items|[InkAnalysisLine[]](inkanalysisline.md)|Коллекция объектов InkAnalysisLine. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-items)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number или string)](#getitemindex-number-или-string)|[InkAnalysisLine](inkanalysisline.md)|Получает объект InkAnalysisLine по идентификатору или индексу в коллекции. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkAnalysisLine](inkanalysisline.md)|Получает объект InkAnalysisLine по его позиции в коллекции.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLineCollection-load)|

## Сведения о методе


### getItem(index: number или string)
Получает объект InkAnalysisLine по идентификатору или индексу в коллекции. Только для чтения.

#### Синтаксис
```js
inkAnalysisLineCollectionObject.getItem(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number или string|Идентификатор объекта InkAnalysisLine или расположение индекса объекта InkAnalysisLine в коллекции.|

#### Возвращаемое значение
[InkAnalysisLine](inkanalysisline.md)

### getItemAt(index: number)
Получает объект InkAnalysisLine по его позиции в коллекции.

#### Синтаксис
```js
inkAnalysisLineCollectionObject.getItemAt(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number|Значение индекса получаемого объекта. Используется нулевой индекс.|

#### Возвращаемое значение
[InkAnalysisLine](inkanalysisline.md)

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
