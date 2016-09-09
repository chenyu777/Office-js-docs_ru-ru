# Объект InkAnalysisParagraphCollection (API JavaScript для OneNote)

_Применяется для OneNote Online_  


Представляет коллекцию объектов InkAnalysisParagraph.

## Свойства

| Свойство     | Тип   |Описание|Отзывы|
|:---------------|:--------|:----------|:-------|
|count|int|Возвращает количество объектов InkAnalysisParagraph в объекте Page. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-count)|
|items|[InkAnalysisParagraph[]](inkanalysisparagraph.md)|Коллекция объектов InkAnalysisParagraph. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-items)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number или string)](#getitemindex-number-или-string)|[InkAnalysisParagraph](inkanalysisparagraph.md)|Получает объект InkAnalysisParagraph по идентификатору или индексу в коллекции. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkAnalysisParagraph](inkanalysisparagraph.md)|Получает объект InkAnalysisParagraph по его позиции в коллекции.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraphCollection-load)|

## Сведения о методе


### getItem(index: number или string)
Получает объект InkAnalysisParagraph по идентификатору или индексу в коллекции. Только для чтения.

#### Синтаксис
```js
inkAnalysisParagraphCollectionObject.getItem(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number или string|Идентификатор объекта InkAnalysisParagraph или расположение индекса объекта InkAnalysisParagraph в коллекции.|

#### Возвращаемое значение
[InkAnalysisParagraph](inkanalysisparagraph.md)

### getItemAt(index: number)
Получает объект InkAnalysisParagraph по его позиции в коллекции.

#### Синтаксис
```js
inkAnalysisParagraphCollectionObject.getItemAt(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number|Значение индекса получаемого объекта. Используется нулевой индекс.|

#### Возвращаемое значение
[InkAnalysisParagraph](inkanalysisparagraph.md)

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
