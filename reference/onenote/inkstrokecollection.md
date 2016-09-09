# Объект InkStrokeCollection (API JavaScript для OneNote)

_Применяется для OneNote Online_   


Представляет коллекцию объектов InkStroke.

## Свойства

| Свойство     | Тип   |Описание|Отзывы|
|:---------------|:--------|:----------|:-------|
|count|int|Возвращает количество объектов InkStroke на странице. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokeCollection-count)|
|items|[InkStroke[]](inkstroke.md)|Коллекция объектов inkStroke. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokeCollection-items)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number или string)](#getitemindex-number-или-string)|[InkStroke](inkstroke.md)|Получает объект InkStroke по идентификатору или индексу в коллекции. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokeCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkStroke](inkstroke.md)|Получает объект InkStroke по его позиции в коллекции.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokeCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokeCollection-load)|

## Сведения о методе


### getItem(index: number или string)
Получает объект InkStroke по идентификатору или индексу в коллекции. Только для чтения.

#### Синтаксис
```js
inkStrokeCollectionObject.getItem(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number или string|Идентификатор объекта InkStroke или расположение индекса объекта InkStroke в коллекции.|

#### Возвращаемое значение
[InkStroke](inkstroke.md)

### getItemAt(index: number)
Получает объект InkStroke по его позиции в коллекции.

#### Синтаксис
```js
inkStrokeCollectionObject.getItemAt(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number|Значение индекса получаемого объекта. Используется нулевой индекс.|

#### Возвращаемое значение
[InkStroke](inkstroke.md)

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
