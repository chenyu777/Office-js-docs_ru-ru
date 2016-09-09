# Объект InkWordCollection (API JavaScript для OneNote)

_Применяется для OneNote Online_  


Представляет коллекцию объектов InkWord.

## Свойства

| Свойство     | Тип   |Описание|Отзывы|
|:---------------|:--------|:----------|:-------|
|count|int|Возвращает количество объектов InkWord на странице. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-count)|
|items|[InkWord[]](inkword.md)|Коллекция объектов inkWord. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-items)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number или string)](#getitemindex-number-или-string)|[InkWord](inkword.md)|Получает объект InkWord по идентификатору или индексу в коллекции. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkWord](inkword.md)|Получает объект InkWord по позиции в коллекции.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-load)|

## Сведения о методе


### getItem(index: number или string)
Получает объект InkWord по идентификатору или индексу в коллекции. Только для чтения.

#### Синтаксис
```js
inkWordCollectionObject.getItem(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number или string|Идентификатор объекта InkWord или расположение индекса объекта InkWord в коллекции.|

#### Возвращаемое значение
[InkWord](inkword.md)

### getItemAt(index: number)
Получает объект InkWord по позиции в коллекции.

#### Синтаксис
```js
inkWordCollectionObject.getItemAt(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number|Значение индекса получаемого объекта. Используется нулевой индекс.|

#### Возвращаемое значение
[InkWord](inkword.md)

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
