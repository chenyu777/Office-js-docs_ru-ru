# Объект TableCellCollection (API JavaScript для OneNote)

_Относится к: OneNote Online_  


Содержит коллекцию объектов TableCell.

## Свойства

| Свойство     | Тип   |Описание|Отзывы|
|:---------------|:--------|:----------|:-------|
|count|int|Возвращает количество объектов TableCell в этой коллекции. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection-count)|
|items|[TableCell[]](tablecell.md)|Коллекция объектов TableCell. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection-items)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[getItem(index: число или строка)](#getitemindex-число-или-строка)|[TableCell](tablecell.md)|Получает объект TableCell по идентификатору или индексу в коллекции. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection-getItem)|
|[getItemAt(index: число)](#getitematindex-число)|[TableCell](tablecell.md)|Получает объект TableCell по его позиции в коллекции.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection-load)|

## Сведения о методе


### getItem(index: число или строка)
Получает объект TableCell по идентификатору или индексу в коллекции. Только для чтения.

#### Синтаксис
```js
tableCellCollectionObject.getItem(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|число или строка|Число, определяющее расположение объекта TableCell по индексу.|

#### Возвращаемое значение
[TableCell](tablecell.md)

### getItemAt(index: число)
Получает объект TableCell по его позиции в коллекции.

#### Синтаксис
```js
tableCellCollectionObject.getItemAt(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number|Значение индекса получаемого объекта. Используется нулевой индекс.|

#### Возвращаемое значение
[TableCell](tablecell.md)

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
