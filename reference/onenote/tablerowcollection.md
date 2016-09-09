# Объект TableRowCollection (API JavaScript для OneNote)

_Относится к: OneNote Online_  


Содержит коллекцию объектов TableRow.

## Свойства

| Свойство     | Тип   |Описание|Отзывы|
|:---------------|:--------|:----------|:-------|
|count|int|Возвращает количество объектов TableRow в этой коллекции. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-count)|
|items|[TableRow[]](tablerow.md)|Коллекция объектов TableRow. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-items)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[getItem(index: число или строка)](#getitemindex-число-или-строка)|[TableRow](tablerow.md)|Получает объект TableRow по идентификатору или индексу в коллекции. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-getItem)|
|[getItemAt(index: число)](#getitematindex-число)|[TableRow](tablerow.md)|Получает объект TableRow по позиции в коллекции.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-load)|

## Сведения о методе


### getItem(index: число или строка)
Получает объект TableRow по идентификатору или индексу в коллекции. Только для чтения.

#### Синтаксис
```js
tableRowCollectionObject.getItem(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|число или строка|Число, определяющее расположение объекта TableRow по индексу.|

#### Возвращаемое значение
[TableRow](tablerow.md)

### getItemAt(index: число)
Получает объект TableRow по позиции в коллекции.

#### Синтаксис
```js
tableRowCollectionObject.getItemAt(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number|Значение индекса получаемого объекта. Используется нулевой индекс.|

#### Возвращаемое значение
[TableRow](tablerow.md)

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
