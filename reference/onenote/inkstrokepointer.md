# Объект InkStrokePointer (API JavaScript для OneNote)

_Применяется для OneNote Online_  


Слабая ссылка на объект рукописного фрагмента и его родительский элемент содержимого

## Свойства

| Свойство     | Тип   |Описание|Отзыв|
|:---------------|:--------|:----------|:-------|
|contentId|string|Представляет соответствующий этому росчерку идентификатор объекта содержимого страницы|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokePointer-contentId)|
|inkStrokeId|string|Представляет идентификатор росчерка пера|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokePointer-inkStrokeId)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStrokePointer-load)|

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
