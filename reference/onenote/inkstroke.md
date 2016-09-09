# Объект InkStroke (API JavaScript для OneNote)

_Применяется для OneNote Online_    


Представляет один росчерк пера.

## Свойства

| Свойство     | Тип   |Описание|Отзывы|
|:---------------|:--------|:----------|:-------|
|id|string|Получает идентификатор объекта InkStroke. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStroke-id)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
| Связь | Тип   |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|floatingInk|[FloatingInk](floatingink.md)|Получает идентификатор объекта InkStroke. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStroke-floatingInk)|

## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkStroke-load)|

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
