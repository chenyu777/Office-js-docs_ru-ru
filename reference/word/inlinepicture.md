# Объект InlinePicture (API JavaScript для Word)

Представляет встроенный рисунок.

_Область применения: Word 2016, Word для iPad, Word для Mac_

## Свойства
| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|altTextDescription|string|Возвращает или задает строку, которая представляет замещающий текст, связанный с встроенным изображением.|
|altTextTitle|string|Возвращает или задает строку, содержащую заголовок встроенного рисунка.|
|hyperlink|string|Возвращает или задает гиперссылку, связанную со встроенным рисунком.|
|lockAspectRatio|Bool|Возвращает или задает значение, указывающее, сохраняет ли встроенный рисунок исходные пропорции при изменении размера.|

## Связи
| Связь | Тип   |Описание|
|:---------------|:--------|:----------|
|height|**с плавающей запятой**|Возвращает или задает высоту встроенного рисунка в точках. |
|parentContentControl|[ContentControl](contentcontrol.md)|Возвращает элемент управления содержимым, который содержит встроенный рисунок. Возвращает значение null, если родительского элемента управления содержимым не существует. Только для чтения.|
|paragraph|[paragraph](paragraph.md)|Возвращает абзац, который содержит встроенный рисунок. Только для чтения.
|width|**с плавающей запятой**|Возвращает или задает ширину встроенного рисунка в точках.|

## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Удаляет рисунок из документа.|
|[getBase64ImageSrc()](#getbase64imagesrc)|object|Возвращает объект, значение которого является строковым представлением встроенного рисунка в кодировке base64.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Вставляет разрыв в указанном расположении. Возможные значения InsertLocation: Before и After.|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Включает встроенный рисунок в элемент управления содержимым форматированного текста.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Вставляет документ в содержимое в заданном расположении. Возможные значения insertLocation: Before и After.|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Вставляет HTML-код в указанном расположении. Возможные значения InsertLocation: Before и After.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Вставляет рисунок в содержимое в заданном расположении. Возможные значения insertLocation: Replace, Before и After. |
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Вставляет OOXML-код в указанном расположении. Возможные значения InsertLocation: Before и After.|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Вставляет абзац в указанном расположении. Возможные значения InsertLocation: Before и After.|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Вставляет текст в содержимое в заданном расположении. Возможные значения insertLocation: Before и After.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Выбирает рисунок и создает к нему путь в пользовательском интерфейсе Word. Возможные значения selectionMode: Select, Start и End.|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|

## Сведения о методе

### delete()
Удаляет рисунок из документа.

#### Синтаксис
```js
inlinePictureObject.delete();
```

#### Параметры
Нет

#### Возвращаемое значение
void

### getBase64ImageSrc()
Возвращает объект, значение которого является строковым представлением встроенного рисунка в кодировке base64.

#### Синтаксис
```js
var base64 = inlinePictureObject.getBase64ImageSrc();
return context.sync().then(function () {    
    console.log("base64 string is " + base64.value);
});

```

#### Параметры
Нет

#### Возвращаемое значение
object 



### insertBreak(breakType: BreakType, insertLocation: InsertLocation)

#### Синтаксис
```js
inlinePictureObject.insertBreak(breakType, insertLocation);
```
#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|breakType|BreakType|Обязательный параметр. Тип разрыва, который необходимо добавить в содержимое.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Before и After.|

#### Возвращаемое значение
void

### insertContentControl()
Включает встроенный рисунок в элемент управления содержимым форматированного текста.

#### Синтаксис
```js
inlinePictureObject.insertContentControl();
```

#### Параметры
Нет

#### Возвращаемое значение
[ContentControl](contentcontrol.md)

### insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Вставляет документ в содержимое в заданном расположении. Возможные значения insertLocation: Before и After.

#### Синтаксис
```js
inlinePictureObject.insertFileFromBase64(base64File, insertLocation);
```
#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|base64File|string|Обязательный параметр. Содержимое DOCX-файла в кодировке base64.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Before и After.|

#### Возвращаемое значение
[Range](range.md)

### insertHtml(html: string, insertLocation: InsertLocation)
Вставляет HTML-код в указанном расположении. Возможные значения InsertLocation: Before и After.

#### Синтаксис
```js
inlinePictureObject.insertHtml(html, insertLocation);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|html|string|Обязательный параметр. HTML-код, который необходимо вставить в документ.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Before и After.|

#### Возвращаемое значение
[Range](range.md)


### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Вставляет рисунок в содержимое в заданном расположении. Возможные значения insertLocation: Before и After.

#### Синтаксис
inlinePictureObject.insertInlinePictureFromBase64(image, insertLocation);

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Обязательный параметр. Вставляемое в основной текст изображение в кодировке base64.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Before и After.|

#### Возвращаемое значение
[InlinePicture](inlinepicture.md)


### insertOoxml(ooxml: string, insertLocation: InsertLocation)
Вставляет OOXML-код в указанном расположении. Возможные значения InsertLocation: Before и After.

#### Синтаксис
```js
inlinePictureObject.insertOoxml(ooxml, insertLocation);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|"ooxml"|string|Обязательный параметр. Вставляемый OOXML-код.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Before и After.|

#### Возвращаемое значение
[Range](range.md)

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Вставляет абзац в указанном расположении. Возможные значения InsertLocation: Before и After.

#### Синтаксис
```js
inlinePictureObject.insertParagraph(paragraphText, insertLocation);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|paragraphText|string|Обязательный параметр. Текст абзаца, который необходимо вставить.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Before и After.|

#### Возвращаемое значение
[Paragraph](paragraph.md)

### insertText(text: string, insertLocation: InsertLocation)
Вставляет текст в содержимое в заданном расположении. Возможные значения insertLocation: Before и After.

#### Синтаксис
```js
inlinePictureObject.insertText(text, insertLocation);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|text|string|Обязательный параметр. Текст, который необходимо вставить.|
|insertLocation|InsertLocation|Обязательный параметр. Возможные значения: Before и After.|

#### Возвращаемое значение
[Range](range.md)

### select(selectionMode: SelectionMode)
Выбирает рисунок и создает к нему путь в пользовательском интерфейсе Word. Возможные значения selectionMode: Select, Start и End.

#### Синтаксис
```js
inlinePictureObject.select(selectionMode);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|Необязательный параметр. Возможные режимы выбора: Select, Start и End. Значение по умолчанию — Select.|

#### Возвращаемое значение
void

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

## Сведения о поддержке
Используйте [набор требований](../office-add-in-requirement-sets.md) в проверках среды выполнения, чтобы обеспечить поддержку ведущей версии Word для своего приложения. Дополнительные сведения о требованиях ведущих приложений и серверов Office см. в статье [Требования для запуска надстроек Office](../../docs/overview/requirements-for-running-office-add-ins.md).