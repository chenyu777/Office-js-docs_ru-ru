# Объект SectionCollection (API JavaScript для OneNote)

_Относится к: OneNote Online_  


Представляет коллекцию разделов.

## Свойства

| Свойство     | Тип   |Описание|Отзывы|
|:---------------|:--------|:----------|:-------|
|count|int|Возвращает количество разделов в коллекции. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-count)|
|items|[Section[]](section.md)|Коллекция объектов Section. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-items)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[getByName(name: string)](#getbynamename-string)|[SectionCollection](sectioncollection.md)|Получает коллекцию разделов с указанным именем.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-getByName)|
|[getItem(index: числовой тип или string)](#getitemindex-number-or-string)|[Раздел](section.md)|Получает раздел по идентификатору или индексу в коллекции. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[Раздел](section.md)|Получает раздел по позиции в коллекции.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionCollection-load)|

## Сведения о методе


### getByName(name: string)
Получает коллекцию объектов Section с указанным именем.

#### Синтаксис
```js
sectionCollectionObject.getByName(name);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|name|string|Имя раздела.|

#### Возвращаемое значение
[SectionCollection](sectioncollection.md)

#### Примеры
```js
OneNote.run(function (context) {

    // Get the sections in the current notebook.
    var sections = context.application.getActiveNotebook().sections;

    // Queue a command to load the sections. 
    // For best performance, request specific properties.
    sections.load("id"); 
    
    // Get the sections with the specified name.
    var groceriesSections = sections.getByName("Groceries");
    
    // Queue a command to load the sections with the specified name.
    groceriesSections.load("id,name");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Iterate through the collection or access items individually by index.
            if (groceriesSections.items.length > 0) {
                console.log("Section name: " + groceriesSections.items[0].name);
                console.log("Section ID: " + groceriesSections.items[0].id);
            }
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### getItem(index: числовой тип или string)
Получает объект Section по его идентификатору или индексу в коллекции. Только для чтения.

#### Синтаксис
```js
sectionCollectionObject.getItem(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number или string|Идентификатор объекта Section или расположение индекса этого объекта в коллекции.|

#### Возвращаемое значение
[Раздел](section.md)

### getItemAt(index: number)
Получает объект Section по его позиции в коллекции.

#### Синтаксис
```js
sectionCollectionObject.getItemAt(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number|Значение индекса получаемого объекта. Используется нулевой индекс.|

#### Возвращаемое значение
[Раздел](section.md)

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
### Примеры доступа к свойствам

**items**
```js
OneNote.run(function (context) {

    // Get the sections in the current notebook.
    var sections = context.application.getActiveNotebook().sections;

    // Queue a command to load the sections. 
    // For best performance, request specific properties.
    sections.load("name"); 

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Iterate through the collection or access items individually by index, for example: sections.items[0]
            $.each(sections.items, function(index, section) {
                if (section.name === "Homework") {
                    section.addPage("Biology");
                    section.addPage("Spanish");
                    section.addPage("Computer Science");
                }
            });
            return context.sync();
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

