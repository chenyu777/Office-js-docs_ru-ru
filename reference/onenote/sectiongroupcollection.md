# Объект SectionGroupCollection (API JavaScript для OneNote)

_Относится к: OneNote Online_  


Представляет коллекцию групп разделов.

## Свойства

| Свойство     | Тип   |Описание|Отзывы|
|:---------------|:--------|:----------|:-------|
|count|int|Возвращает количество групп разделов в коллекции. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-count)|
|items|[SectionGroup[]](sectiongroup.md)|Коллекция объектов SectionGroup. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-items)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[getByName(name: string)](#getbynamename-string)|[SectionGroupCollection](sectiongroupcollection.md)|Получает коллекцию групп разделов с указанным именем.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-getByName)|
|[getItem(index: числовой тип или string)](#getitemindex-number-or-string)|[SectionGroup](sectiongroup.md)|Получает группу разделов по идентификатору или индексу в коллекции. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[SectionGroup](sectiongroup.md)|Получает объект SectionGroup по позиции в коллекции.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroupCollection-load)|

## Сведения о методе


### getByName(name: string)
Получает коллекцию объектов SectionGroup с указанным именем.

#### Синтаксис
```js
sectionGroupCollectionObject.getByName(name);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|name|string|Имя группы разделов.|

#### Возвращаемое значение
[SectionGroupCollection](sectiongroupcollection.md)

#### Примеры
```js
OneNote.run(function (context) {

    // Get the section groups that are direct children of the current notebook.
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;

    // Queue a command to load the section groups. 
    // For best performance, request specific properties.
    sectionGroups.load("id"); 

    // Get the section groups with the specified name.
    var labsSectionGroups = sectionGroups.getByName("Labs");

    // Queue a command to load the section groups with the specified properties.
    labsSectionGroups.load("id,name"); 
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Iterate through the collection or access items individually by index.
            if (labsSectionGroups.items.length > 0) {
                console.log("Section group name: " + labsSectionGroups.items[0].name);
                console.log("Section group ID: " + labsSectionGroups.items[0].id);
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
Получает объект SectionGroup по его идентификатору или индексу в коллекции. Только для чтения.

#### Синтаксис
```js
sectionGroupCollectionObject.getItem(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number или string|Идентификатор объекта SectionGroup или расположение индекса этого объекта в коллекции.|

#### Возвращаемое значение
[SectionGroup](sectiongroup.md)

### getItemAt(index: number)
Получает объект SectionGroup по его позиции в коллекции.

#### Синтаксис
```js
sectionGroupCollectionObject.getItemAt(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number|Значение индекса получаемого объекта. Используется нулевой индекс.|

#### Возвращаемое значение
[SectionGroup](sectiongroup.md)

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

    // Get the section groups that are direct children of the current notebook.
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;

    // Queue a command to load the section groups. 
    // For best performance, request specific properties.
    sectionGroups.load("name"); 

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Iterate through the collection or access items individually by index, for example: sectionGroups.items[0]
            $.each(sectionGroups.items, function(index, sectionGroup) {
                console.log("Section group name: " + sectionGroup.name);  
                console.log("Section group ID: " + sectionGroup.id);  
            });
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

