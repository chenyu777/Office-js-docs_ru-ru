# Объект Notebook (API JavaScript для OneNote)

_Относится к: OneNote Online_   


Представляет записную книжку OneNote. Записные книжки содержат группы разделов и разделы.

## Свойства

| Свойство     | Тип   |Описание|Отзыв|
|:---------------|:--------|:----------|:-------|
|clientUrl|string|URL-адрес записной книжки для клиента. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-clientUrl)|
|id|string|Получает идентификатор записной книжки. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-id)|
|name|string|Получает имя записной книжки. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-name)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
| Связь | Тип   |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|sectionGroups|[SectionGroupCollection](sectiongroupcollection.md)|Группы разделов в записной книжке. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-sectionGroups)|
|sections|[SectionCollection](sectioncollection.md)|Разделы записной книжки. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-sections)|

## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[addSection(name: String)](#addsectionname-string)|[Раздел](section.md)|Добавляет новый раздел в конец записной книжки.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-addSection)|
|[addSectionGroup(name: String)](#addsectiongroupname-string)|[SectionGroup](sectiongroup.md)|Добавляет новую группу разделов в конец записной книжки.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-addSectionGroup)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebook-load)|

## Сведения о методе


### addSection(name: String)
Добавляет новый раздел в конец записной книжки.

#### Синтаксис
```js
notebookObject.addSection(name);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|name|String|Имя нового раздела.|

#### Возвращаемое значение
[Раздел](section.md)

#### Примеры
```js          
OneNote.run(function (context) {

    // Gets the active notebook.
    var notebook = context.application.getActiveNotebook();

    // Queue a command to add a new section. 
    var section = notebook.addSection("Sample section");
    
    // Queue a command to load the new section. This example reads the name property later.
    section.load("name");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log("New section name is " + section.name);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```


### addSectionGroup(name: String)
Добавляет новую группу разделов в конец записной книжки.

#### Синтаксис
```js
notebookObject.addSectionGroup(name);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|name|String|Имя нового раздела.|

#### Возвращаемое значение
[SectionGroup](sectiongroup.md)

#### Примеры
```js          
OneNote.run(function (context) {

    // Gets the active notebook.
    var notebook = context.application.getActiveNotebook();

    // Queue a command to add a new section group.
    var sectionGroup = notebook.addSectionGroup("Sample section group");

    // Queue a command to load the new section group.
    sectionGroup.load();

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log("New section group name is " + sectionGroup.name);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```

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
**id**
```js
OneNote.run(function (context) {
        
    // Get the current notebook.
    var notebook = context.application.getActiveNotebook();
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('id');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("Notebook ID: " + notebook.id);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**name**
```js
OneNote.run(function (context) {
        
    // Get the current notebook.
    var notebook = context.application.getActiveNotebook();
            
    // Queue a command to load the notebook. 
    // For best performance, request specific properties.           
    notebook.load('name');
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("Notebook name: " + notebook.name);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**sectionGroups**
```js          
OneNote.run(function (context) {

    // Get the section groups in the notebook. 
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;

    // Queue a command to load the sectionGroups. 
    sectionGroups.load("name");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            $.each(sectionGroups.items, function(index, sectionGroup) {
                console.log("Section group name: " + sectionGroup.name);
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

**sections**
```js
OneNote.run(function (context) {

    // Gets the active notebook.
    var notebook = context.application.getActiveNotebook();
    
    // Queue a command to get immediate child sections of the notebook. 
    var childSections = notebook.sections;

    // Queue a command to load the childSections. 
    context.load(childSections);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            $.each(childSections.items, function(index, childSection) {
                console.log("Immediate child section name: " + childSection.name);
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

