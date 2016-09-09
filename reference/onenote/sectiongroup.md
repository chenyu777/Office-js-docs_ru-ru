# Объект SectionGroup (API JavaScript для OneNote)

_Относится к: OneNote Online_   


Представляет группу разделов OneNote. Объекты SectionGroup могут содержать объекты Section и другие объекты SectionGroup.

## Свойства

| Свойство     | Тип   |Описание|Отзыв|
|:---------------|:--------|:----------|:-------|
|clientUrl{|string|URL-адрес клиента группы разделов. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-clientUrl{)|
|id|string|Получает идентификатор объекта SectionGroup. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-id)|
|name|string|Имя группы разделов. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-name)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
| Связь | Тип   |Описание| Отзывы|
|:---------------|:--------|:----------|:-------|
|notebook|[Notebook](notebook.md)|Получает записную книжку, содержащую группу разделов. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-notebook)|
|parentSectionGroup|[SectionGroup](sectiongroup.md)|Получает группу разделов, содержащую другую группу разделов. Возвращает значение ItemNotFound, если групп разделов является прямым потомком записной книжки. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-parentSectionGroup)|
|parentSectionGroupOrNull|[SectionGroup](sectiongroup.md)|Получает группу разделов, содержащую другую группу разделов. Возвращает значение null, если объект SectionGroup является прямым потомком объекта Notebook. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-parentSectionGroupOrNull)|
|sectionGroups|[SectionGroupCollection](sectiongroupcollection.md)|Коллекция групп разделов в группе разделов. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-sectionGroups)|
|sections|[SectionCollection](sectioncollection.md)|Коллекция разделов в группе разделов. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-sections)|

## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[addSection(title: String)](#addsectiontitle-string)|[Раздел](section.md)|Добавляет новый раздел в конец группы разделов.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-addSection)|
|[addSectionGroup(name: String)](#addsectiongroupname-string)|[SectionGroup](sectiongroup.md)|Добавляет новую группу разделов в конец sectionGroup.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-addSectionGroup)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-sectionGroup-load)|

## Сведения о методе


### addSection(title: String)
Добавляет новый раздел в конец группы разделов.

#### Синтаксис
```js
sectionGroupObject.addSection(title);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|title|String|Имя нового раздела.|

#### Возвращаемое значение
[Раздел](section.md)

#### Примеры
```js
OneNote.run(function (context) {

    // Get the section groups that are direct children of the current notebook.
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;
    
    // Queue a command to load the section groups.
    // For best performance, request specific properties.
    sectionGroups.load("id");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            
            // Add a section to each section group.
            $.each(sectionGroups.items, function(index, sectionGroup) {
                sectionGroup.addSection("Agenda");
            });
            
            // Run the queued commands.
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


### addSectionGroup(name: String)
Добавляет новую группу разделов в конец sectionGroup.

#### Синтаксис
```js
sectionGroupObject.addSectionGroup(name);
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
    var sectionGroup;
    var nestedSectionGroup;

    // Gets the active notebook.
    var notebook = context.application.getActiveNotebook();

    // Queue a command to add a new section group.
    var sectionGroups = notebook.sectionGroups;

    // Queue a command to load the new section group.
    sectionGroups.load();

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function(){
            sectionGroup = sectionGroups.items[0];
            sectionGroup.load();
            return context.sync();
        })
        .then(function(){
            nestedSectionGroup = sectionGroup.addSectionGroup("Sample nested section group");
            nestedSectionGroup.load();
            return context.sync();
        })
        .then(function() {
            console.log("New nested section group name is " + nestedSectionGroup.name);
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
        
    // Get the parent section group that contains the current section.
    var sectionGroup = context.application.getActiveSection().parentSectionGroup;
            
    // Queue a command to load the section group. 
    // For best performance, request specific properties.           
    sectionGroup.load("id,name");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Write the properties.
            console.log("Section group name: " + sectionGroup.name);
            console.log("Section group ID: " + sectionGroup.id);
            
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**name и notebook**
```js
OneNote.run(function (context) {
        
    // Get the parent section group that contains the current section.
    var sectionGroup = context.application.getActiveSection().parentSectionGroup;
            
    // Queue a command to load the section group with the specified properties.           
    sectionGroup.load("name,notebook/name"); 
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Write the properties.
            console.log("Section group name: " + sectionGroup.name);
            console.log("Parent notebook name: " + sectionGroup.notebook.name);
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

    // Get the section groups that are direct children of the current notebook.
    var sectionGroups = context.application.getActiveNotebook().sectionGroups;

    // Queue a command to load the section groups.
    // For best performance, request specific properties.
    sectionGroups.load("name");
    
    // Get the child section groups of the first section group in the notebook.
    var nestedSectionGroups = sectionGroups._GetItem(0).sectionGroups;
    
    // Queue a command to load the ID and name properties of the child section groups.
    nestedSectionGroups.load("id,name");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            
            // Write the properties for each child section group.
            $.each(nestedSectionGroups.items, function(index, sectionGroup) {
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

**sections**
```js
OneNote.run(function (context) {

    // Get the sections that are siblings of the current section.
    var sections = context.application.getActiveSection().parentSectionGroup.sections;

    // Queue a command to load the section groups.
    // For best performance, request specific properties.
    sections.load("id,name");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            
            // Write the properties for each section.
            $.each(sections.items, function(index, section) {
                console.log("Section name: " + section.name);  
                console.log("Section ID: " + section.id);  
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

