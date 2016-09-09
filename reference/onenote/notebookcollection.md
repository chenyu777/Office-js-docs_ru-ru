# Объект NotebookCollection (API JavaScript для OneNote)

_Относится к: OneNote Online_  


Представляет коллекцию записных книжек.

## Свойства

| Свойство     | Тип   |Описание|Отзывы|
|:---------------|:--------|:----------|:-------|
|count|int|Возвращает количество записных книжек в коллекции. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-count)|
|items|[Notebook[]](notebook.md)|Коллекция объектов Notebook. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-items)|

_См. [примеры](#примеры) доступа к свойствам._

## Связи
Нет


## Методы

| Метод           | Возвращаемый тип    |Описание| Отзыв|
|:---------------|:--------|:----------|:-------|
|[getByName(name: string)](#getbynamename-string)|[NotebookCollection](notebookcollection.md)|Получает коллекцию записных книжек с указанным именем, открытую в экземпляре приложения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-getByName)|
|[getItem(index: числовой тип или string)](#getitemindex-number-or-string)|[Записная книжка](notebook.md)|Получает записную книжку по ее идентификатору или индексу в коллекции. Только для чтения.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[Записная книжка](notebook.md)|Получает записную книжку по ее позиции в коллекции.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Заполняет прокси-объект, созданный в слое JavaScript, значениями свойства и объекта, указанными в параметре.|[Перейти](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-notebookCollection-load)|

## Сведения о методе


### getByName(name: string)
Получает коллекцию записных книжек с указанным именем, открытую в экземпляре приложения.

#### Синтаксис
```js
notebookCollectionObject.getByName(name);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|name|string|Имя записной книжки.|

#### Возвращаемое значение
[NotebookCollection](notebookcollection.md)

#### Примеры
```js
OneNote.run(function (context) {

    // Get the notebooks that are open in the application instance and have the specified name.
    var notebooks = context.application.notebooks.getByName("Homework");

    // Queue a command to load the notebooks. 
    // For best performance, request specific properties.           
    notebooks.load("id,name");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Iterate through the collection or access items individually by index, for example: notebooks.items[0]
            if (notebooks.items.length > 0) {
                console.log("Notebook name: " + notebooks.items[0].name);
                console.log("Notebook ID: " + notebooks.items[0].id);
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
Получает записную книжку по ее идентификатору или индексу в коллекции. Только для чтения.

#### Синтаксис
```js
notebookCollectionObject.getItem(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number или string|Идентификатор записной книжки или расположение индекса записной книжки в коллекции.|

#### Возвращаемое значение
[Записная книжка](notebook.md)

### getItemAt(index: number)
Получает записную книжку по ее позиции в коллекции.

#### Синтаксис
```js
notebookCollectionObject.getItemAt(index);
```

#### Параметры
| Параметр    | Тип   |Описание|
|:---------------|:--------|:----------|
|index|number|Значение индекса получаемого объекта. Используется нулевой индекс.|

#### Возвращаемое значение
[Записная книжка](notebook.md)

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

    // Get the notebooks that are open in the application instance and have the specified name.
    var notebooks = context.application.notebooks.getByName("Homework");

    // Queue a command to load the notebooks. 
    // For best performance, request specific properties.           
    notebooks.load("id");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Iterate through the collection or access items individually by index, for example: notebooks.items[0]
            $.each(notebooks.items, function(index, notebook) {
                notebook.addSection("Biology");
                notebook.addSection("Spanish");
                notebook.addSection("Computer Science");
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

