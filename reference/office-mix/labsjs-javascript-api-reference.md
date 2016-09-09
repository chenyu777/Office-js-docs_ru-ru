
# Справка по API JavaScript библиотеки LabsJS
Обзор объектной модели JavaScript LabsJS.

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Справка по LabsJS содержит ссылку на файл [TypeScript](http://www.typescriptlang.org/), **labs-1.0.42.d.ts**, который делит объектную модель LabsJS на модули.

## Объектная модель LabsJS

Объектная модель LabsJS состоит из пяти модулей:


- [LabsJS.Labs](../../reference/office-mix/labsjs.labs.md). Модуль лабораторий содержит набор основных API, с помощью которых непосредственно создаются лаборатории. Они обеспечивают точки входа для разработки лабораторий.
    
- [LabsJS.Labs.Core](../../reference/office-mix/labsjs.labs.core.md). Основные интерфейсы, структуры данных и классы, используемые библиотекой LabsJS и драйвером презентаций (в данном случае — Office Mix) в качестве общего канала связи.
    
- [LabsJS.Labs.Core.Actions](../../reference/office-mix/labsjs.labs.core.actions.md). Эти API представляют действия лабораторий, указывая на их текущее поведение. Они полезны разработчикам, которые создают компоненты (кроме компонентов по умолчанию) или разрабатывают связи с новым драйвером (кроме Office Mix).
    
- [LabsJS.Labs.Core.GetActions](../../reference/office-mix/labsjs.labs.core.getactions.md). Эти API позволяют запрашивать действия, ранее выполненные на сервере.
    
- [LabsJS.Labs.Components](../../reference/office-mix/labsjs.labs.components.md). Эти API представляют четыре типа компонентов, доступные для лабораторий в настоящее время (компоненты действия (Activity), компоненты выбора (Choice), компоненты ввода (Input) и динамические компоненты (Dynamic)).
    
Каждый модуль включает в себя набор элементов, состоящих из одного или нескольких типов:


- Классы
    
- Интерфейсы
    
- Функции
    
- Перечисления
    
- Переменные
    



## Дополнительные ресурсы



- [TypeScript](http://www.typescriptlang.org/)
    
