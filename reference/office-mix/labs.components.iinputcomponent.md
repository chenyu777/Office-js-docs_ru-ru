
# Labs.Components.IInputComponent

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Позволяет взаимодействовать с компонентом ввода.

```
interface IInputComponent extends Labs.Core.IComponent
```


## Свойства


|Имя|Описание|
|:-----|:-----|
| `maxScore: number`|Максимальный допустимый показатель для компонента ввода.|
| `timeLimit: number`|Предельное время для проблемы ввода.|
| `hasAnswer: boolean`|Возвращает значение **True**, если у компонента есть ответ.|
| `answer: any`|Ответ на проблему компонента (при наличии).|
| `secure: boolean`|Возвращает значение **True**, если компонент ввода защищен.|
