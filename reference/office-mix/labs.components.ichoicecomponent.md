
# Labs.Components.IChoiceComponent

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Позволяет взаимодействовать с компонентом выбора.

```
interface IChoiceComponent extends Labs.Core.IComponent
```


## Свойства


|Имя|Описание|
|:-----|:-----|
| `choices: Components.IChoice[]`|Массив, представляющий список вариантов выбора, связанных с проблемой.|
| `timeLimit: number`|Предельное время решения проблемы.|
| `maxAttempts: number`|Максимальное число попыток, разрешенных для проблемы.|
| `maxScore: number`|Максимальный показатель проблемы.|
| `hasAnswer: boolean`|Возвращает значение **True**, если для проблемы есть ответ.|
| `answer: any`|Ответ для проблемы. Массив, если поддерживается несколько ответов, либо один идентификатор, если поддерживается только один ответ.|
| `secure: boolean`|Показывает, является ли тест защищенным, то есть скрыты ли от пользователя защищенные поля.|
