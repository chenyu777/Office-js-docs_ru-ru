

# diagnostics

## [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics

Предоставляет надстройке Outlook диагностические сведения.

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|

### Элементы

####  hostName :String

Получает строку, представляющую имя ведущего приложения.

Строка, которая может иметь одно из следующих значений: `Outlook`, `Mac Outlook` или `OutlookWebApp`.

##### Тип:

*   String

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|
####  hostVersion :String

Получает строку, которая представляет версию ведущего приложения или Exchange Server.

Если надстройка почты работает в клиенте Outlook для настольных ПК, свойство `hostVersion` возвращает версию Outlook как ведущего приложения. В Outlook Web App это свойство возвращает версию Exchange Server, например строку `15.0.468.0`.

##### Тип:

*   String

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|
####  OWAView :String

Получает строку, отображающую текущее представление Outlook Web App.

Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.

Если Outlook Web App — не ведущее приложение, при получении доступа к этому свойству будет выдаваться значение `undefined`.

Outlook Web App включает три представления, которые соответствуют ширине экрана и окна, а также числу отображаемых столбцов.

*   `OneColumn` используется в случае узкого экрана: Outlook Web App использует этот макет размером в один столбец на экране смартфона.
*   `TwoColumns` используется при более широком экране: Outlook Web App использует это представление на большинстве планшетных ПК.
*   `ThreeColumns` используется для полноразмерных экранов. Например, Outlook Web App использует это представление в полноэкранном режиме на настольных компьютерах.

##### Тип:

*   String

##### Требования

|Requirement| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](./tutorial-api-requirement-sets.md)| 1.0|
|[Минимальный уровень разрешений](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Применимый режим Outlook| Создание или чтение|
