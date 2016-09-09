# Объект OfficeExtension.Error (API JavaScript для Excel)

Представляет ошибки, которые возникают при использовании API JavaScript для Excel.

_Относится к: Excel 2016, Excel Online, Excel для iOS, Office 2016_

## Свойства
| Свойство     | Тип   |Описание
|:---------------|:--------|:----------|
|code|string|Возвращает тип ошибки. Возможные значения: AccessDenied, ActivityLimitReached, BadPassword, GeneralException, InsertDeleteConflict, InvalidArgument, InvalidBinding, InvalidOperation, InvalidReference, InvalidSelection, ItemAlreadyExists, ItemNotFound, NotImplemented и UnsupportedOperation. |
|debugInfo|string|Возвращает значение, которое указывает, что произошло при возникновении ошибки. Это значение предназначено для использования только во время разработки и отладки.  |
|сообщение |string| Возвращает локализованную понятную для пользователя строку, которая соответствует коду ошибки.|
|name |string| Возвращает значение OfficeExtension.Error. |
|traceMessages |string[]| Возвращает массив значений, которые соответствуют сообщениям инструментирования, заданным с помощью синтаксиса context.trace(); |

## Методы

| Метод           | Возвращаемый тип    |Описание|
|:---------------|:--------|:----------|
|[toString()](#tostring)|строка|Возвращает код ошибки и сообщение в следующем формате: "{0}: {1}", код, сообщение.|

## Сведения о методе

### toString()
Возвращает код ошибки и сообщение в следующем формате: "{0}: {1}", код, сообщение.

#### Синтаксис
```js
error.toString()
```

#### Параметры
Нет

#### Возвращаемое значение
string
