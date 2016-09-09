
# Labs.Core.IConnectionResponse

 _**Область применения:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Сведения об ответе, возвращаемые при вызове подключения.

```
interface IConnectionResponse
```


## Свойства


|||
|:-----|:-----|
| `initializationInfo: Core.IConfigurationInfo`|Сведения о конфигурации инициализации (либо значение **null**, если приложение не было инициализировано).|
| `mode: Core.LabMode`|Режим, в котором на данный момент работает лаборатория.|
| `hostVersion: Core.IVersion`|Сведения о версии ([Labs.Core.IVersion](../../reference/office-mix/labs.core.iversion.md)) для сервера.|
| `userInfo: Core.IUserInfo`|Сведения о пользователе ([Labs.Core.IUserInfo](../../reference/office-mix/labs.core.iuserinfo.md)).|
