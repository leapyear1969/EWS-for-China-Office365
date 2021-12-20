## Exchange Web Services(EWS) 
Some EWS management API samples for China Office 365(``21vianet Office 365``). 
- Authentication of China Office 365
- Make EWS calls, like get mailbox folders, get calendar information, get appointments.

更多功能，陆续施工中。

## 使用方法：

### 授权认证
下载本项目之后，使用Visual Studio打开，打开之后在解决方案管理器中找到``App.config``配置文件，配置对应的``appID``,``tenantID``,``clientSecret``。<br><br>
![image](https://user-images.githubusercontent.com/18607988/146703976-fe3a921e-d604-4077-a6e0-69f589a80759.png)<br>
配置应用信息：
![a](https://user-images.githubusercontent.com/18607988/146703763-fa6471a4-89c5-4490-96a5-7f354393e572.png)<br>

> ``appID``,``tenantID``,``clientSecret``的设置参考文档：[注册应用](https://docs.microsoft.com/en-us/graph/auth-register-app-v2?context=graph%2Fapi%2F1.0&view=graph-rest-1.0)

### 发出EWS请求
在授权成功之后，就通过EWS向Office 365进行数据的请求了，我这边写了3个例子，
- 获取当前账户的邮箱中的所有文件夹
- 获取日历信息包括日历ID和时间
- 获取日历的资源信息
![image](https://user-images.githubusercontent.com/18607988/146706485-f6eb0db7-e27f-413c-bd8e-ac32f2a76573.png)

更多EWS接口请参考如下：
[Get started with EWS Managed API](https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications)
