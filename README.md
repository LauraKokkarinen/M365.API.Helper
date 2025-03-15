# M365.API.Helper

C# .NET 8 class library that contains all the necessary code for calling the Microsoft Graph and SharePoint Online REST APIs. The code is as minimalistic as possible while still containing support for paging and throttling. Offers easy extensibility for the developer. When running on Azure, the authentication happens using the resource's managed identity. While debugging, a client secret will be used for authenticating to Microsoft Graph, and a client certificate will be used for authenticating to the SharePoint Online REST API.
