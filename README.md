# python-ssas

### Motivation:
I’ve been working for some time analyzing data using two unrelated tools – Python (primarily pandas), and DAX in Microsoft’s Tabular models. They are very different – Python is open source, Microsoft’s products are (obviously) not. It was frustrating to not be able to merge them. If I wanted to get data from a DAX data model into a Pandas dataframe, I would typically need to first export it to a file (like CSV) and then read it from there.

Also, I’ve been building a data mart using the Microsoft platform, but on a shoestring budget. Notably, I don’t have a SQL Server license. So part of what I’m missing is a tool to automate tasks like SQL Server Agent. 

### Solution:
Inspired by [@akavalar's great post](https://github.com/akavalar/SSAS-on-a-shoestring), I discovered a nice workaround:
- DAX models (or any Analysis Services model) have two .Net APIs:  
  - [ADOMD.NET](https://msdn.microsoft.com/en-us/library/mt465769.aspx) for admin tasks like refreshing the model 
  - [AMO](https://msdn.microsoft.com/en-us/library/mt436122.aspx) to query it.
- Also, there is a fantastic Python library called [Pythonnet](https://github.com/pythonnet/pythonnet) that enables near seamless integration between CPython (mainstream, not Iron Python) and .Net. 

Using these ingredients, I created my `ssas_api.py` module with some utilities that I use frequently. Note that this just uses the parts of the APIs that I needed; there is a wealth more available, just dig through the documentation.

**Note:** I've only been using Azure Analysis Services, so the code is designed for that regarding the URL of the server and authentication string.

I haven't found anything like this online, so feel free to use it.
