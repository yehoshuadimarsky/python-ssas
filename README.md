# python-ssas

## Prerequisites:
- Good working knowledge of:
  - Microsoft SQL Server Analysis Services (Tabular models)
  - Python (specifically pandas)
- Some knowledge required:
  - General Microsoft .Net familiarity


### Motivation:
I’ve been working for some time analyzing data using two unrelated tools – Python (primarily pandas), and DAX in Microsoft’s Tabular models. They are very different – Python is open source, Microsoft’s products are (obviously) not. It was frustrating to not be able to merge them. If I wanted to get data from a DAX data model into a Pandas dataframe, I would typically need to first export it to a file (like CSV) and then read it from there.

Also, I wanted a way to programatically "refresh" the data model (called "processing" it) from Python.

### Solution:
Inspired by [@akavalar's great post](https://github.com/akavalar/SSAS-on-a-shoestring), I discovered a nice workaround:
- DAX models (or any Analysis Services model) have several .Net APIs, see [here](https://docs.microsoft.com/en-us/sql/analysis-services/analysis-services-developer-documentation) for the Microsoft documentation
- Also, there is a fantastic Python library called [Pythonnet](https://github.com/pythonnet/pythonnet) that enables near seamless integration between Python and .Net. This is for the mainstream Python, called CPython, and not to be confused with the .Net implementation of Python which is called IronPython.

Using these ingredients, I created my `ssas_api.py` module with some utilities that I use frequently. Note that this just uses the parts of the APIs that I needed; there is a wealth more available, just dig through the documentation.

**Note:** I've only been using Azure Analysis Services, so the code is designed for that regarding the URL of the server and authentication string.

I haven't found anything like this online, so feel free to use it.

## Quickstart
```
In [1]: import ssas_api
   ...: 
   ...: conn = ssas_api.set_conn_string(
   ...:     ssas_server='<YOUR_SERVER>',
   ...:     db_name='<YOUR_DATABASE>',
   ...:     username='<USERNAME>',
   ...:     password='<PASSWORD>'
   ...: )

In [2]: dax_string = '''
   ...: //any valid DAX query
   ...: EVALUATE
   ...: CALCULATETABLE(MyTable)
   ...: '''

In [3]: df = ssas_api.get_DAX(connection_string=conn, dax_string=dax_string)
```
