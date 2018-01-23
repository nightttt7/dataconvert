# esp python2.7
- transform excel/sql/pandas to each other and analyze product data or similar
- sometimes you must use excel, but you like pandas and sql more, use esp

### excel layout 
1. sheet names  
'info','cat1','cat2'...
2. datas  
![](./jpg/120180123155746.png)  
*index can't be string, use date or int*  
3. info  
![](./jpg/220180123160111.png)  
### sql layout
all data in one database  
one table name 'info'(can rename)  
others are tables for each kind  
### pandas data
1. product_data  
dict which key is string and value is pandas.DataFrame
2. product_panel  
pandas.panel,index are name,date and index
3. product_data_multi  
pandas's multi-DataFrame,index is index and data,columns is name
4. product_information  
pandas.DataFrame including 'info' (can rename)
## Main functions
### MyProduct.excel_to_frame(path,sheets,header,index_col,pruduct_name_col,data_breaks,header_info,index_col_info,sheet_info)  
returns(2):  
dict which key is string and value is pandas.DataFrame  
pandas.DataFrame including 'info' (can rename)  
must parameter:  
path : string  
excel path  
optional parameter:  
sheets : [string,string,...]  
names of sheets need read, default all sheets except info sheet  
header : int default 1  
header of data, count form 0  
index_col : int default 2  
first index column of data  
pruduct_name_col : int default 0  
column of name of data  
data_breaks : int default 1  
the separate columns in data  
header_info : int default 0  
header of info  
index_col_info : int default 0  
column of index of info  
sheet_info : string default u'info'  
name of info sheet  

### MyProduct.sql_to_frame(login,sheet_info,index_name)
returns(2):  
dict which key is string and value is pandas.DataFrame  
pandas.DataFrame including 'info' (can rename)  
must parameter:  
login: string  
sql login string. such as"mysql+mysqldb://user:password@hostname/databasename?charset=utf8"  
optional parameter:  
sheet_info : string default u'info'  
name of info sheet  
index_name : string default u'date'  
name of index  

### MyProduct.frame_to_sql(data,info,login,sheet_info)  
returns:None,only import product_data and product_information to sql, cover the original one  
must parameter:  
data : dict  
dict which key is string and value is pandas.DataFrame, often comes form excel_to_frame or sql_to_frame  
info : pandas.DataFrame  
pandas.DataFrame of info, often comes form excel_to_frame or sql_to_frame  
login: string  
sql login string. such as"mysql+mysqldb://user:password@hostname/databasename?charset=utf8"  
optional parameter:  
sheet_info : string default u'info'  
name of info sheet  

### MyProduct.sql_to_excel(login,path,header,index_col,pruduct_name_col,sheet_info,kind_name,data_name)
returns:None,only import data and info of sql to excel, cover the original one  
must parameter:  
login: string  
sql login string. such as"mysql+mysqldb://user:password@hostname/databasename?charset=utf8"  
path : string  
excel path  
optional parameter:  
header : int default 1  
header of data, count form 0  
index_col : int default 2  
first index column of data  
pruduct_name_col : int default 0  
column of name of data  
sheet_info : string default u'info'  
name of info sheet  
kind_name : string default u'cat'  
the cats name in info sheet  
data_name : string default u'name'  
the name of every data in info sheet  

## Inner function
### sql_to_frame1(login,sheet_info)  
the difference from sql_to_frame is not let one column as the index, and loss the index_name parameter  
returns(2):  
dict which key is string and value is pandas.DataFrame  
pandas.DataFrame including 'info' (can rename)  
must parameter:  
login: string  
sql login string. such as"mysql+mysqldb://user:password@hostname/databasename?charset=utf8"  
optional parameter:  
sheet_info : string default u'info'  
name of info sheet  

### frame1_to_excel(data,info,path,header,index_col,pruduct_name_col,sheet_info,kind_name,data_name)
returns:None,only import product_data and product_information to excel, cover the original one  
must parameter:  
data : dict  
dict which key is string and value is pandas.DataFrame, often comes form excel_to_frame or sql_to_frame  
info : pandas.DataFrame  
pandas.DataFrame of info, often comes form excel_to_frame or sql_to_frame  
path : string  
excel path  
optional parameter   
header : int default 1  
header of data, count form 0 
index_col : int default 2  
first index column of data  
pruduct_name_col : int default 0  
column of name of data  
sheet_info : string default u'info'  
name of info sheet
kind_name : string default u'cat'  
the cats name in info sheet
data_name : string default u'name'  
the name of every data in info sheet  
  
## how to use

```
import MyProduct as mp
path='example/data.xlsx'
login="mysql+mysqldb://user:password@hostname/databasename?charset=utf8"
#excel_to_frame
data,info=mp.excel_to_frame(path=path)
#sql_to_frame
data,info=mp.sql_to_frame(login=login)
#frame_to_sql
mp.frame_to_sql(data=data,info=info,login=login)
#sql_to_excel
mp.sql_to_excel(login=login,path=path)
```



