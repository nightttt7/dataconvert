# esp 
**python2.7**
- 让Excel,MySQL和pandas的数据可以快速相互转换
- 主要通过pandas对数据进行处理和可视化
- 适合因为各种因素,必须要有excel格式的数据,但是使用pandas和数据库可以大幅提高工作效率的情况

### 需要使用固定格式的Excel表格,格式如下
1. sheet名  
'产品信息','分类1','分类2'...
2. 产品数据表  
![](./jpg/6f51cb25cb25453795d5d73fdcd86174.jpg)  
*index不能为字符串,推荐日期或整数*  
3. 产品信息表  
![](./jpg/82ecfcfd71bb4e8788c2fa1978840166.jpg)  
### sql安排
所有数据在同一个database中  
有一个名为'产品信息'(可自定义命名)的信息汇总表  
其余为每个品种的数据表  
### pandas数据
1. product_data  
key为string,value为pandas.DataFrame的dict
2. product_panel  
pandas的面板数据,三维分别是产品名,日期和指标
3. product_data_multi  
pandas的多重索引DataFrame,index是指标和日期,columns是产品名
4. product_information  
包含'产品信息'(可自定义命名)的pandas.DataFrame
## 主要函数
### MyProduct.excel_to_frame(path,sheets,header,index_col,pruduct_name_col,data_breaks,header_info,index_col_info,sheet_info)  
返回值(两个):  
product_data,是一个key是string,value是pandas.DataFrame的dict  
product_info,是一个pandas.DataFrame  
必要参数:  
path : string  
需要导入的excel文件路径  
可选参数:  
sheets : [string,string,...]  
指定需要读取的sheet名,默认为所有sheet除掉产品信息表  
header : int default 1  
数据表的数据部分表头所在行(columns所在的那一行),这里及以下都是从0开始数数的  
index_col : int default 2  
数据表的数据部分第一个index所在的列  
pruduct_name_col : int default 0  
数据表的产品名列所在的列  
data_breaks : int default 1  
数据表的数据部分的每个小表格的间隔列  
header_info : int default 0  
产品信息表的表头所在行  
index_col_info : int default 0  
产品信息表的index所在的列  
sheet_info : string default u'产品信息'  
产品信息表的名字  

### MyProduct.sql_to_frame(login,sheet_info,index_name)
返回值(两个):  
product_data,是一个key是string,value是pandas.DataFrame的dict  
product_info,是一个pandas.DataFrame  
必要参数:  
login: string  
sql的登录信息,形如"mysql+mysqldb://user:password@hostname/databasename?charset=utf8"  
可选参数:  
sheet_info : string default u'产品信息'  
产品信息表的名字  
index_name : string default u'日期'  
index的名字  

### MyProduct.frame_to_sql(data,info,login,sheet_info)
返回值:无,仅将product_data和product_information导入sql,已有的table则覆盖
必要参数:  
data : dict  
是一个key是string,value是pandas.DataFrame的dict,一般由excel_to_frame或者sql_to_frame生成  
info : pandas.DataFrame  
产品信息的pandas.DataFrame,一般由excel_to_frame或者sql_to_frame生成  
login: string  
sql的登录信息,形如"mysql+mysqldb://user:password@hostname/databasename?charset=utf8"  
可选参数:  
sheet_info : string default u'产品信息'  
产品信息表的名字  

### MyProduct.sql_to_excel(login,path,header,index_col,pruduct_name_col,sheet_info,kind_name,data_name)
返回值:无,仅将sql中的数据表和产品信息表导入excel,已有则覆盖  
必要参数:  
login: string  
sql的登录信息,形如"mysql+mysqldb://user:password@hostname/databasename?charset=utf8"  
path : string  
需要导入的excel文件路径  
可选参数:  
header : int default 1  
数据表的数据部分表头所在行(columns所在的那一行),这里及以下都是从0开始数数的  
index_col : int default 2  
数据表的数据部分第一个index所在的列  
pruduct_name_col : int default 0  
数据表的产品名列所在的列  
sheet_info : string default u'产品信息'  
产品信息表的名字  
kind_name : string default u'合作方式'  
分类的依据,应为产品信息表中的列名  
data_name : string default u'产品'  
每个小表的名字在产品信息表中的列名  

## 内部函数
### sql_to_frame1(login,sheet_info)  
与sql_to_frame的区别在于不指定某列为index,少了index_name参数  
返回值(两个):  
product_data,是一个key是string,value是pandas.DataFrame的dict  
product_info,是一个pandas.DataFrame  
必要参数:  
login: string  
sql的登录信息,形如"mysql+mysqldb://user:password@hostname/databasename?charset=utf8"  
可选参数:  
sheet_info : string default u'产品信息'  
产品信息表的名字  

### frame1_to_excel(data,info,path,header,index_col,pruduct_name_col,sheet_info,kind_name,data_name)
返回值:无,仅将product_data和product_information导入excel,已有则覆盖  
必要参数:  
data : dict  
是一个key是string,value是pandas.DataFrame的dict,一般由excel_to_frame或者sql_to_frame生成  
info : pandas.DataFrame  
产品信息的pandas.DataFrame,一般由excel_to_frame或者sql_to_frame生成  
path : string  
需要导入的excel文件路径  
可选参数   
header : int default 1  
数据表的数据部分表头所在行(columns所在的那一行),这里及以下都是从0开始数数的  
index_col : int default 2  
数据表的数据部分第一个index所在的列  
pruduct_name_col : int default 0  
数据表的产品名列所在的列  
sheet_info : string default u'产品信息'  
产品信息表的名字
kind_name : string default u'合作方式'
分类的依据,应为产品信息表中的列名
data_name : string default u'产品'
每个小表的名字在产品信息表中的列名

## 如何使用

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



