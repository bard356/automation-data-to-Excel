# нумерация ячеек для интеграции в файлы формата тетрадей .ypinb
# In[1]:


import cx_Oracle
import pandas as pd


# In[2]:


#загружаем справочник имен в виде таблице для создания словарей 
df = pd.read_excel('Справочник.xlsx',sep=';') 

#удаляем строки с пустыми значениями специально для файлов gpc
df_gpc = pd.read_excel('Справочник.xlsx',sep=';')
df_gpc.dropna(subset=['Название excel файла (GPC)'],inplace=True)


# In[3]:


#сохраняем названия в значения и ключ будущих словарей 
val_ppr = df['Название excel файла (ППР)']
val_gpc = df_gpc['Название excel файла (GPC)']

keys_ppr = df['Таблица в oracle']
keys_gpc = df_gpc['Таблица в oracle']

#создаем словарь для имен клиентов 
sl_ppr = dict(zip(keys_ppr,val_ppr))
sl_gpc = dict(zip(keys_gpc,val_gpc))


# In[4]:


#данные для входа
USER = 'name'
PASSWORD = 'password'

#константна path - содержит адрес папки для сохранения файлов
FINAL_PATH ='K:/......./Списки/'


# In[5]:


#функция для запроса таблицы из БД
def oracle_querry(from_table):
    dsn_tns = cx_Oracle.makedsn('dm.ru'
                                ,0000
                                ,service_name='name_user'
                               )
    conn = cx_Oracle.connect(
                            user=USER
                            ,password=PASSWORD
                            ,dsn=dsn_tns
                            )
    query_string = "select * from " + from_table
    #query = ()
    request = pd.read_sql(query_string,conn)
    conn.close()
    return (request)


# In[6]:
print('Конект к БД установлен успешно. Выгружаем')

#передаем в переменную списком исхдные названия таблиц в БД, они у нас являются ключами для PPR
TABLE_LIST = list(sl_ppr.keys())

#создаем результирующий список таблиц result_table, в нем хранятся наши таблицы под индексами []
result_table = []
for table_name in TABLE_LIST:
    ########
    table_name = oracle_querry(table_name) #запрос в БД
    result_table.append(table_name) #добавление в список

#создаем переменную file_name_path с названием каждой таблицы, и передаем новое название с помощью словаря   PPR
for n,table_name in enumerate(TABLE_LIST):
    #строка с путем сохранения файла, названием и форматом
    file_name_path = FINAL_PATH + sl_ppr[table_name] + '.xlsx' 
    #для каждой таблицы по индексу сохраняем без индексов в файл
    result_table[n].to_excel(file_name_path,index=False)  
    
    print('Таблица для PPR {} выгружена'.format(sl_ppr[table_name])) 


# In[7]:


#GPC
TABLE_LIST_GPC = list(sl_gpc.keys())

#создаем результирующий список таблиц result_table, в нем хранятся наши таблицы под индексами []
result_table_gpc = []
for table_name in TABLE_LIST_GPC:
    ########
    table_name = oracle_querry(table_name) #запрос в БД
    result_table_gpc.append(table_name) #добавление в список
    
#создаем переменную file_name_path с названием каждой таблицы, и передаем новое название с помощью словаря      GPC
for n,table_name in enumerate(TABLE_LIST_GPC):
    #строка с путем сохранения файла, названием и форматом
    file_name_path_gpc = FINAL_PATH + sl_gpc[table_name] + '.xlsx' 
    #для каждой таблицы по индексу сохраняем без индексов в файл
    result_table_gpc[n].to_excel(file_name_path_gpc,index=False) 
    
    print('Таблица для GPC {} выгружена'.format(table_name)) 


# In[ ]:




