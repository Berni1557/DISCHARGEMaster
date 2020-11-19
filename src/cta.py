import pandas as pd
from sqlalchemy import create_engine
import time

# list of the tables necessary and the corresponding tablenames in mysql
# path_name_list = [
#                     ('discharge_tags_28022020.xlsx', 'discharge'),
#                     ('discharge_master_20200501.xlsx', 'master_huhu'),
#                     ('ITT.xlsx','itt'),
#                     ('phase_exclude_stenosis.xlsx','phase_exclude_stenosis'),
#                     ('stenosis_bigger_20_phases.xlsx','stenosis_bigger_20_phases'),
#                     ('prct.xlsx','prct'),
#                     ('ecrf.xlsx','ecrf')
#                     ]

# Name of the mysql database to be created
DATABASE = 'tests'
# path of the mysql scripts used
#SQL_SCRIPTS = ['ecrf_queries.sql', 'connect_dicom.sql'] 
SQL_SCRIPTS = ['ecrf_queries.sql', 'connect_dicom.sql'] 
#mysql path, see sqlalchemy create_engine documentation
#MYSQL_ENGINE = 'mysql://root:password@localhost/?charset=utf8'

#MYSQL_ENGINE = 'mysql://root:Lmbjbhj6ich123!M@localhost/?charset=utf8'
#MYSQL_ENGINE = 'mysql://root:Lmbjbhj6ich123!M@DESKTOP-458KDQP'
#MYSQL_ENGINE = 'mysql://DESKTOP-458KDQP:Lmbjbhj6ich123!M@localhost'
MYSQL_ENGINE = 'mysql://root:password@localhost/?charset=utf8'
#MYSQL_ENGINE = 'mysql://root:password@localhost/?charset=UTF8MB4'

# Path of the excel that will be created
EXCEL_OUT = 'master_new.xlsx'


def replace_table(path,name,database, engine):
    """
    path -- filepath of the excel table that should be replaced
    name of the table to be replaced
    database in mysql thats used
    engine --  sqlalchemy engine
    """
    with engine.connect() as con:
        con.execute("use " + database + "; drop table if exists " + name + ";")
    df = pd.read_excel(path)
    df.to_sql(name, engine, index=False)

def execute_mysql_script(path, engine, database):
    """
    path -- path of the mysql script

    """
    # prepare script so it is one line
    temp = open(path, 'r').read().splitlines()
    command = ''
    for i in temp:
        command += i
    
    # Create Engine and execute command
    with engine.connect() as con:
        rs = con.execute("use " + database + ";" + command)
    print(type(rs))

def save_table(table_name, database,excelpath, engine):
    """
    table_name -- name of the table in mysql including the database used
    excelpath -- path where to store the excel table
    engine -- sqlalchemy engine
    """
    df = pd.read_sql_table(table_name, engine, schema=database)
    df.to_excel(excelpath)

def update_table(path_name_list, database, script_path_list, mysql_path, excelpath):
    """
    path_name_list -- list of tuples with path of excel table and corresponding tablename in mysql
    database -- database in mysql to be used
    mysqlpath -- dialext+driver://username:password@host:port/?charset=utf8 e.g.'mysql://root:1234@localhost/?charset=utf8'
    """
    
    path_name_list = path_name_list
    database = DATABASE
    script_path_list = SQL_SCRIPTS
    mysql_path = MYSQL_ENGINE
    excelpath = EXCEL_OUT

    engine = create_engine(mysql_path, encoding="utf-8", echo=False)
    #engine = create_engine('sqlite:///H:\\cloud\\cloud_data\\Projects\\CACSFilter\\src\\christian\\sqlite\\tests.db')
    with engine.connect() as con:
        con.execute('create database if not exists ' + database + ';')
    
    for path, name in path_name_list:
        print(path, name)
        replace_table(path, name, database, engine)
        time.sleep(5)
        
    for script in script_path_list:
        execute_mysql_script(script, engine, database)
        time.sleep(5)
        
        
    #df = pd.read_sql_table('master_plus_ecrf', engine, schema=database)
    
    #return df

    #save_table('master_plus_ecrf',database,excelpath,engine)
    

    has_table = False
    while not has_table:
        time.sleep(10)
        has_table = engine.has_table('master_plus_ecrf', schema=database)
        if has_table:
            df = pd.read_sql_table('master_plus_ecrf', engine, schema=database)
        else:
            print('Waiting for table creation.')

    return df

 
    
#update_table(path_name_list,DATABASE,SQL_SCRIPTS ,MYSQL_ENGINE,EXCEL_OUT)
