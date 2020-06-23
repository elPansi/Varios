# -*- coding: utf-8 -*-
import cx_Oracle
import pypyodbc
import pymssql 
import MySQLdb
import sqlite3
import fdb


def consultaOracle(query=''):
    conn =  cx_Oracle.connect('usuario/password@192.168.5.252:1521/orcl')
    cursor = conn.cursor()         # Crear un cursor
    cursor.execute(query)          # Ejecutar una consulta

    if query.upper().startswith('SELECT'):
        data = cursor.fetchall()   # Traer los resultados de un select
    else:
        conn.commit()              # Hacer efectiva la escritura de datos
        data = None

    cursor.close()                 # Cerrar el cursor
    conn.close()                   # Cerrar la conexión
    return(data)

def consultaMysql(query=''):
   DB_HOST = '127.0.0.1'
   DB_USER = 'root'
   DB_PASS = 'passUsuario'
   DB_NAME = 'nombreBD'
   datos = [DB_HOST, DB_USER, DB_PASS, DB_NAME]

   conn = MySQLdb.connect(*datos) # Conectar a la base de datos
   cursor = conn.cursor()         # Crear un cursor
   cursor.execute(query)          # Ejecutar una consulta

   if query.upper().startswith('SELECT'):
       data = cursor.fetchall()   # Traer los resultados de un select
   else:
       conn.commit()              # Hacer efectiva la escritura de datos
       data = None

   ##f.write("Cons. MySQL devuelve: " + str(data) +"\r\n")
   cursor.close()                 # Cerrar el cursor
   conn.close()                   # Cerrar la conexión
   return(data)

def consultaMSSQL(query=''):
    servidor = '192.168.2.6'
    usuario = 'usuario'
    password = 'password'
    bd = 'nombreBD'
    driver = '{ODBC Driver 13 for SQL Server}'
    #conn = pypyodbc.connect(driver=driver, server=servidor, database=bd, uid=usuario, pwd=password ,Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;)
    conn = pypyodbc.connect(driver=driver, server=servidor, database=bd, uid=usuario, pwd=password)
    cursor = conn.cursor()
    cursor.execute(query)
    cabeceras = []

    for c in cursor.description:
        print(c)
        cabeceras.append(c)

    if query.upper().startswith('SELECT'):
        data = cursor.fetchall()   # Traer los resultados de un select
    else:
        conn.commit()              # Hacer efectiva la escritura de datos
        data = None

    cursor.close()
    conn.close()
    return(data, cabeceras)

def consultaSQLite(query=''): 
    conn = sqlite3.connect('datos.sqlite')
    cursor = conn.cursor()         # Crear un cursor
    cursor.execute(query)          # Ejecutar una consulta

    if query.upper().startswith('SELECT'):
        data = cursor.fetchall()   # Traer los resultados de un select
    else:
        conn.commit()              # Hacer efectiva la escritura de datos
        data = None

    cursor.close()
    conn.close()
    return(data)
	
def consultaFB(query=''):
    servidor = '172.26.0.104'
    usuario = 'sysdba'
    passw = 'passwordBD'
    bd = 'c:\\bdReloj\\tadata.fdb'
    conn = fdb.connect(host=servidor, database=bd, user=usuario, password=passw) #, charset='UTF8'
    cursor = conn.cursor()
    cursor.execute(query)
	
    if query.upper().startswith('SELECT'):
        data = cursor.fetchall()
    else:
        conn.commit()
        data = None
    cursor.close()
    conn.close()
    return(data)
