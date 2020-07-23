from flask import Flask, flash, redirect, url_for, render_template, request, send_file, Response
from werkzeug.utils import secure_filename
import collections
import pandas as pd
import xlsxwriter
import os
import json
from datetime import datetime, time
import psycopg2

"""
maquinaria_excel = pd.read_excel('plantilla_carga_datos.xlsx',
                                sheet_name='maquinaria_tipo',
                                header=0)

#print("maquinaria excel")
#print(maquinaria_excel)
#Conexión a db

conn = sqlite3.connect('elecnor.db')
cur = conn.cursor()

maquinaria_bd = pd.read_sql("SELECT * FROM maquinaria_tipo", conn)

#print("maquinaria bd")
#print(maquinaria_bd)
#maquinaria_tipo.to_sql('maquinaria_tipo', conn, if_exists = 'replace', index=False)
#sss = pd.read_sql("SELECT * FROM maquinaria_tipo LIMIT 10", conn)
conn.close()

for index, nuevo_dato in maquinaria_excel.iterrows():
    excel_id = nuevo_dato['subir']
    print(excel_id)



# create table maquinaria_tipo(id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, nombre varchar(120), tiene_patente bool, comentarios text, activo bool)

maquinaria_tipo = pd.read_excel('plantilla_carga_datos.xlsx',
                                sheet_name='maquinaria_tipo',
                                header=0,
                                dtype={'nombre': str,
                                       'tiene_patente': bool,
                                       'comentarios': str,
                                       'activo': bool,
                                       'subir': bool})

del maquinaria_tipo['subir']

maquinaria_tipo.index +=1

print(maquinaria_tipo)

#Conexión a db

conn = sqlite3.connect('elecnor.db')
cur = conn.cursor()

nnn = pd.read_sql("SELECT * FROM maquinaria_tipo", conn)

print("Sin importacion")
print(nnn)

maquinaria_tipo.to_sql('maquinaria_tipo', conn, if_exists='append', index=False)

sss = pd.read_sql("SELECT * FROM maquinaria_tipo", conn)

print("Con importacion")
print(sss)

conn.close()


#Conexión a db

conn = sqlite3.connect('elecnor.db')
cur = conn.cursor()

df = pd.read_sql("SELECT * FROM maquinaria_tipo", conn)

del df['id']

df.index +=1

print(df)

conn.close()

#Xlsx write

writer = pd.ExcelWriter('nueva_planilla.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='maquinaria_tipo')

writer.save() """

def db():
    try:
        con = psycopg2.connect(database='defaultdb',
                                user='doadmin',
                                password='eyepgge23jcpydtl',
                                host='db-postgresql-innspatial-do-user-7628107-0.a.db.ondigitalocean.com',
                                port='25060')
        print('Conexión exitosa a la base de datos')
        #cur = db().cursor()
        #cur.execute('SELECT * FROM maq_tipo;')
        return con
    except:
        print('La conexión ha fallado')

#funcionalidades
def excel_to_db(new_excel_name, db_table='maq_tipo'):
# new_excel_name type: str
# db_table type: str

#Connection to db
    conn = db()
    cur = conn.cursor()
#Excel to dataframe
    raw_data = pd.read_excel(new_excel_name,
                                    sheet_name='maquinaria_tipo',
                                    header=0
                                    )

    nueva_entrada = raw_data[raw_data['subir']==True]

    del nueva_entrada['subir']

    nueva_entrada.index +=1

    print('Datos a actualizar \n')
    print(nueva_entrada)

    print('\n Base de Datos sin actualizar \n')
    dbdb = pd.read_sql("SELECT * FROM maq_tipo", conn)
    print(dbdb)

#Update information
    try:
        for fila in nueva_entrada.itertuples(name='Dato'):
            if str(fila.nombre)=='nan':
                nom = None
            else:
                nom = str(fila.nombre)

            if str(fila.tiene_patente)=='nan':
                tp = None
            else:
                tp = fila.tiene_patente

            if str(fila.comentarios)=='nan':
                com = None
            else:
                com = str(fila.comentarios)

            if str(fila.activo)=='nan':
                act = None
            else:
                act = fila.activo

            number = int(fila.Index)

            #print(f"UPDATE {db_table} SET nombre = {nom}, tiene_patente = {tp}, comentarios = {com}, activo = {act} WHERE id = {int(fila.Index)} \n")
            if number in dbdb['id']:
                cur.execute("UPDATE maq_tipo SET nombre = %s, tiene_patente = %s, comentarios = %s, activo = %s WHERE id = %s", (nom, tp, com, act, number))
            else:
                cur.execute("INSERT INTO maq_tipo VALUES(%s,%s,%s,%s,%s)",(number, nom, tp, com, act))
    except:
        print("Parese k no funciona la wa")

    dbdb2 = pd.read_sql("SELECT * FROM maq_tipo", conn)

    print('\n Base de datos ACTUALIZADA \n')
    print(dbdb2)

    conn.commit()
    conn.close()

excel_to_db('plantilla_carga_datos.xlsx')
