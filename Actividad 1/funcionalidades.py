import psycopg2
import sqlite3
import pandas as pd
import collections
import xlsxwriter
import json
from datetime import datetime, time


def excel_to_db(new_excel_name, db_table='maquinaria_tipo'):
# new_excel_name type: str
# db_table type: str

#Connection to db
    conn = sqlite3.connect('elecnor.db')
    cur = conn.cursor()

#Excel to dataframe
    raw_data = pd.read_excel(new_excel_name,
                                    sheet_name='maquinaria_tipo',
                                    header=0,
                                    dtype={'nombre': str,
                                           'tiene_patente': bool,
                                           'comentarios': str,
                                           'activo': bool,
                                           'subir': bool})

    nueva_entrada = raw_data[raw_data['subir']==True]

    del nueva_entrada['subir']

    nueva_entrada.index +=1

    print('Datos a actualizar \n')
    print(nueva_entrada)

    print('\n Base de Datos sin actualizar \n')
    dbdb = pd.read_sql("SELECT * FROM maquinaria_tipo", conn)
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

            cur.execute("UPDATE maquinaria_tipo SET nombre = ?, tiene_patente = ?, comentarios = ?, activo = ? WHERE id = ?", (nom, tp, com, act, number))
    except:
        print("Parese k no funciona la wa")


    dbdb2 = pd.read_sql("SELECT * FROM maquinaria_tipo", conn)

    print('\n Base de datos ACTUALIZADA \n')
    print(dbdb2)

    conn.commit()

    conn.close()

def db_to_excel(database='elecnor.db'):
#Database: 'elecnor.db'

#Conexión a db
    conn = sqlite3.connect('elecnor.db')
    cur = conn.cursor()

    df = pd.read_sql("SELECT * FROM maquinaria_tipo", conn)

    print(df)

    conn.close()

#Modificación Dataframe
    del df['id']

    replace_values = {0:'false', 1:'true'}

    df = df.replace({"tiene_patente": replace_values})
    df = df.replace({"activo": replace_values})
    df['subir'] = 'false'

    print(df)

#Xlsx write
    writer = pd.ExcelWriter('nueva_planilla.xlsx', engine='xlsxwriter')

    df.to_excel(writer, sheet_name='maquinaria_tipo', index=False)

    writer.save()

def get_mq_tipo( json_str = False ):
    #db to json
    conn = sqlite3.connect('elecnor.db')
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    rows = cur.execute('''SELECT * FROM maquinaria_tipo''').fetchall()

    conn.commit()
    conn.close()

    if json_str:
        return json.dumps([dict(ix) for ix in rows], indent=4) #CREATE JSON

    return rows

t = get_mq_tipo( json_str = True )

print(t)
