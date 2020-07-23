from flask import Flask, flash, redirect, url_for, render_template, request, send_file, Response
from werkzeug.utils import secure_filename
import collections
import pandas as pd
import xlsxwriter
import os
import json
from datetime import datetime, time
import psycopg2

# Base Aplicación web
app = Flask(__name__)

UPLOAD_FOLDER = 'static/uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
#app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

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

def desconexion():
    try:
        conn.commit()
        conn.close()
    except:
        print("Esta conexión no existe")

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

def db_to_excel():
#Database: 'elecnor.db'

#Conexión a db
    conn = db()
    cur = conn.cursor()

    df = pd.read_sql("SELECT * FROM maq_tipo", conn)

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

def query_db(query, args=(), one=False):
    cur = db().cursor()
    cur.execute(query, args)
    r = [dict((cur.description[i][0], value) \
               for i, value in enumerate(row)) for row in cur.fetchall()]
    cur.connection.close()
    return (r[0] if r else None) if one else r

def to_json():
    my_query = query_db("select * from maq_tipo")

    json_output = json.dumps(my_query)
    print(json_output)

    return json_output

#decorador inicial
@app.route('/')
def index():
    return "Wena los cabros"

@app.route('/upload')
def upload_file():
   return render_template('upload.html')

@app.route('/uploader', methods = ['GET', 'POST'])
def uploader_file():
   if request.method == 'POST':
      f = request.files['file']
      filename = secure_filename(f.filename)
      f.save(filename)
      excel_to_db(filename)

      return f'el archivo {filename} se guardó y actualizó la base de datos'

@app.route('/download')
def file_downloads():
    return render_template('downloads.html')

@app.route('/return-files/', methods=['GET', 'POST'])
def download():
    try:
        db_to_excel()
        filename = 'nueva_planilla.xlsx'
        return send_file(filename, cache_timeout=0)
    except Exception as e:
        return str(e)

@app.route('/up_json')
def up_json():
    data = to_json()
    return data


if __name__ == '__main__':
   app.run(debug = True)
