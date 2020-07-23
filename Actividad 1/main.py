# -*- coding: latin-1 -*-

from flask import Flask, flash, redirect, url_for, render_template, request, send_file, Response
from werkzeug.utils import secure_filename
import sqlite3
import collections
import pandas as pd
import xlsxwriter
import os
import json
from datetime import datetime, time

# Base Aplicación web
app = Flask(__name__)

UPLOAD_FOLDER = 'static/uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
#app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

#funcionalidades
def excel_to_db(new_excel_name, db_table='maquinaria_tipo'):
# new_excel_name type: str
# db_table type: str

#Connection to db
    conn = sqlite3.connect('elecnor.db')
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
            if number in dbdb['id']:
                cur.execute("UPDATE maquinaria_tipo SET nombre = ?, tiene_patente = ?, comentarios = ?, activo = ? WHERE id = ?", (nom, tp, com, act, number))
            else:
                cur.execute("INSERT INTO maquinaria_tipo VALUES(?,?,?,?,?)",(number, nom, tp, com, act))
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

    data = cur.execute('''SELECT * FROM maquinaria_tipo''').fetchall()

    conn.commit()
    conn.close()

    if json_str:
        return json.dumps([dict(ix) for ix in data], indent=10) #CREATE JSON

    return data

def json_file(data):
    with open('data.json', 'w') as outfile:
        json.dump(data, outfile)

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

      return f'el archivo {nombre_excel} se guardó y actualizó la base de datos'

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
    data = get_mq_tipo( json_str = True )
    return data

"""    data = get_mq_tipo( json_str = True )
    json_file(data)
    content = str(request.form)
    return Response(content,
            mimetype='application/json',
            headers={'Content-Disposition':'attachment;filename=zones.geojson'})"""

if __name__ == '__main__':
   app.run(debug = True)
