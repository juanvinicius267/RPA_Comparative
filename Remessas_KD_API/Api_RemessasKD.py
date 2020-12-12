import pandas as pd
from flask import Flask, render_template, request, redirect, session, flash, url_for, jsonify
import json
import sqlite3
import numpy as np

app = Flask(__name__)

@app.route('/')
def index():
    listArray = []
    sqliteConnection = sqlite3.connect('sqlite/rpa.db')
    cursor = sqliteConnection.cursor()       
    cursor.execute("SELECT * FROM tb_remessas_kd WHERE comparative_done == '0'")
    sqliteConnection.commit()
    records = cursor.fetchall()
    for record in records:
        data = str(record).split(", ")
        infoJson= { 'id': data[0].replace("'", "").replace("(", ""),
                    'remessa': data[1].replace("'", ""),
                    'popId': data[2].replace("'", ""),
                    'chassi': data[3].replace("'", ""),
                    'cod_Cliente': data[4].replace("'", ""),
                    'part_Period': data[5].replace("'", ""),
                    'type': data[6].replace("'", ""),
                    'mercado': data[7].replace("'", ""),
                    'pru': data[8].replace("'", ""),
                    'modelo': data[9].replace("'", ""),
                    'pedidos': data[10].replace("'", ""),
                    'altura_chassi': data[11].replace("'", ""),
                    'suspensao': data[12].replace("'", ""),
                    'axle_distance': data[13].replace("'", ""),
                    'cab_type': data[14].replace("'", ""),
                    'cab_lenght': data[15].replace("'", ""),
                    'roof_height': data[16].replace("'", ""),
                    'comparative_done':data[17].replace("'", "").replace(")", "")
                    }
        listArray.append(infoJson)
        print(infoJson)
        
    cursor.close()    
    return str(listArray)
   

    


@app.route('/uploadFile')
def inicio():
    return render_template('uploadFile.html', titulo='Jogos')

@app.route('/fileUpload', methods=['POST',])
def file():
    #Recebe os arquivos da requisição HTTP
    file = request.files['file']  
    #Abre o excel, Escolhe a sheet, Escolhe as colunas e me retorna tudo em strings 
    data = pd.read_excel(file, sheet_name='MIX KD',header=None, names=['REMESSA','POPID','CHASSI','CÓDIGO DO CLIENTE','OM','TIPO','MERCADO',
    'PRU','MODELO','PEDIDOS','Altura Chassis','Suspensão Traseira','Axle distance (1406)','Cab type (42)','Cab length (116)','Roof height (584)'], usecols=[0,1,2,3,5,6,7,8,11,41,42,43,44,45,46,47])
    #Converte os dados em JSON      
    data = data.to_json(orient="records", date_format='iso')  
    setRemessas(data)
    return 'Ok'



@app.route('/update', methods=['POST',])
def update():
    content = request.json
    print(content['id'] + content['comparative_done'])    
    sqliteConnection = sqlite3.connect('sqlite/rpa.db')
    cursor = sqliteConnection.cursor()
    dados = (content['comparative_done'],content['id'])
    cursor.execute("UPDATE tb_remessas_kd SET comparative_done =? WHERE tb_remessas_kd.id =?",dados)
    sqliteConnection.commit()
    records = cursor.fetchall()
    return 'Ok'




def setRemessas(dados):    
    info = json.loads(dados)
    x = 100
    for i in info:         
        varQuery = '''insert into tb_remessas_kd (remessa , popid , chassi , codigo_do_cliente , om , tipo , mercado , pru, modelo , pedidos , altura_chassi , suspensao_traseira , axle_distance , cab_type , cab_lenght , roof_height )values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?);'''
        val = (i['REMESSA'],i['POPID'],i['CHASSI'],i['CÓDIGO DO CLIENTE'],i['OM'],i['TIPO'],i['MERCADO'],i['PRU'],i['MODELO'],i['PEDIDOS'],i['Altura Chassis'],i['Suspensão Traseira'],i['Axle distance (1406)'],i['Cab type (42)'],i['Cab length (116)'],i['Roof height (584)'])
        dbConnection(varQuery,val)
        

def dbConnection(query,val):
    try:
        sqliteConnection = sqlite3.connect('sqlite/rpa.db')
        cursor = sqliteConnection.cursor()
        print("Database created and Successfully Connected to SQLite")
        cursor.execute(query,val)
        sqliteConnection.commit()
        record = cursor.fetchall()
        print("SQLite Database Version is: ", record)
        cursor.close()
    except sqlite3.Error as error:
        print("Error while connecting to sqlite", error)
    finally:
        if (sqliteConnection):
            sqliteConnection.close()
            print("The SQLite connection is closed")

app.debug = True
app.run(host='10.251.16.133',port = 5001)

