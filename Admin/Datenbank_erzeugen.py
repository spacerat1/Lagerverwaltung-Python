import sqlite3
import os
import pandas as pd
from tkinter import filedialog
from standard_table_layout import tables


name_of_database = 'Lagerverwaltung_Datenbank.db'
standardmaterial_file = 'Standardmaterial.xlsx'
kleinstmaterial_file = 'Kleinstmaterial.xlsx'
wareneingang_file = 'Wareneingang.xlsx'
actual_path = os.getcwd()

def create_standard_table():
    file_path_db = filedialog.askdirectory()
    if not file_path_db:
        return "\nKeinen Ordner ausgesucht. Erstellung abgebrochen.\n"
    if not os.path.exists(f"{file_path_db}/{name_of_database}"):
        with open(name_of_database, 'w'):
            pass
    
    database = f"{file_path_db}/{name_of_database}"
    connection = sqlite3.connect(database)
    cursor = connection.cursor()

    for table in tables.split(';'):
        cursor.execute(table)
    
    connection.commit()
    return
    # load standard material and write it into the database
    standardmaterial = get_data_from_file(standardmaterial_file)
    kleinstmaterial = get_data_from_file(kleinstmaterial_file)
    wareneingang = get_data_from_file(wareneingang_file)


    for _, material in standardmaterial.iterrows():
        cursor.execute('''
                        INSERT INTO 
                            Standardmaterial(MatNr, Bezeichnung,Einheit,Grenzwert,Auffüllen)
                        VALUES
                            (?,?,?,?,?)
                        ''', (material['MNr'], material['Bezeichnung'], material['Einheit'],
                              material['Grenzwert'], material['auffüllen auf'])
                        )
        
    for _, material in kleinstmaterial.iterrows():
        cursor.execute('''
                        INSERT INTO 
                            Kleinstmaterial(MatNr,Bezeichnung,Einheit,Grenzwert,Auffüllen)
                        VALUES
                            (?,?,?,?,?)
                        ''', (material['MNr'], material['Bezeichnung'], material['Einheit'],
                              material['Grenzwert'], material['auffüllen auf'])
                        )
            
    for _, daten in wareneingang.iterrows():
        cursor.execute('''
                       INSERT INTO
                            Wareneingang (MatNr, Bezeichnung, Menge)
                       VALUES 
                            (?,?,?)
                       ''',(daten['MNr'], daten['Bezeichnung'], daten['Menge'])
                       )
    
    
    connection.commit()
    cursor.close()
    connection.close()
    return "\nDatenbank erfolgreich erstellt.\n"

def get_data_from_file(file_name):
    path = f"{actual_path}/{file_name}"
    return pd.read_excel(path)
    


if __name__ == '__main__':
    print(create_standard_table())
   