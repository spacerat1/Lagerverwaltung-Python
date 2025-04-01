import os
import ctypes
import sqlite3
import pandas as pd
import tkinter as tk
import tkinter.ttk as ttk
import application
import re
import datetime
import win32com.client
import openpyxl as xl
from tkinter import filedialog
from collections import defaultdict


def open_db() -> str:
    settings = f"{os.environ['USERPROFILE']}\\Documents\\Lagerverwaltung_settings.txt"
    if not os.path.exists(settings):
        with open(settings, 'w'):
            pass
    with open(settings, 'r') as file:
        path_to_db = file.read()
    if not os.path.exists(path_to_db):
        ctypes.windll.user32.MessageBoxW(0,
                                         "Die Datenbank 'Lagerverwaltung_Datanbank.db' konnte nicht gefunden werden.\nBitte im nächsten Fenster die Datenbank auswählen", 
                                         "Pfad zur Datenbank suchen...", 
                                         64)
        path_to_db = change_db_path()
    return path_to_db


def change_db_path(app:application.App = None) -> str:
    settings = f"{os.environ['USERPROFILE']}\\Documents\\Lagerverwaltung_settings.txt"
    path_to_db = filedialog.askopenfile(initialfile = 'Lagerverwaltung_Datenbank.db', 
                                        defaultextension = '.db', 
                                        filetypes = [('Datenbank', '*.db'), ('All files', '*.*') ])
    if not path_to_db:
        return
    with open(settings, 'w') as file:
        file.write(path_to_db.name)
        path_to_db = path_to_db.name
    if app:
        app.path_to_db = path_to_db
        app.label_top_path_to_db.configure(text = f'Pfad zur Datenbank: {app.path_to_db}')
        if 'Service-Center' in app.path_to_db:
            color = 'forest green'
        else:
            color = 'orange red'
        app.label_top_path_to_db.configure(foreground = color)
        app.connection.close()
        app.connection = sqlite3.connect(path_to_db)
        app.connection.row_factory = sqlite3.Row
        app.cursor = app.connection.cursor()
        app.window.call(app.disabled_button['command'])
    return path_to_db


def set_widget_status(enabled: list[ttk.Widget], disabled:list[ttk.Widget]) -> None:
    for entry in enabled:
        if entry:
            entry.config(state = 'enabled')
    for entry in disabled:
        if entry:
            if isinstance(entry, ttk.Entry):
                entry.delete(0, tk.END)
            entry.config(state = 'disabled')
    

def show_critical_material(app:application.App) -> None:
    app.disabled_button.config(state = 'enabled')
    app.disabled_button = app.button_krit_mat
    app.button_krit_mat.config(state = 'disabled')
    
    enabled = [app.bestellt_button, app.bestellt_label, ]
    disabled = [app.sm_entry, app.posnr_entry, app.matnr_entry, app.delete_button, app.print_button, app.combobox_loeschen ]
    set_widget_status(enabled, disabled)
    
    app.matnr_entry.unbind('<Return>')
    app.matnr_entry.unbind('<KeyRelease>')
    app.sm_entry.unbind('<Return>')
    app.sm_entry.unbind('<KeyRelease>')
    app.posnr_entry.unbind('<Return>')
    app.posnr_entry.unbind('<KeyRelease>')
    app.combobox_loeschen.unbind('<<ComboboxSelected>>')

    cursor = app.cursor
    output_box = app.output_box

    ingoing_dict = defaultdict(int)
    outgoing_dict = defaultdict(int)
    threshhold_dict = defaultdict(int)
    auffuellen_dict = defaultdict(int)
    namen_dict = defaultdict(str)
    bestellt_dict = defaultdict(int)
    date_dict = defaultdict(str)
    ingoing_query = "SELECT * FROM Wareneingang"
    outgoing_query = "SELECT * FROM Warenausgang"
    outgoing_kleinst_query = "SELECT * FROM Warenausgang_Kleinstmaterial_ohne_SM_BEZUG"
    threshhold_query = "SELECT * FROM Standardmaterial UNION SELECT * FROM Kleinstmaterial"
    standardmaterial_query = "SELECT * FROM Standardmaterial"
    kleinstmaterial_query = "SELECT * FROM Kleinstmaterial"
    
    cursor.execute(ingoing_query)
    ingoing = cursor.fetchall()
    cursor.execute(outgoing_query)
    outgoing = cursor.fetchall()
    cursor.execute(outgoing_kleinst_query)
    outgoing_kleinst = cursor.fetchall()
    cursor.execute(threshhold_query)
    threshhold = cursor.fetchall()
    cursor.execute(standardmaterial_query)
    standard = cursor.fetchall()
    cursor.execute(kleinstmaterial_query)
    kleinst = cursor.fetchall()
    standardmaterial = [row['MatNr'] for row in standard]
    kleinstmaterial = [row['MatNr'] for row in kleinst]

    for row in ingoing:
        ingoing_dict[row['MatNr']] += row['Menge']

    for row in outgoing:
        if row['Warenausgangsmenge'] != 0:
            menge = row['Warenausgangsmenge']
        else:
            menge = row['Bedarfsmenge']

        if row['PosTyp'] == 8:
            outgoing_dict[row['MatNr']] -= menge
        elif row['PosTyp'] == 9:
            outgoing_dict[row['MatNr']] += menge
    for row in outgoing_kleinst:
        outgoing_dict[row['MatNr']] += row['Menge']

    for row in threshhold:
        threshhold_dict[row['MatNr']] = row['Grenzwert']
        auffuellen_dict[row['MatNr']] = row['Auffüllen']
        namen_dict[row['MatNr']] = row['Bezeichnung']
        bestellt_dict[row['MatNr']] = 'nein' if row['bestellt'] == 0 else 'ja'
        date_dict[row['MatNr']] = row['Datum'] if row['bestellt'] else ''

    kleinstmaterial_output= []
    standardmaterial_output = []
    for matnr, booked in sorted(outgoing_dict.items()):
        if ingoing_dict[matnr] - booked <= threshhold_dict[matnr]:
            text_line = f"  {str(matnr).ljust(10)}{namen_dict[matnr].ljust(50)}{str(ingoing_dict[matnr]-booked).center(7)}{str(auffuellen_dict[matnr]).center(18)}{bestellt_dict[matnr].center(10)}{str(date_dict[matnr]).center(19)}"
            if matnr in standardmaterial:
                standardmaterial_output.append(text_line)
            elif matnr in kleinstmaterial:
                kleinstmaterial_output.append(text_line)
    
    headline = (f"  {'MatNr.'.ljust(10)}{'Bezeichnung'.ljust(50)}{'Bestand'.ljust(9)}{'empfohlene Menge'.ljust(18)}{'bestellt?'.ljust(11)}{' '.ljust(2)}Datum\n")
    output_string_standard = '\n'.join(standardmaterial_output)
    output_string_kleinst = '\n'.join(kleinstmaterial_output)
    output_string = ("\n Das unten aufgeführte Material wird langsam knapp. Bitte nachbestellen.\n Um bestelltes Material als 'bestellt' zu markieren, bitte die entsprechenden Materialnummern mit der Maus markieren und dann den Knopf drücken.\n\n  Kritische Lagerbestände\n\n  Standardmaterial\n")
    output_string += headline
    output_string += ("-"*130)
    output_string += '\n'
    output_string += output_string_standard
    output_string += ("\n\n  Kleinstmaterial\n")
    output_string += headline
    output_string += ('-'*130)
    output_string += '\n'
    output_string += output_string_kleinst
    output_box.delete(1.0, 'end')
    output_box.insert(1.0, output_string)
    

def toggle_ordered_status(app:application.App) -> None:
    connection:sqlite3.Connection = app.connection
    cursor:sqlite3.Cursor = app.cursor
    try:
        selected_text = app.output_box.get(tk.SEL_FIRST, tk.SEL_LAST)
    except tk.TclError:
        return
    mat_numbers = re.findall(r'\d{8}', selected_text)
    if not mat_numbers:
        return
    
    for mat_number in mat_numbers:
        
        cursor.execute(f'''SELECT bestellt FROM Standardmaterial WHERE MatNr = {mat_number} 
                    UNION SELECT bestellt FROM Kleinstmaterial WHERE MatNr = {mat_number}''')
        status = cursor.fetchone()[0]
        status = not status

        if not status:
            cursor.execute(f''' UPDATE Standardmaterial
                                SET bestellt = {status}
                                WHERE MatNr = {mat_number}
                        ''')
            cursor.execute(f''' UPDATE Kleinstmaterial
                                SET bestellt = {status}
                                WHERE MatNr = {mat_number}
                        ''')
        else:
            datum = datetime.datetime.strftime(datetime.datetime.now(), r'%d.%m.%Y')
            #datum = datetime.datetime.now()
            cursor.execute(''' UPDATE Standardmaterial
                                SET bestellt = ?,
                                    Datum = ?
                                WHERE MatNr = ? 
                           ''', (status, datum, mat_number))
            cursor.execute(''' UPDATE Kleinstmaterial
                                SET bestellt = ?,
                                    Datum = ?
                                WHERE MatNr = ?
                            ''',(status, datum, mat_number))

            
    connection.commit()
    show_critical_material(app)    


def show_stock(app: application.App) -> None:
    app.disabled_button.config(state = 'enabled')
    app.disabled_button = app.button_bestand
    app.button_bestand.config(state = 'disabled')
    
    enabled = [app.matnr_entry]
    disabled = [app.posnr_entry, app.sm_entry, app.bestellt_button, app.bestellt_label, app.delete_button, app.print_button, app.combobox_loeschen]
    set_widget_status(enabled, disabled)
    
    app.sm_entry.unbind('<Return>')
    app.sm_entry.unbind('<KeyRelease>')
    app.posnr_entry.unbind('<Return>')
    app.posnr_entry.unbind('<KeyRelease>')
    app.combobox_loeschen.unbind('<<ComboboxSelected>>')
    app.matnr_entry.bind('<Return>', lambda _ : show_stock(app))
    app.matnr_entry.bind('<KeyRelease>', lambda _: show_stock(app))
    
    app.matnr_entry.focus()

    cursor = app.cursor
    matnr_entry = app.matnr_entry
    output_box = app.output_box
    matnr = matnr_entry.get()
    
    standardmaterial_query = "SELECT * FROM Standardmaterial"
    cursor.execute(standardmaterial_query)
    standardmaterial = cursor.fetchall()
    standardmaterial_list = [row['MatNr'] for row in standardmaterial]
    
    ingoing_query = "SELECT * FROM Wareneingang WHERE MatNr LIKE ?"
    cursor.execute(ingoing_query, (f'%{matnr}%',))
    ingoing = cursor.fetchall()
    ingoing_dict = defaultdict(int)
    for row in ingoing:
        ingoing_dict[row['MatNr']] += row['Menge']
    
    outgoing_query = "SELECT * FROM Warenausgang WHERE MatNr LIKE ?"
    cursor.execute(outgoing_query, (f'%{matnr}%',))
    outgoing = cursor.fetchall()
    
    outgoing_dict = defaultdict(int)
    for row in outgoing:
        if row['Warenausgangsmenge'] != 0:
            menge = row['Warenausgangsmenge']
        else:
            menge = row['Bedarfsmenge']
        if row['PosTyp'] == 8:
            outgoing_dict[row['MatNr']] -= menge
        elif row['PosTyp'] == 9:
            outgoing_dict[row['MatNr']] += menge

    outgoing_kleinst_query = "SELECT * FROM Warenausgang_Kleinstmaterial_ohne_SM_BEZUG"
    cursor.execute(outgoing_kleinst_query)
    outgoing_kleinst = cursor.fetchall()
    for row in outgoing_kleinst:
        outgoing_dict[row['MatNr']] += row['Menge']
    
    material_query = "SELECT * FROM Standardmaterial UNION SELECT * FROM Kleinstmaterial"
    cursor.execute(material_query)
    material = cursor.fetchall()
    material_dict = defaultdict(list)
    for row in material:
        if row['MatNr'] in standardmaterial_list:
            Zuordnung = 'Standardmaterial'
        else:
            Zuordnung = 'Kleinstmaterial'
        material_dict[row['MatNr']] = (row['Bezeichnung'], row['Einheit'], Zuordnung)
    
    result = []
    for matnr, menge in ingoing_dict.items():
        bestand = menge - outgoing_dict.get(matnr, 0)
        bezeichnung = material_dict[matnr][0]
        einheit = material_dict[matnr][1]
        zuordnung = material_dict[matnr][2]
        result.append((matnr, bezeichnung, bestand, einheit, zuordnung))
    
    output_string = f"  {'MatNr.'.ljust(8)}\t\t{'Bezeichnung'.ljust(50)}\tBestand{' '*5}Einheit{' '*10}{'Zuordnung'.center(16)}\n"
    output_material = ''
    for matnr, bezeichnung, bestand, einheit, zuordnung in sorted(result):
        output_material += f"  {matnr}\t\t{bezeichnung.ljust(50)}{str(bestand).center(7)}{' '*5}{einheit.center(7)}{' ' *10}{zuordnung}\n"
    if not output_material:
        output_string = '\nKeine Datensätze gefunden. Bitte Filter überprüfen.'
    else:
        output_string += output_material
    output_box.delete(1.0, 'end')
    output_box.insert(1.0, output_string)
    
     
def show_ingoing_material(app:application.App) -> None:
    app.disabled_button.config(state = 'enabled')
    app.disabled_button = app.button_wareneingang
    app.button_wareneingang.config(state = 'disabled')
    
    enabled = [app.matnr_entry]
    disabled = [app.sm_entry, app.posnr_entry, app.bestellt_label, app.bestellt_button, app.delete_button, app.print_button, app.combobox_loeschen]
    set_widget_status(enabled, disabled)
    app.matnr_entry.focus()
    app.sm_entry.unbind('<Return>')
    app.sm_entry.unbind('<KeyRelease>')
    app.posnr_entry.unbind('<Return>')
    app.posnr_entry.unbind('<KeyRelease>')
    app.matnr_entry.bind('<Return>', lambda _ : show_ingoing_material(app))
    app.matnr_entry.bind('<KeyRelease>', lambda _: show_ingoing_material(app))
    app.combobox_loeschen.unbind('<<ComboboxSelected>>')

    cursor:sqlite3.Cursor = app.cursor
    matnr = app.matnr_entry.get()
    cursor.execute("SELECT * FROM Wareneingang WHERE MatNr LIKE ?", (f'%{matnr}%',))
    selection = cursor.fetchall()
    einheits_query = "SELECT * FROM Kleinstmaterial UNION SELECT * FROM Standardmaterial"
    cursor.execute(einheits_query)
    einheiten = cursor.fetchall()
    einheitsdict = {row['MatNr']:row['Einheit'] for row in einheiten}
    headline = (f"{'ID'.rjust(5)}\t{'MatNr.'.ljust(8)} \t{'Bezeichnung'.ljust(50)}  {'Menge'.ljust(8)}\
\t\tEinheit\t{'Datum'.ljust(10)}\n")
    output = headline if selection else '\nKeine Datensätze gefunden. Bitte Filter überprüfen.'
    for row in selection[::-1]:
        text = (f"{str(row['ID']).rjust(5)}\t{str(row['MatNr']).ljust(8)} \t{row['Bezeichnung'].ljust(50)}  {str(row['Menge']).center(8)}\
\t\t{einheitsdict[row['MatNr']].center(7)}\t{row['Datum'].ljust(10)}\n")
        output += text
    
    app.output_box.delete(1.0, 'end')
    app.output_box.insert(1.0, output)


def show_outgoing_material(app:application.App) -> None:
    app.disabled_button.config(state = 'enabled')
    app.disabled_button = app.button_warenausgang
    app.button_warenausgang.config(state = 'disabled')
    
    enabled = [app.sm_entry, app.matnr_entry]
    disabled = [app.posnr_entry, app.bestellt_label, app.bestellt_button, app.delete_button, app.print_button, app.combobox_loeschen]
    set_widget_status(enabled, disabled)
    if str(app.matnr_entry.focus_get()) != '.!frame3.!entry3': # .!frame3.!entry3 ist der Name vom SM_ENTRY Widget
        app.matnr_entry.focus()
    app.sm_entry.bind('<Return>', lambda _ : show_outgoing_material(app))
    app.sm_entry.bind('<KeyRelease>', lambda _ : show_outgoing_material(app))
    app.posnr_entry.unbind('<Return>')
    app.posnr_entry.unbind('<KeyRelease>')
    app.matnr_entry.bind('<Return>', lambda _ : show_outgoing_material(app))
    app.matnr_entry.bind('<KeyRelease>', lambda _ : show_outgoing_material(app))
    app.combobox_loeschen.unbind('<<ComboboxSelected>>')

    cursor:sqlite3.Cursor = app.cursor
    matnr = app.matnr_entry.get()
    smnr = app.sm_entry.get()
    cursor.execute("SELECT * FROM Warenausgang WHERE MatNr LIKE ? AND SM_Nummer LIKE ? AND PosTyp = 9", (f'%{matnr}%', f'%{smnr}%',))
    selection = cursor.fetchall()
    output = ''
    headline = (f"{'Nr.'.rjust(5)}\t{'SM Nummer'.ljust(10)}\t{'Pos'.ljust(3)}\t{'Typ'.ljust(3)}  {'MatNr'.ljust(8)}\
\t\t{'Bezeichnung'.ljust(50)}\t{'SD Beleg'.ljust(10)}\t{'Bedarfsmenge'.ljust(12)}\
\t\t{'Warenausgangsmenge'.ljust(18)}\t{'Umbuchungsmenge'.ljust(15)}\t{'Lieferschein'.ljust(12)}\t{'Materialbeleg'.ljust(13)}\n")
    #output = headline if selection else 'Keine Daten gefunden. Bitte Filter überprüfen.'
    start = 1
    if selection:
        output += headline
        for number, row in enumerate(selection, start = start):
            #print(f"{row['SM_Nummer']}\t{row['Warenausgangsmenge']}")
            text = (f"{str(number).rjust(5)}\t{row['SM_Nummer'].ljust(10)}\t{row['Position'].ljust(3)}\t{str(row['PosTyp']).ljust(3)}  {str(row['MatNr']).ljust(8)}\
    \t\t{row['Bezeichnung'].ljust(50)}\t{row['SD_Beleg'][:10].ljust(10)}\t{str(row['Bedarfsmenge']).center(12)}\
    \t\t{str(row['Warenausgangsmenge']).center(18)}\t{str(row['Umbuchungsmenge']).center(15)}\t{row['Lieferschein'][:10].ljust(12)}\t{row['Materialbeleg'][:10].ljust(13)}\n")
            output += text
        start = number+1

    selection_kleinst = ''
    if not app.sm_entry.get():
        unit_query = "SELECT * FROM Kleinstmaterial"
        cursor.execute(unit_query)
        units = cursor.fetchall()
        units_dict = {unit['MatNr']: unit['Einheit'] for unit in units}
        cursor.execute("SELECT * FROM Warenausgang_Kleinstmaterial_ohne_SM_Bezug WHERE MatNr LIKE ?", (f'%{matnr}%',))
        selection_kleinst = cursor.fetchall()
        if selection_kleinst:
            output += '\nAusgabe Kleinstmaterial ohne SM Bezug\n'
            output += (f"{'Nr.'.rjust(5)}\t{'MatNr'.ljust(8)}\t{'Bezeichnung'.ljust(50)}\t{'Menge'.ljust(5)}\t{'Einheit'.ljust(7)}\n")
            for number,row in enumerate(selection_kleinst, start = start):
                output += (f"{str(number).rjust(5)}\t{str(row['Matnr']).ljust(8)}\t{row['Bezeichnung'].ljust(50)}\t{str(row['Menge']).center(5)}\t{units_dict[row['MatNr']].center(6)}\n")
    if not selection and not selection_kleinst:
        output += 'Keine Daten gefunden. Bitte Filter überprüfen.'
    app.output_box.delete(1.0, 'end')
    app.output_box.insert(1.0, output)
    # app.output_box.tag_configure('headline - bold', )


def show_material_for_order(app:application.App) -> None:
    app.disabled_button.config(state = 'enabled')
    app.disabled_button = app.button_sm_auftrag
    app.button_sm_auftrag.config(state = 'disabled')
    
    enabled = [app.sm_entry, app.print_button]
    disabled = [app.posnr_entry, app.matnr_entry, app.bestellt_label, app.bestellt_button, app.delete_button, app.combobox_loeschen]
    set_widget_status(enabled, disabled)
    
    app.sm_entry.focus()
    app.sm_entry.bind('<Return>', lambda _ : show_material_for_order(app))
    app.sm_entry.bind('<KeyRelease>', lambda _ : show_material_for_order(app))
    app.posnr_entry.unbind('<Return>')
    app.posnr_entry.unbind('<KeyRelease>')
    app.matnr_entry.unbind('<Return>')
    app.matnr_entry.unbind('<KeyRelease>')
    app.combobox_loeschen.unbind('<<ComboboxSelected>>')

    cursor:sqlite3.Cursor = app.cursor
    smnr = app.sm_entry.get()
    cursor.execute("SELECT * FROM Warenausgabe_Comline WHERE SM_Nummer LIKE ?", (f'%{smnr}%',))
    selection = cursor.fetchall()
    cursor.execute("SELECT * FROM Standardmaterial")
    standard = cursor.fetchall()
    cursor.execute("SELECT * FROM Kleinstmaterial")
    kleinst = cursor.fetchall()
    standard_dict = {row['MatNr']:row['Einheit'] for row in standard}
    kleinst_dict= {row['MatNr']:row['Einheit'] for row in kleinst}

    standard_material = []
    kleinst_material = []
    telekom_material = []
    for row in selection:
        if row['MatNr'] in standard_dict:
            standard_material.append(row)
        elif row['MatNr'] in kleinst_dict:
            kleinst_material.append(row)
        else:
            telekom_material.append(row)

    if not selection:
        app.output_box.delete(1.0, 'end')
        app.output_box.insert(1.0, 'Keine Daten gefunden. Bitte Filter überprüfen.')
        return

    headline = (f"{'SM Nummer'.ljust(10)}\t{'MatNr'.ljust(8)}\t{'Bezeichnung'.ljust(45)}\t{'Menge'.ljust(5)}\t{'Einheit'}\n")
    
    if standard_material:
        output_text = '\n Standardmaterial\n'
        output_text += headline
    else:
        output_text = ''
    for row in standard_material:
        einheit = standard_dict.get(row['MatNr'], None) or kleinst_dict.get(row['MatNr'], 'n/a')
        text = (f"{row['SM_Nummer'].ljust(10)}\t{str(row['MatNr']).ljust(8)}\t{row['Bezeichnung'].ljust(45)}\t{str(row['Bedarfsmenge']).center(5)}\t{einheit.center(6)}\n")
        output_text += text
    if kleinst_material:
        output_text += '\n Kleinstmaterial\n'  
        output_text += headline
    
    for row in kleinst_material:
        einheit = standard_dict.get(row['MatNr'], None) or kleinst_dict.get(row['MatNr'], 'n/a')
        text = (f"{row['SM_Nummer'].ljust(10)}\t{str(row['MatNr']).ljust(8)}\t{row['Bezeichnung'].ljust(45)}\t{str(row['Bedarfsmenge']).center(5)}\t{einheit.center(6)}\n")
        output_text += text
    if telekom_material:
        output_text  += '\n Telekom Material (per SM Nummer bestellt)\n'
        output_text += headline
    for row in telekom_material:
        einheit = standard_dict.get(row['MatNr'], None) or kleinst_dict.get(row['MatNr'], 'n/a')
        text = (f"{row['SM_Nummer'].ljust(10)}\t{str(row['MatNr']).ljust(8)}\t{row['Bezeichnung'].ljust(45)}\t{str(row['Bedarfsmenge']).center(5)}\t{einheit.center(6)}\n")
        output_text += text
    
    app.output_box.delete(1.0, 'end')
    app.output_box.insert(1.0, output_text)


def print_screen(app:application.App) -> None:
    file_path = f"{os.getcwd()}\\Ausgabe.xlsx"
    if not app.sm_entry.get():
        return
    answer = ctypes.windll.user32.MessageBoxW(0,"Soll der angezeigte Inhalt gedruckt werden?", "Drucken...", 68)
    if answer == 6:
        selection = app.output_box.get(1.0, 'end')
        wb = xl.Workbook()
        sheet = wb.active
        sheet.page_setup.orientation = 'landscape'
        sheet.page_setup.fitToPage = True
        sheet.column_dimensions['A'].width = 12
        sheet.column_dimensions['B'].width = 11
        sheet.column_dimensions['C'].width = 44
        sheet.column_dimensions['D'].width = 7
        sheet.column_dimensions['E'].width = 7
        sheet.column_dimensions['F'].width = 3
        for x in range(7,37):
            if x % 2 == 1:
                sheet.column_dimensions[xl.utils.cell.get_column_letter(x)].width = 1
            else:
                sheet.column_dimensions[xl.utils.cell.get_column_letter(x)].width = 3
        rows = selection.split('\n')
        for row, line in enumerate(rows, start = 1):
            for col, word in enumerate(line.split('\t'), start = 1):
                sheet.cell(column =col, row = row, value = word)
        border = xl.styles.Side(border_style = 'thin', color = '000000')
        for idx, row in enumerate(sheet, start = 1):
            try:
                _ = int(row[0].value)
                # draw cell bottom lines
                for col in range(1,6):
                    sheet.cell(row = idx, column = col).border = xl.styles.Border(bottom = border)
                # change font color to red when amount > 1 and unit is M or n/a
                if int(sheet.cell(row = idx, column = 4).value) > 1 and sheet.cell(row= idx, column = 5).value.strip() in ('M', 'n/a'):
                    for col in range(1,6):
                        sheet.cell(row = idx, column = col).font = xl.styles.Font(color="FF0000", bold = True)
                if sheet.cell(row= idx, column = 5).value.strip() in ('ST', 'SA'):
                    amount = int(sheet.cell(row= idx, column=4).value) * 2
                else:
                    amount = 2
                for col in range(amount):
                    if col % 2 == 0:
                        continue
                    sheet.cell(row = idx, column = col+7).border = xl.styles.Border(bottom = border, top = border, left = border, right = border)
                
            except ValueError:
                continue
        
        wb.save(file_path)
        try:
            excel_app = win32com.client.Dispatch('Excel.Application')
            excel_app.Visible = False
            wb = excel_app.Workbooks.Open(file_path)
            wb.PrintOut()
            #ws = wb.Worksheets[1]
            #ws.PrintOut()
        finally:
            wb.Close(False)
            excel_app.Quit()
            os.remove(file_path)

        
def book_outgoing_from_excel_file(app:application.App) -> None:
    # ctypes.windll.user32.MessageBoxW(0,"Bitte im nächsten Fenster die exportierte Materialliste aus PSL auswählen.", "Warenausgang buchen...", 64)
    cursor:sqlite3.Cursor = app.cursor
    connection:sqlite3.Connection = app.connection
    path_to_excel = filedialog.askopenfilenames(defaultextension = 'xlsx', title = 'Bitte die EXPORT Datei aus PSL wählen - Mehrfachauswahl ist möglich...')
    if not path_to_excel:
        return
    all_rows = new_entries = no_booking = already_exists = 0
    
    for file in path_to_excel:
        excel_file = pd.read_excel(file)
        valid_entries = []
        absolute_valid_entries = []
        ausgabe_entries = []
        valid_ausgabe_entries = []

        # sammle Kleinstmaterial und Standardmaterial
        cursor.execute('SELECT MatNr FROM Kleinstmaterial')
        selection_kleinstmaterial = cursor.fetchall()
        kleinstmaterial = [row[0] for row in selection_kleinstmaterial]
        cursor.execute('SELECT MatNr FROM Standardmaterial')
        selection_standardmaterial = cursor.fetchall()
        standardmaterial = [row[0] for row in selection_standardmaterial]

        for _, line in excel_file.iterrows():
            if line['Positionstyp'] in (9,'9','N'):
                ausgabe_entries.append(line)
            if line['Material'] in kleinstmaterial:
                valid_entries.append(line)
                continue
            
            if line['Material'] in standardmaterial and (not pd.isnull(line['Lieferungen(ab/bis)']) or line['Positionstyp'] in (8,'8')):
                valid_entries.append(line)
        
        # check if position already exists in db
        for entry in valid_entries:
            cursor.execute(f"SELECT * FROM Warenausgang WHERE SM_Nummer = {entry['Auftrag']} AND Position = {entry['Stl.Position']}")
            result = cursor.fetchall()
            if not result:
                absolute_valid_entries.append(entry)

        for entry in ausgabe_entries:
            cursor.execute(f"SELECT * FROM Warenausgabe_Comline WHERE SM_Nummer = {entry['Auftrag']} AND Position = {entry['Stl.Position']}")
            result = cursor.fetchall()
            if not result:
                valid_ausgabe_entries.append(entry)
        
        for line in absolute_valid_entries:
            entry = line.fillna(" ")
            cursor.execute('''
                            INSERT INTO 
                                Warenausgang (SM_Nummer, Position, PosTyp, MatNr, Bezeichnung, SD_Beleg, Bedarfsmenge, Warenausgangsmenge, Umbuchungsmenge, Lieferschein, Materialbeleg)
                            VALUES
                                (?,?,?,?,?,?,?,?,?,?,?)
                        ''', (entry['Auftrag'],entry['Stl.Position'],entry['Positionstyp'],entry['Material'],
                entry['Materialkurztext'],entry['SD Beleg'],entry['Bedarfsmenge'],entry['Warenausgangsmenge'],
                entry['Umbuchungsmenge'],entry['Lieferschein'],entry['Materialbeleg'])
                            )

        for line in valid_ausgabe_entries:
            entry = line.fillna(" ")
            cursor.execute('''
                            INSERT INTO 
                                Warenausgabe_Comline (SM_Nummer, Position, PosTyp, MatNr, Bezeichnung, SD_Beleg, Bedarfsmenge, Warenausgangsmenge, Umbuchungsmenge, Lieferschein, Materialbeleg)
                            VALUES
                                (?,?,?,?,?,?,?,?,?,?,?)
                        ''', (entry['Auftrag'],entry['Stl.Position'],entry['Positionstyp'],entry['Material'],
                entry['Materialkurztext'],entry['SD Beleg'],entry['Bedarfsmenge'],entry['Warenausgangsmenge'],
                entry['Umbuchungsmenge'],entry['Lieferschein'],entry['Materialbeleg'])
                            )    

        connection.commit()         
        
        rows_in_excel = excel_file.shape[0]
        valid_count = len(valid_entries)
        absolute_valid_count = len(absolute_valid_entries)
        all_rows += rows_in_excel
        new_entries += absolute_valid_count
        no_booking += (rows_in_excel - valid_count)
        already_exists += (valid_count - absolute_valid_count)
        
    ctypes.windll.user32.MessageBoxW(0,f'''
                {new_entries} von {all_rows} Datensätze in Datenbank aufgenommen.
                {no_booking} Einträge enthielten keine Bestellung.
                {already_exists} Einträge waren bereits in der Datenbank enthalten.       
            ''', 'Ergebnis', 64)
    show_critical_material(app)


def book_ingoing_position(app:application.App) -> None:
    connection = app.connection
    cursor = app.cursor
    
    app.disabled_button.config(state = 'enabled')
    app.disabled_button = app.button_wareneingang_buchen
    app.button_wareneingang_buchen.config(state = 'disabled')
    selection = 'SELECT * FROM Kleinstmaterial UNION SELECT * FROM Standardmaterial'
    app.open_booking_window(selection)
    try:
        matnr = re.findall(r'\d{8}', app.ingoing_mat_string_var.get())[0]
    except IndexError:
        show_ingoing_material(app)
        return
    menge = app.ingoing_menge_var.get()
    if menge == 0:
        show_ingoing_material(app)
        return
    material_query =  "SELECT * FROM Standardmaterial WHERE MatNr = ? UNION SELECT * FROM Kleinstmaterial WHERE MatNr = ?"
    cursor.execute(material_query, (matnr, matnr))
    material = cursor.fetchall()[0]
    bezeichnung = material['Bezeichnung']
    einheit = material['Einheit']
    answer = ctypes.windll.user32.MessageBoxW(0,f"Soll folgendes Material gebucht werden?\n\n{matnr}    '{bezeichnung}'       {menge} {einheit}", "Warenausgang buchen...", 68)
    if answer != 6:
        #app.button_wareneingang_buchen.config(state = 'enabled')
        show_ingoing_material(app)
        return
    cursor.execute('''
                        INSERT INTO
                                Wareneingang (MatNr, Bezeichnung, Menge)
                        VALUES 
                                (?,?,?)
                        ''',(matnr, bezeichnung, menge)
                        )
    cursor.execute(f''' UPDATE Standardmaterial
                            SET bestellt = 0
                            WHERE MatNr = {matnr}
                       ''')
    cursor.execute(f''' UPDATE Kleinstmaterial
                            SET bestellt = 0
                            WHERE MatNr = {matnr}
                       ''') 

    connection.commit()
    app.button_wareneingang_buchen.config(state = 'enabled')
    show_ingoing_material(app)

def book_outgoing_kleinstmaterial(app:application.App) -> None:
    connection = app.connection
    cursor = app.cursor
    app.disabled_button.config(state = 'enabled')
    app.disabled_button = app.button_warenausgang_buchen_Kleinstmaterial
    app.button_warenausgang_buchen_Kleinstmaterial.config(state = 'disabled')
    selection = 'SELECT * FROM Kleinstmaterial'
    app.open_booking_window(selection)
    try:
        matnr = re.findall(r'\d{8}', app.ingoing_mat_string_var.get())[0]
    except IndexError:
        show_critical_material(app)
        return
    menge = app.ingoing_menge_var.get()
    if menge == 0:
        show_critical_material(app)
        return
    material_query =  "SELECT * FROM Kleinstmaterial WHERE MatNr = ?"
    cursor.execute(material_query, (matnr,))
    material = cursor.fetchall()[0]
    bezeichnung = material['Bezeichnung']
    einheit = material['Einheit']
    answer = ctypes.windll.user32.MessageBoxW(0,f"Soll folgendes Material gebucht werden?\n\n{matnr}    '{bezeichnung}'       {menge} {einheit}", "Warenausgang buchen...", 68)
    if answer != 6:
        show_critical_material(app)
        return
    cursor.execute('''
                        INSERT INTO
                                Warenausgang_Kleinstmaterial_ohne_SM_Bezug (MatNr, Bezeichnung, Menge)
                        VALUES 
                                (?,?,?)
                        ''',(matnr, bezeichnung, menge)
                        )
    connection.commit()
    app.button_warenausgang_buchen_Kleinstmaterial.config(state = 'enabled')
    show_critical_material(app)


def filter_entries_to_delete(app:application.App) -> None:
    cursor:sqlite3.Cursor = app.cursor
    
    app.disabled_button.config(state = 'enabled')
    app.disabled_button = app.button_loeschen
    app.button_loeschen.config(state = 'disabled')
    
    combobox_current = app.combobox_loeschen.get()
    if not combobox_current:
        app.combobox_loeschen.current(0)

    enabled = [app.sm_entry, app.posnr_entry,app.matnr_entry, app.delete_button, app.combobox_loeschen]
    disabled = [app.bestellt_label, app.bestellt_button, app.print_button]
    enabled.extend(app.filter_dict[app.combobox_loeschen.get()][0])
    disabled.extend(app.filter_dict[app.combobox_loeschen.get()][1])
    set_widget_status(enabled, disabled)

    app.sm_entry.bind('<KeyRelease>', lambda _ : filter_entries_to_delete(app))
    app.posnr_entry.bind('<KeyRelease>', lambda _ : filter_entries_to_delete(app))
    app.matnr_entry.bind('<KeyRelease>', lambda _ : filter_entries_to_delete(app))
    app.combobox_loeschen.bind('<<ComboboxSelected>>', lambda _: filter_entries_to_delete(app))

    sm = app.sm_entry.get()
    posnr = app.posnr_entry.get()
    matnr = app.matnr_entry.get()

    execution_string = app.execution_dict[app.combobox_loeschen.get()][0]
    execute_values = app.execution_dict[app.combobox_loeschen.get()][1]
    execute_values = execute_values.replace('matnr', matnr).replace('sm', sm).replace('posnr', posnr)
    execution_tuple = tuple(execute_values.split(','))
    execution_tuple = tuple([entry for entry in execution_tuple if entry])
    if not execution_tuple:
        execution_tuple = ('',)    

    cursor.execute(execution_string, (execution_tuple))
    app.execution_string = execution_string
    app.execution_tuple = execution_tuple
    selection = cursor.fetchall()
    output = ''
    for row in selection:
        for field in row:
            output += f"{field}\t"
        output += '\n'
    app.output_box.delete(1.0, 'end')
    app.output_box.insert(1.0, output)


def delete_selected_entries(app:application.App) -> None:
    answer = ctypes.windll.user32.MessageBoxW(0,
                                              "Achtung!!\nAlle im Fenster angezeigten Einträge werden unwiderruflich gelöscht.\nBist du sicher, dass du die Einträge löschen willst?", 
                                              "Einträge löschen", 
                                              68)
    if answer != 6:
        return
    execution_string = app.execution_string.replace('SELECT *', 'DELETE')
    execution_tuple = app.execution_tuple
    app.connection.execute(execution_string, (execution_tuple))
    app.connection.commit()
    filter_entries_to_delete(app)

   
def add_kleinstmaterial(connection:sqlite3.Connection, cursor:sqlite3.Cursor, data:str) -> None:
    ''' Adds a new entry in table 'Kleinstmaterial'
        connection = sqlite3.Connection
        cursor = sqlite3.Cursor
        data = (matnr [int], bezeichnung [str], einheit [str], grenzwert [int], up_to [int], bestellt [bool])
    '''
    matnr, bezeichnung, einheit, grenzwert, up_to, bestellt = data
    cursor.execute('''
                        INSERT INTO 
                            Kleinstmaterial(MatNr,Bezeichnung,Einheit,Grenzwert,Auffüllen, bestellt)
                        VALUES
                            (?,?,?,?,?,?)
                        ''', (matnr, bezeichnung, einheit, grenzwert, up_to, bestellt)
                        )
    connection.commit()


def add_standardmaterial(connection:sqlite3.Connection, cursor:sqlite3.Cursor,  data:tuple):
    ''' Adds a new entry in table 'Standardmaterial'
        connection = sqlite3.Connection
        cursor = sqlite3.Cursor
        data = (matnr [int], bezeichnung [str], einheit [str], grenzwert [int], up_to [int], bestellt [bool])
    '''
    
    matnr, bezeichnung, einheit, grenzwert, up_to, bestellt = data
    cursor.execute('''
                        INSERT INTO 
                            Standardmaterial(MatNr,Bezeichnung,Einheit,Grenzwert,Auffüllen, bestellt)
                        VALUES
                            (?,?,?,?,?,?)
                        ''', (matnr, bezeichnung, einheit, grenzwert, up_to, bestellt)
                        )
    connection.commit()