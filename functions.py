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
    

def unbind_all_widgets(app:application.App) -> None:
    app.matnr_entry.unbind('<Return>')
    app.matnr_entry.unbind('<KeyRelease>')
    app.sm_entry.unbind('<Return>')
    app.sm_entry.unbind('<KeyRelease>')
    app.posnr_entry.unbind('<Return>')
    app.posnr_entry.unbind('<KeyRelease>')
    app.combobox_loeschen.unbind('<<ComboboxSelected>>')


def show_context_menu(event:tk.Event, app:application.App) -> None:
    #print(app.output_listbox['columns'])
    selections =  app.output_listbox.selection()
    if not selections:
        return
    app.context_menu.delete(0, 'end')
    columns = app.output_listbox['columns']
    if 'MatNr.' in columns:
        app.context_menu.add_command(label = 'Kopieren: Materialnummer', command  = lambda: copy_matnr(app))
    if 'SM Nummer / ID' in columns or 'SM Nummer' in columns:
        app.context_menu.add_command(label = 'Kopieren: SM Nummer', command = lambda: copy_sm_nr(app))
    if 'Bezeichnung' in columns:
        app.context_menu.add_command(label = 'Kopieren: Bezeichnung', command = lambda: copy_name(app))
    app.context_menu.add_command(label = 'Kopieren: Ganze Zeile', command = lambda: copy_line(app))
    app.context_menu.post(event.x_root, event.y_root)


def copy_matnr(app:application.App) -> None:
    mat_numbers = []
    selections =  app.output_listbox.selection()
    for selection in selections:
        matnr = app.output_listbox.set(selection, column = 'MatNr.')
        if not matnr:
            continue
        mat_numbers.append(matnr)
    output_string = '\n'.join(mat_numbers)
    app.window.clipboard_clear()
    app.window.clipboard_append(output_string)
    app.window.update()

def copy_sm_nr(app:application.App) -> None:
    sm_numbers = []
    if 'SM Nummer / ID' in app.output_listbox['columns']:
        sm_column = 'SM Nummer / ID'
    else:
        sm_column = 'SM Nummer'
    selections =  app.output_listbox.selection()
    
    for selection in selections:
        smnr = app.output_listbox.set(selection, column = sm_column)
        if not smnr:
            continue
        sm_numbers.append(smnr)
    output_string = '\n'.join(sm_numbers)
    app.window.clipboard_clear()
    app.window.clipboard_append(output_string)
    app.window.update()


def copy_name(app:application.App) -> None:
    names = []
    selections =  app.output_listbox.selection()
    for selection in selections:
        name = app.output_listbox.set(selection, column = 'Bezeichnung')
        if not name:
            continue
        names.append(name)
    output_string = '\n'.join(names)
    app.window.clipboard_clear()
    app.window.clipboard_append(output_string)
    app.window.update()

def copy_line(app:application.App) -> None:
    lines = []
    selections =  app.output_listbox.selection()
    for selection in selections:
        values = app.output_listbox.item(selection, 'values')
        if not values:
            continue
        lines.append(values)
    output_list = []
    for value in lines:
        val_str = str(value).strip('()').replace(", '", '\t').replace("'","")
        output_list.append(val_str)
    output_string = '\n'.join(output_list)
    app.window.clipboard_clear()
    app.window.clipboard_append(output_string)
    app.window.update()

def on_mouse_drag(event, app:application.App):
    row_id = app.output_listbox.identify_row(event.y)
    if row_id:
        app.output_listbox.selection_add(row_id)    
        

def show_critical_material(app:application.App) -> None:
    app.disabled_button.config(state = 'enabled')
    app.disabled_button = app.button_krit_mat
    app.button_krit_mat.config(state = 'disabled')
    
    enabled = [app.bestellt_button, app.bestellt_label, ]
    disabled = [app.sm_entry, app.posnr_entry, app.matnr_entry, app.delete_button, app.print_button, app.combobox_loeschen ]
    set_widget_status(enabled, disabled)
    unbind_all_widgets(app)

    cursor = app.cursor
    ingoing_dict = defaultdict(int)
    outgoing_dict = defaultdict(int)
    already_ordered_dict = defaultdict(int)
    date_dict = defaultdict(str)
    
    ingoing = cursor.execute("SELECT * FROM Wareneingang").fetchall()
    outgoing = cursor.execute("SELECT * FROM Warenausgang").fetchall()
    outgoing_small_material = cursor.execute("SELECT * FROM Warenausgang_Kleinstmaterial_ohne_SM_BEZUG").fetchall()
    all_materials = cursor.execute("SELECT * FROM Standardmaterial UNION SELECT * FROM Kleinstmaterial").fetchall()
    # calculate the stock
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
    for row in outgoing_small_material:
        outgoing_dict[row['MatNr']] += row['Menge']

    for row in all_materials:
        already_ordered_dict[row['MatNr']] = 'nein' if row['bestellt'] == 0 else 'ja'
        date_dict[row['MatNr']] = row['Datum'] if row['bestellt'] else ''

    # formatting the putput
    small_material_output= []
    standard_material_output = []
    for matnr, booked in sorted(outgoing_dict.items()):
        if ingoing_dict[matnr] - booked <= app.threshhold_dict[matnr]:
            output = [matnr, app.materialnames_dict[matnr], ingoing_dict[matnr]-booked, app.units_dict[matnr], app.recommended_amount_dict[matnr], already_ordered_dict[matnr], date_dict[matnr]]
            if matnr in app.standard_materials:
                standard_material_output.append(output)
            elif matnr in app.small_materials:
                small_material_output.append(output)
    
    # fill the treeview widget
    for item in app.output_listbox.get_children():
        app.output_listbox.delete(item)
    columns = ['MatNr.', 'Bezeichnung', 'Bestand', 'Einheit', 'empfohlene Menge', 'bestellt', 'Datum', 'LAST_COLUMN']
    app.output_listbox.configure(columns = columns)
    for column in columns:
        if column == 'LAST_COLUMN':
            continue
        width, anchor = app.columns_dict[column]
        app.output_listbox.heading(column, text = column)
        app.output_listbox.column(column, width = width, anchor = anchor, stretch = tk.NO)
    
    
    app.output_listbox.column('#0', width = 200, stretch = tk.NO) #erste Spalte fixieren (dort, wo das Parent erscheint)
    if standard_material_output:
        tree_standard = app.output_listbox.insert('', 'end', text = 'Standardmaterial', open = True, tags = ('green',))
        for entry in standard_material_output:
            app.output_listbox.insert(tree_standard, "end", values = entry)
    if small_material_output:
        tree_small_material = app.output_listbox.insert('', 'end', text = 'Kleinstmaterial', open = True, tags = ('green',))
        for entry in small_material_output:
            app.output_listbox.insert(tree_small_material, "end", values = entry)
    if  not standard_material_output and not small_material_output:
        app.output_listbox.insert('', 'end', values = ('', 'Sieht gut aus, wir haben alles. :)'))
    

def toggle_ordered_status(app:application.App) -> None:
    connection:sqlite3.Connection = app.connection
    cursor:sqlite3.Cursor = app.cursor

    selections =  app.output_listbox.selection()
    mat_numbers = []
    for selection in selections:
        values = app.output_listbox.item(selection, 'values')
        if not values:
            continue
        mat_numbers.append(values[0])

    for mat_number in mat_numbers:
        cursor.execute('''SELECT bestellt 
                       FROM Standardmaterial 
                       WHERE MatNr = ? 
                       UNION SELECT bestellt 
                       FROM Kleinstmaterial 
                       WHERE MatNr = ?
                       ''', (mat_number, mat_number)
                       )
        status = cursor.fetchone()[0]
        status = not status
        if not status:
            cursor.execute( ''' UPDATE Standardmaterial
                                SET bestellt = ?
                                WHERE MatNr = ?
                            ''',(status, mat_number)
                            )
            cursor.execute( ''' UPDATE Kleinstmaterial
                                SET bestellt = ?
                                WHERE MatNr = ?
                            ''', (status, mat_number)
                           )
        else:
            datum = datetime.datetime.strftime(datetime.datetime.now(), r'%d.%m.%Y %H:%M')
            #datum = datetime.datetime.now()
            cursor.execute( ''' UPDATE Standardmaterial
                                SET bestellt = ?,
                                    Datum = ?
                                WHERE MatNr = ? 
                            ''', (status, datum, mat_number)
                            )
            cursor.execute( ''' UPDATE Kleinstmaterial
                                SET bestellt = ?,
                                    Datum = ?
                                WHERE MatNr = ?
                            ''',(status, datum, mat_number)
                            )   
    connection.commit()
    show_critical_material(app)    


def show_stock(app: application.App) -> None:
    app.disabled_button.config(state = 'enabled')
    app.disabled_button = app.button_bestand
    app.button_bestand.config(state = 'disabled')
    
    enabled = [app.matnr_entry]
    disabled = [app.posnr_entry, app.sm_entry, app.bestellt_button, app.bestellt_label, app.delete_button, app.print_button, app.combobox_loeschen]
    set_widget_status(enabled, disabled)
    unbind_all_widgets(app)
    app.matnr_entry.bind('<Return>', lambda _ : show_stock(app))
    app.matnr_entry.bind('<KeyRelease>', lambda _: show_stock(app))
    app.matnr_entry.focus()

    cursor = app.cursor
    matnr_entry = app.matnr_entry
    matnr = matnr_entry.get()
    
    ingoing = cursor.execute("SELECT * FROM Wareneingang WHERE MatNr LIKE ?", (f'%{matnr}%',)).fetchall()
    ingoing_dict = defaultdict(int)
    for row in ingoing:
        ingoing_dict[row['MatNr']] += row['Menge']
    
    outgoing = cursor.execute("SELECT * FROM Warenausgang WHERE MatNr LIKE ?", (f'%{matnr}%',)).fetchall()
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

    outgoing_small_material = cursor.execute("SELECT * FROM Warenausgang_Kleinstmaterial_ohne_SM_BEZUG").fetchall()
    for row in outgoing_small_material:
        outgoing_dict[row['MatNr']] += row['Menge']

    
    
    standardmaterial_list = []
    small_material_list = []
    for matnr, menge in ingoing_dict.items():
        bestand = menge - outgoing_dict.get(matnr, 0)
        bezeichnung = app.materialnames_dict[matnr]
        einheit = app.units_dict[matnr]
        if matnr in app.standard_materials:
            standardmaterial_list.append((matnr, bezeichnung, bestand, einheit))
        elif matnr in app.small_materials:
            small_material_list.append((matnr, bezeichnung, bestand, einheit))
    
    for item in app.output_listbox.get_children():
        app.output_listbox.delete(item)
    columns = ['MatNr.', 'Bezeichnung', 'Bestand', 'Einheit', 'LAST_COLUMN']
    app.output_listbox.configure(columns = columns)
    for column in columns:
        if column == 'LAST_COLUMN':
            continue
        width, anchor = app.columns_dict[column]
        app.output_listbox.heading(column, text = column)
        app.output_listbox.column(column, width = width, anchor = anchor, stretch = tk.NO)
    app.output_listbox.column('#0', width = 200, stretch = tk.NO) #erste Spalte fixieren (dort, wo das Parent erscheint)
    if standardmaterial_list:
        tree_standard = app.output_listbox.insert('', 'end', text = 'Standardmaterial', open = True, tags = ('green',))
        for entry in sorted(standardmaterial_list, key = lambda x: int(x[0])):
            app.output_listbox.insert(tree_standard, "end", values = entry)
    if small_material_list:
        tree_small_material = app.output_listbox.insert('', 'end', text = 'Kleinstmaterial', open = True, tags = ('green',))
        for entry in sorted(small_material_list, key = lambda x: int(x[0])):
            app.output_listbox.insert(tree_small_material, "end", values = entry)
    
    if not standardmaterial_list and not small_material_list:
        app.output_listbox.insert('','end', values = ('','Keine Daten gefunden. Bitte Filter prüfen'))
    
     
def show_ingoing_material(app:application.App) -> None:
    app.disabled_button.config(state = 'enabled')
    app.disabled_button = app.button_wareneingang
    app.button_wareneingang.config(state = 'disabled')
    
    enabled = [app.matnr_entry]
    disabled = [app.sm_entry, app.posnr_entry, app.bestellt_label, app.bestellt_button, app.delete_button, app.print_button, app.combobox_loeschen]
    set_widget_status(enabled, disabled)
    unbind_all_widgets(app)
    app.matnr_entry.focus()
    app.matnr_entry.bind('<Return>', lambda _ : show_ingoing_material(app))
    app.matnr_entry.bind('<KeyRelease>', lambda _: show_ingoing_material(app))
    
    cursor:sqlite3.Cursor = app.cursor
    matnr = app.matnr_entry.get()
    selection = cursor.execute("SELECT * FROM Wareneingang WHERE MatNr LIKE ?", (f'%{matnr}%',)).fetchall()
    
    for item in app.output_listbox.get_children():
        app.output_listbox.delete(item)
    
    columns = ['ID', 'MatNr.', 'Bezeichnung', 'Menge', 'Einheit', 'Datum', 'LAST_COLUMN']
    app.output_listbox.configure(columns = columns)
    for column in columns:
        if column == 'LAST_COLUMN':
            continue
        width, anchor = app.columns_dict[column]
        app.output_listbox.heading(column, text = column)
        app.output_listbox.column(column, width = width, anchor = anchor, stretch = tk.NO)
    
    for entry in selection[::-1]:
        date = datetime.datetime.strftime(datetime.datetime.strptime(entry['Datum'], r"%Y-%m-%d %H:%M:%S"), r"%d.%m.%Y %H:%M:%S")
        app.output_listbox.insert('', "end", values = (entry['ID'],
                                                       entry['MatNr'],
                                                       entry['Bezeichnung'],
                                                       entry['Menge'],
                                                       app.units_dict[entry['MatNr']],
                                                       date
                                                       ))
    if not selection:
        app.output_listbox.insert('', 'end', values = ('','','Keine Daten gefunden. Bitte Filter prüfen.'))


def show_outgoing_material(app:application.App) -> None:
    app.disabled_button.config(state = 'enabled')
    app.disabled_button = app.button_warenausgang
    app.button_warenausgang.config(state = 'disabled')
    
    enabled = [app.sm_entry, app.matnr_entry]
    disabled = [app.posnr_entry, app.bestellt_label, app.bestellt_button, app.delete_button, app.print_button, app.combobox_loeschen]
    set_widget_status(enabled, disabled)
    unbind_all_widgets(app)
    
    app.sm_entry.bind('<Return>', lambda _ : show_outgoing_material(app))
    app.sm_entry.bind('<KeyRelease>', lambda _ : show_outgoing_material(app))
    app.matnr_entry.bind('<Return>', lambda _ : show_outgoing_material(app))
    app.matnr_entry.bind('<KeyRelease>', lambda _ : show_outgoing_material(app))
    
    if str(app.matnr_entry.focus_get()) != '.!frame3.!entry2': # .!frame3.!entry2 ist der Name vom SM_ENTRY Widget (2.entry im 3. frame)
        app.matnr_entry.focus()
    
    cursor:sqlite3.Cursor = app.cursor
    matnr = app.matnr_entry.get()
    smnr = app.sm_entry.get()
    selection_with_sm = cursor.execute("SELECT * FROM Warenausgang WHERE MatNr LIKE ? AND SM_Nummer LIKE ? AND PosTyp = 9", (f'%{matnr}%', f'%{smnr}%',)).fetchall()
    selection_without_sm = cursor.execute("SELECT * FROM Warenausgang_Kleinstmaterial_ohne_SM_Bezug WHERE MatNr LIKE ?", (f'%{matnr}%',)).fetchall()

    for item in app.output_listbox.get_children():
        app.output_listbox.delete(item)
    
    columns = ['Nr.', 'SM Nummer / ID', 'Position', 'Pos.Typ', 'MatNr.', 'Bezeichnung', 'Menge', 'Einheit', 'Lieferschein', 'SD Beleg',  'Materialbeleg', 'LAST_COLUMN']
    app.output_listbox.configure(columns = columns)
    for column in columns:
        if column == 'LAST_COLUMN':
            continue
        width, anchor = app.columns_dict[column]
        app.output_listbox.heading(column, text = column)
        app.output_listbox.column(column, width = width, anchor = anchor, stretch = tk.NO)
    app.output_listbox.column('#0', width = 280, stretch = tk.NO) #erste Spalte fixieren (dort, wo das Parent erscheint)
    
    if selection_with_sm:
        selection_with_sm = sorted(selection_with_sm, key = lambda x : x['SM_Nummer'])
        tree_with_sm = app.output_listbox.insert('', 'end', text = 'Warenausgang mit SM Bezug', open = True, tags = ('green',))
        for number, entry in enumerate(selection_with_sm, start = 1):
            app.output_listbox.insert(tree_with_sm, "end", values = (number,
                                                                    entry['SM_Nummer'],
                                                                    entry['Position'],
                                                                    entry['PosTyp'],
                                                                    entry['MatNr'],
                                                                    entry['Bezeichnung'],                                                                
                                                                    entry['Bedarfsmenge'],
                                                                    app.units_dict[entry['MatNr']],
                                                                    entry['Lieferschein'][:-2],
                                                                    entry['SD_Beleg'][:-2],
                                                                    entry['Materialbeleg'][:-2]
                                                                    ))
    if selection_without_sm and not smnr:
        tree_without_sm = app.output_listbox.insert('', 'end', text = 'Warenausgang ohne SM Bezug', open = True, tags = ('green',))
        for number, entry in enumerate(selection_without_sm, start = len(selection_with_sm)+1):
            app.output_listbox.insert(tree_without_sm, "end", values = (number,
                                                                    entry['ID'],
                                                                    '',
                                                                    '',
                                                                    entry['MatNr'],
                                                                    entry['Bezeichnung'],
                                                                    entry['Menge'],
                                                                    app.units_dict[entry['MatNr']]
                                                                    ))
    if not selection_with_sm and not selection_without_sm:
        app.output_listbox.insert('','end', values = ('','','','','','Keine Daten gefunden. Bitte Filter prüfen'))


def show_material_for_order(app:application.App) -> None:
    app.disabled_button.config(state = 'enabled')
    app.disabled_button = app.button_sm_auftrag
    app.button_sm_auftrag.config(state = 'disabled')
    
    enabled = [app.sm_entry, app.print_button]
    disabled = [app.posnr_entry, app.matnr_entry, app.bestellt_label, app.bestellt_button, app.delete_button, app.combobox_loeschen]
    set_widget_status(enabled, disabled)
    unbind_all_widgets(app)
    app.sm_entry.bind('<Return>', lambda _ : show_material_for_order(app))
    app.sm_entry.bind('<KeyRelease>', lambda _ : show_material_for_order(app))
    app.sm_entry.focus()

    cursor:sqlite3.Cursor = app.cursor
    smnr = app.sm_entry.get()
    selection = cursor.execute("SELECT * FROM Warenausgabe_Comline WHERE SM_Nummer LIKE ?", (f'%{smnr}%',)).fetchall()
    selection_addresses = cursor.execute("SELECT * FROM Adresszuordnung WHERE SM_Nummer LIKE ?", (f'%{smnr}%',)).fetchall()
    standard_material = []
    small_material = []
    telekom_material = []
    for row in selection:
        values = (row['SM_Nummer'], row['MatNr'], row['Bezeichnung'], row['Bedarfsmenge'], app.units_dict[row['MatNr']] or 'n/a')
        if row['MatNr'] in app.standard_materials:
            standard_material.append(values)
        elif row['MatNr'] in app.small_materials:
            small_material.append(values)
        else:
            telekom_material.append(values)
    address_values = []
    for row in selection_addresses:
        values = (row['SM_Nummer'], row['VPSZ'], row['Adresse'])
        address_values.append(values)

    for item in app.output_listbox.get_children():
        app.output_listbox.delete(item)
    
    columns = ['SM Nummer', 'MatNr.', 'Bezeichnung', 'Menge', 'Einheit', 'LAST_COLUMN']
    app.output_listbox.configure(columns = columns)
    for column in columns:
        if column == 'LAST_COLUMN':
            continue
        width, anchor = app.columns_dict[column]
        app.output_listbox.heading(column, text = column)
        app.output_listbox.column(column, width = width, anchor = anchor, stretch = tk.NO)
    app.output_listbox.column('#0', width = 250, stretch = tk.NO) #erste Spalte fixieren (dort, wo das Parent erscheint)
    
    if not selection:
        app.output_listbox.insert('','end', values = ('','','Keine Daten gefunden. Bitte Filter prüfen.'))
        return
    if address_values:
        tree_address = app.output_listbox.insert('', 'end', text = 'Adressen', open = True, tags = ('green',))
        for entry in address_values:
            app.output_listbox.insert(tree_address, "end", values = entry)
    
    if standard_material:
        tree_standard = app.output_listbox.insert('', 'end', text = 'Standardmaterial', open = True, tags = ('green',))
        for entry in standard_material:
            app.output_listbox.insert(tree_standard, "end", values = entry)
    if small_material:
        tree_small_material = app.output_listbox.insert('', 'end', text = 'Kleinstmaterial', open = True, tags = ('green',))
        for entry in small_material:
            app.output_listbox.insert(tree_small_material, "end", values = entry)    
    if telekom_material:
        tree_small_material = app.output_listbox.insert('', 'end', text = 'Telekommaterial', open = True, tags = ('green',))
        for entry in telekom_material:
            app.output_listbox.insert(tree_small_material, "end", values = entry)  


def print_screen(app:application.App) -> None:
    file_path = f"{os.getcwd()}\\Ausgabe.xlsx"
    if len(app.sm_entry.get()) < 4:
        ctypes.windll.user32.MessageBoxW(0,
                                         "Um Fehlausdrucke zu vermeiden müssen im Filter SM-Nummer \nmindestens 4 Zeichen eingetragen sein.", 
                                         "Zu wenig Zeichen bei SM Nummer", 
                                         64)
        return
    answer = ctypes.windll.user32.MessageBoxW(0,"Soll der angezeigte Inhalt gedruckt werden?", "Drucken...", 68)
    if answer == 6:
        standard_material = []
        small_material = []
        telekom_material = []
        address = []
        parents = app.output_listbox.get_children()
        for parent in parents:
            children = app.output_listbox.get_children(parent)
            for child in children:
                values = app.output_listbox.item(child)['values']
                values = [str(x) for x in values]
                output = '\t'.join(values)
                if app.output_listbox.item(parent)['text'] == 'Standardmaterial':
                    standard_material.append(output)
                elif app.output_listbox.item(parent)['text'] == 'Kleinstmaterial':
                    small_material.append(output)
                elif app.output_listbox.item(parent)['text'] == 'Adressen':
                    address.append(output)
                else:
                    telekom_material.append(output)
        wb:xl.Workbook = xl.Workbook()
        sheet = wb.active
        sheet.page_setup.orientation = 'landscape'
        sheet.page_setup.fitToPage = True
        sheet.column_dimensions['A'].width = 12
        sheet.column_dimensions['B'].width = 16
        sheet.column_dimensions['C'].width = 44
        sheet.column_dimensions['D'].width = 7
        sheet.column_dimensions['E'].width = 7
        sheet.column_dimensions['F'].width = 3
        for x in range(7,37):
            if x % 2 == 1:
                sheet.column_dimensions[xl.utils.cell.get_column_letter(x)].width = 1
            else:
                sheet.column_dimensions[xl.utils.cell.get_column_letter(x)].width = 3
        start = 1
        if address:
            sheet.cell(column = 1, row = start,value = 'Adresse')
            sheet.cell( column = 1, row = start).font = xl.styles.Font(size = 12, bold = True)
            start +=1
            for row, line in enumerate(address, start = start):
                for col, word in enumerate(line.split('\t'), start = 1):
                    sheet.cell(column = col, row = row, value = word)
            start += len(address) + 1
        
        if standard_material:
            sheet.cell(column = 1, row = start,value = 'Standardmaterial')
            sheet.cell( column = 1, row = start).font = xl.styles.Font(size = 12, bold = True)
            start +=1
            for row, line in enumerate(standard_material, start = start):
                for col, word in enumerate(line.split('\t'), start = 1):
                    sheet.cell(column = col, row = row, value = word)
            start += len(standard_material) + 1
        if small_material:
            sheet.cell(column = 1, row = start, value = 'Kleinstmaterial')
            sheet.cell( column = 1, row = start).font = xl.styles.Font(size = 12, bold = True)
            start +=1
            for row, line in enumerate(small_material, start = start):
                for col, word in enumerate(line.split('\t'), start = 1):
                    sheet.cell(column =col, row = row, value = word)
            start += len(small_material) + 1
        if telekom_material:
            sheet.cell(column = 1, row = start, value = 'Telekommaterial')
            sheet.cell( column = 1, row = start).font = xl.styles.Font(size = 12, bold = True)
            start +=1
            for row, line in enumerate(telekom_material, start = start):
                for col, word in enumerate(line.split('\t'), start = 1):
                    sheet.cell(column =col, row = row, value = word)
        border = xl.styles.Side(border_style = 'thin', color = '000000')
        for idx, row in enumerate(sheet, start = 1):
            if not row[0].value:
                continue
            try:
                # test for number in first column
                _ = int(row[0].value)
                # if no type in column 5, ignore this row
                if not (sheet.cell(row= idx, column = 5).value):
                    for col in range(1,4):
                            sheet.cell(row = idx, column = col).font = xl.styles.Font(color="3282F6")
                    continue
                # draw cell bottom lines
                for col in range(1,6):
                    sheet.cell(row = idx, column = col).border = xl.styles.Border(bottom = border)
                # change font color to red when amount > 1 and unit is M or n/a
                try:
                    if int(sheet.cell(row = idx, column = 4).value) > 1 and sheet.cell(row= idx, column = 5).value.strip() in ('M', 'n/a'):
                        for col in range(1,6):
                            sheet.cell(row = idx, column = col).font = xl.styles.Font(color="FF0000", bold = True)
                except TypeError:
                    pass
                if sheet.cell(row= idx, column = 5).value.strip() in ('ST', 'SA', 'PAK'):
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
        finally:
            wb.Close(False)
            excel_app.Quit()
            os.remove(file_path)

        
def book_outgoing_from_excel_file(app:application.App) -> None:
    
    app.disabled_button.config(state = 'enabled')
    app.disabled_button = app.button_warenausgang_buchen
    enabled = []
    disabled = [app.sm_entry, app.posnr_entry, app.matnr_entry, app.delete_button, app.print_button, app.combobox_loeschen,app.bestellt_button, app.bestellt_label, ]
    set_widget_status(enabled, disabled)
    # ctypes.windll.user32.MessageBoxW(0,"Bitte im nächsten Fenster die exportierte Materialliste aus PSL auswählen.", "Warenausgang buchen...", 64)
    cursor:sqlite3.Cursor = app.cursor
    connection:sqlite3.Connection = app.connection
    path_to_excel = filedialog.askopenfilenames(defaultextension = 'xlsx', title = 'Bitte die EXPORT Datei aus PSL wählen - Mehrfachauswahl ist möglich...')
    if not path_to_excel:
        show_critical_material(app)
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
    enabled = []
    disabled = [app.sm_entry, app.posnr_entry, app.matnr_entry, app.delete_button, app.print_button, app.combobox_loeschen,app.bestellt_button, app.bestellt_label, ]
    set_widget_status(enabled, disabled)
    title = 'Wareneingang buchen (Anlieferung CTDI)' 
    selection = 'SELECT * FROM Kleinstmaterial UNION SELECT * FROM Standardmaterial'
    app.open_booking_window(title, selection)
    if app.user_closed_window:
        show_ingoing_material(app)
        return
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
    enabled = []
    disabled = [app.sm_entry, app.posnr_entry, app.matnr_entry, app.delete_button, app.print_button, app.combobox_loeschen,app.bestellt_button, app.bestellt_label, ]
    set_widget_status(enabled, disabled)
    
    title = 'Kleinstmaterial ohne SM Bezug herausgeben'
    selection = 'SELECT * FROM Kleinstmaterial'
    app.open_booking_window(title, selection)
    if app.user_closed_window:
        show_critical_material(app)
        return
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
    table = app.combobox_loeschen.get()
    execution_string = app.execution_dict[app.combobox_loeschen.get()][0]
    execute_values = app.execution_dict[app.combobox_loeschen.get()][1]
    execute_values = execute_values.replace('matnr', matnr).replace('sm', sm).replace('posnr', posnr)
    execution_tuple = tuple(execute_values.split(','))
    execution_tuple = tuple([entry for entry in execution_tuple if entry])
    if not execution_tuple:
        execution_tuple = ('',)    

    # app.execution_string = execution_string
    # app.execution_tuple = execution_tuple
    
    for item in app.output_listbox.get_children():
        app.output_listbox.delete(item)
    columns_info = cursor.execute(f"PRAGMA table_info ({table})").fetchall()
    columns = [info[1] for info in columns_info]
    columns.append('LAST_COLUMN')
    app.output_listbox.configure(columns = columns)
    for column in columns:
        if column == 'LAST_COLUMN':
            continue
        width, anchor = app.columns_dict[column]
        app.output_listbox.heading(column, text = column)
        app.output_listbox.column(column, width = width, anchor = anchor, stretch = tk.NO)
    app.output_listbox.column('#0', width = 250, stretch = tk.NO) #erste Spalte fixieren (dort, wo das Parent erscheint)
    
    selection = cursor.execute(execution_string, (execution_tuple)).fetchall()
    for row in selection:
        values = [x for x in row]
        app.output_listbox.insert('', 'end', values = values)


def delete_selected_entries(app:application.App) -> None:
    selections =  app.output_listbox.selection()
    if not selections:
        ctypes.windll.user32.MessageBoxW(0,
                                         "Es sind keine Einträge markiert.\nNur markierte Einträge werden gelöscht.", 
                                         "Markier was..", 
                                         64)
        return
    answer = ctypes.windll.user32.MessageBoxW(0,
                                              "Achtung!!\nDie markierten Einträge werden unwiderruflich gelöscht.\nBist du sicher, dass du die Einträge löschen willst?", 
                                              "Einträge löschen", 
                                              68)
    if answer != 6:
        return
    
    for selection in selections:
        values = app.output_listbox.item(selection, 'values')
        if not values:
            continue
        try:
            matnr = app.output_listbox.set(selection, column = 'MatNr')
        except tk.TclError:
            matnr = ''
        try:
            sm = app.output_listbox.set(selection, column = 'SM_Nummer')
        except tk.TclError:
            sm = ''
        try:
            posnr = app.output_listbox.set(selection, column = 'ID')
        except tk.TclError:
            try:
                posnr = app.output_listbox.set(selection, column = 'Position')
            except tk.TclError:
                posnr = ''

        execution_string = app.deletion_dict[app.combobox_loeschen.get()][0]
        execute_values = app.deletion_dict[app.combobox_loeschen.get()][1]
        execute_values = execute_values.replace('matnr', matnr).replace('sm', sm).replace('posnr', posnr)
        execution_tuple = tuple(execute_values.split(','))
        execution_tuple = tuple([entry for entry in execution_tuple if entry])
        if not execution_tuple:
            execution_tuple = ('',)    
        app.connection.execute(execution_string, (execution_tuple))
    app.connection.commit()
    filter_entries_to_delete(app)

   
def confirm_user_input(app:application.App):
    app.stck_entry.unbind('<Return>') 
    app.user_closed_window = False
    app.ingoing_window.quit() 
    app.ingoing_window.destroy()


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


def read_adresses_from_workorder_list(connection:sqlite3.Connection, cursor:sqlite3.Cursor) ->None:
    path_to_excel = filedialog.askopenfilenames(initialfile = 'ExcelPYTHON.xlsx', 
                                        defaultextension = '.db', 
                                        filetypes = [('Excel', '*.xlsx'), ('All files', '*.*') ])
    if not path_to_excel:
        return
    for file in path_to_excel:
        excel_file = pd.read_excel(file)
        
        all_entries_in_db = cursor.execute('SELECT SM_Nummer FROM Adresszuordnung').fetchall()
        all_known_sm_numbers = [row['SM_Nummer'] for row in all_entries_in_db]
        new_addresses = []
        seen = set()
        for _, line in excel_file.iterrows():
            if pd.isnull(line['Was']):
                continue
            if 'OLT' in line['Was']:
                if str(line['SM']) in all_known_sm_numbers or line['SM'] in seen:
                    continue
                new_addresses.append((str(line['SM']), line['VPSZ'], line['ORT']))
                seen.add(line['SM'])
    for smnr, vpsz, address in new_addresses:
        cursor.execute('''
                            INSERT INTO 
                                Adresszuordnung (SM_Nummer, VPSZ, Adresse)
                            VALUES
                                (?,?,?)
                        ''', (smnr, vpsz, address)
                            )
        print(smnr, vpsz, address)
    connection.commit()
    print(f"Done. Added {len(new_addresses)} new addresses.")
