import sqlite3
import tkinter as tk
import tkinter.ttk as ttk
import functions as fc


class App:
    def __init__(self, connection:sqlite3.Connection, cursor:sqlite3.Cursor, user:str, path_to_db:str):
        self.user = user
        self.connection = connection
        self.cursor = cursor
        self.path_to_db = path_to_db
        self.window = tk.Tk()
        self.window.title('Lagerverwaltung Comline')
        self.window.minsize(1200,768)
        self.window.configure(bg='black')
        self.execution_string = ''
        self.execution_tuple = ''
        
        
        self.style = ttk.Style()
        # print(self.style.theme_names())
        self.style.theme_use('classic') # theme 'alt' nutzen und ändern
        # Verhalten der Buttons festlegen - Farbänderung beim Hovern etc.
        self.style.map('Comline.TButton', 
                    background=[('active','forest green'),('disabled', 'forest green')],
                    foreground=[('active', 'black'), ('disabled', 'black')],
                    relief=[('pressed', '!disabled', 'sunken'), ('disabled', 'sunken')],
                    )

        self.style.map('Comline2.TButton', 
                    width = [('disabled', 23)],
                    justify = [('disabled', tk.CENTER)],
                    background=[('active','forest green'), ('disabled', 'black')],
                    foreground=[('active', 'black'), ('disabled', 'black')],
                    borderwidth = [('disabled', 0)],
                    relief=[('pressed', '!disabled', 'sunken'),('disabled', 'flat')],
                    highlightcolor = [('disabled', 'black')]
                    )
        
        
        self.style.map('Blue.TButton', 
                    background=[('active','deep sky blue'), ('disabled', 'deep sky blue')],
                    foreground=[('active', 'black'), ('disabled', 'black')],
                    relief=[('pressed', '!disabled', 'sunken'),('disabled', 'sunken')],
                    )
        
        
        self.style.map('Blue2.TButton', 
                    background=[('active','deep sky blue'), ('disabled', 'black')],
                    foreground=[('active', 'black'), ('disabled', 'black')],
                    borderwidth = [('disabled', 0)],
                    relief=[('pressed', '!disabled', 'sunken'),('disabled', 'flat')],
                    highlightcolor = [('disabled', 'black')]
                    )

        
        self.style.map('Red.TButton', 
                    background=[('active','orange red'), ('disabled', 'orange red')],
                    foreground=[('active', 'black'), ('disabled', 'black')],
                    relief=[('pressed', '!disabled', 'sunken'), ('disabled', 'sunken')],
                    )
        
        
        self.style.map('Red_Delete.TButton', 
                    width = [('disabled', 23)],
                    font = [('disabled', 'Verdana 10')],
                    justify = [('disabled', tk.CENTER)],
                    background=[('active','orange red'), ('disabled', 'black')],
                    foreground=[('active', 'black'), ('disabled', 'black')],
                    relief=[('pressed', '!disabled', 'sunken'), ('disabled', 'flat')],
                    borderwidth = [('disabled', 0)],
                    highlightcolor = [('disabled', 'black')]
                    )
    
        self.style.map('Blue.TLabel', 
                       foreground=[('disabled', 'black')],
                       background=[('disabled', 'black')],
                        )
        self.style.map('Comline.TEntry', fieldbackground = [('disabled', 'grey10')])

        self.style.map('Red.TCombobox',
                       foreground = [('active', 'orange red'), ('disabled', 'black')],
                       background = [('active', 'black'), ('disabled', 'black')],
                       borderwidth = [('disabled', 0)],
                       highlightcolor = [('disabled', 'black')],
                       fieldbackground = [('disabled', 'black')]
                       )


        # grundlegende Sachen konfigurieren - Schriftart, -größe etc.
        self.style.configure('Comline.TButton',  font = 'Verdana 10',foreground = 'forest green', background = 'black', justify = tk.CENTER)
        self.style.configure('Comline2.TButton',  font = 'Verdana 10',foreground = 'forest green', background = 'black', justify = tk.CENTER)
        self.style.configure('Blue.TButton', font = 'Verdana 10',foreground = 'deep sky blue', background = 'black', justify = tk.CENTER)
        self.style.configure('Blue2.TButton', font = 'Verdana 10',foreground = 'deep sky blue', background = 'black', justify = tk.CENTER)
        self.style.configure('Red.TButton', font = 'Verdana 10',foreground = 'orange red', background = 'black', justify = tk.CENTER)
        self.style.configure('Red_Delete.TButton', font = 'Verdana 10',foreground = 'orange red', background = 'black', justify = tk.CENTER)
        self.style.configure('Comline.TLabel', font ='Verdana 10', foreground = 'forest green', background = 'black')
        self.style.configure('Comline2.TLabel', font ='Verdana 10 bold', foreground = 'forest green', background = 'black')
        self.style.configure('Red.TLabel', font ='Verdana 10 bold', foreground = 'orange red', background = 'black')
        self.style.configure('Blue.TLabel', font ='Verdana 10', foreground = 'deep sky blue', background = 'black')
        self.style.configure('Comline.TEntry', font ='Verdana 10 bold', foreground = 'black', fieldbackground = 'white')
        self.style.configure('Frame_grey.TFrame', borderwidth = 2, bordercolor = 'grey', background = 'black', relief = 'groove')
        self.style.configure('Frame_grey2.TFrame', borderwidth =2, background = 'black', relief = 'groove')
        self.style.configure('Frame_filter.TFrame', borderwidth = 2, bordercolor = 'grey', background = 'black', relief = 'groove')
        self.style.configure('Comline.TCombobox', background = 'black', foreground = 'forest green', font = 'Verdana 10')
        self.style.configure('Red.TCombobox', background = 'black', foreground = 'orange red', font = 'Verdana 10')
        #print(self.style.layout('Frame_grey2.TFrame'))
        #print(self.style.element_options('Frame.border'))

        # Überschrift Frame und Label erstellen
        self.frame_top = ttk.Frame(self.window, style = 'Frame_grey2.TFrame')
        self.label_top = ttk.Label(self.frame_top, 
                                   style = 'Comline2.TLabel', 
                                   text = 'Lagerverwaltung Comline',
                                   anchor = 'e')
        self.label_top_path_to_db = ttk.Label(self.frame_top, 
                                   style = 'Red.TLabel', 
                                   text = f'Pfad zur Datenbank: {self.path_to_db}',
                                   anchor = 'w')
        
        self.button_change_db = ttk.Button(self.frame_top, 
                                           style = 'Comline.TButton',
                                           text = 'Pfad ändern',
                                           command = lambda : fc.change_db_path(self))
        
        self.button_change_db.pack(side = tk.RIGHT, padx = 50, pady = 5)
        self.label_top.pack(side = tk.LEFT,padx = 100, pady = 10, fill = 'x', expand = True)
        self.label_top_path_to_db.pack(side = tk.RIGHT, padx = 10, pady = 10, fill = 'x', expand = True)

        # Buttonzeile -  Frames und Buttons erstellen
        self.button_frame = ttk.Frame(self.window, style = 'Frame_grey.TFrame')
        self.frame_button_left = ttk.Frame(self.button_frame, style = 'Frame_grey.TFrame')
        self.button_krit_mat = ttk.Button(self.frame_button_left, 
                                    style = 'Comline.TButton', 
                                    text = 'Kritisches Material \nanzeigen', 
                                    command = lambda: fc.show_critical_material(self)
                                    )
        self.button_krit_mat.pack(side = tk.LEFT, padx = 5, pady = 5)
        self.button_bestand = ttk.Button(self.frame_button_left, 
                                    style = 'Comline.TButton', 
                                    text = 'Bestand\nanzeigen', 
                                    command = lambda:fc.show_stock(self)
                                    )
        self.button_bestand.pack(side = tk.LEFT, padx = 5, pady = 5)
        self.button_wareneingang = ttk.Button(self.frame_button_left, 
                                              style = 'Comline.TButton', 
                                              text = 'Wareneingang\nanzeigen',
                                              command = lambda:fc.show_ingoing_material(self)
                                              )
        self.button_wareneingang.pack(side = tk.LEFT, padx = 5, pady = 5)
        self.button_warenausgang = ttk.Button(self.frame_button_left, 
                                              style = 'Comline.TButton', 
                                              text = 'Warenausgang\nanzeigen',
                                              command = lambda:fc.show_outgoing_material(self)
                                              )
        self.button_warenausgang.pack(side = tk.LEFT, padx = 5, pady = 5)
        self.button_sm_auftrag = ttk.Button(self.frame_button_left,
                                            style = 'Comline.TButton',
                                            text = 'Material für\nSM-Auftrag anzeigen',
                                            command = lambda:fc.show_material_for_order(self)
                                            )
        self.button_sm_auftrag.pack(side = tk.LEFT, padx = 5, pady = 5)
        self.frame_button_left.pack(side = tk.LEFT, padx = 5, pady = 5)

        # ADMIN Knopf und Combobox
        self.frame_button_right = ttk.Frame(self.button_frame, style = 'Frame_grey.TFrame')
        self.combobox_loeschen = ttk.Combobox(self.frame_button_right, 
                                                style = 'Red.TCombobox',
                                                width = 50, 
                                                values = ['Kleinstmaterial', 'Standardmaterial', 'Warenausgabe_Comline',
                                                        'Warenausgang', 'Wareneingang'])
        self.combobox_loeschen.current(0)
        self.button_loeschen = ttk.Button(self.frame_button_right, 
                                            style = 'Red.TButton', 
                                            text = 'Einträge aus \nDatenbank löschen',
                                            command  = lambda: fc.filter_entries_to_delete(self))
        self.button_loeschen.pack(side = tk.LEFT, padx = 5, pady = 5)
        self.combobox_loeschen.pack(side = tk.LEFT, padx = 5, pady = 5)
        
        
        
        if self.user == fc.ADMIN:
            self.frame_button_right.pack(side = tk.RIGHT, padx = 5, pady = 5)

        if self.user in (fc.ADMIN, fc.EXPERT):
            self.frame_button_middle = ttk.Frame(self.button_frame, style = 'Frame_grey.TFrame')
            self.button_wareneingang_buchen = ttk.Button(self.frame_button_middle, 
                                                         style = 'Blue.TButton', 
                                                         text = 'Wareneingang buchen\n(Anlieferung CTDI)',
                                                         command = lambda:fc.book_ingoing_position(self))
            self.button_wareneingang_buchen.pack(side = tk.LEFT, padx = 5, pady = 5)
            self.button_warenausgang_buchen = ttk.Button(self.frame_button_middle, 
                                                         style = 'Blue.TButton', 
                                                         text = 'Warenausgang buchen\n(Excel aus PSL)',
                                                         command = lambda: fc.book_outgoing_from_excel_file(self))
            self.button_warenausgang_buchen.pack(side = tk.LEFT, padx = 5, pady = 5)
            self.button_warenausgang_buchen_Kleinstmaterial = ttk.Button(self.frame_button_middle, 
                                                         style = 'Blue.TButton', 
                                                         text = 'Warenausgang buchen\n(nur Kleinstmaterial)',
                                                         command = lambda: fc.book_outgoing_kleinstmaterial(self))
            self.button_warenausgang_buchen_Kleinstmaterial.pack(side = tk.LEFT, padx = 5, pady = 5)
            if self.user == fc.ADMIN:
                self.frame_button_middle.pack(side = tk.TOP, padx = 5, pady = 5)
            else:
                self.frame_button_middle.pack(side = tk.LEFT, padx = 5, pady = 5)

        # Filter Frames und Felder erstellen
        self.filter_frame = ttk.Frame(self.window, style = 'Frame_filter.TFrame')
        
        # button und entry anlegen, aber nur im Adminmodus packen
        self.posnr_entry = ttk.Entry(self.filter_frame, style = 'Comline.TEntry')
        
            
        # Filterbutton, die immer angezeigt werden
        self.matnr_label = ttk.Label(self.filter_frame, style = 'Comline.TLabel', text = ' Daten filtern:     Materialnummer:')
        self.matnr_label.pack(side = tk.LEFT, padx = 5, pady = 5)
        self.matnr_entry = ttk.Entry(self.filter_frame, style = 'Comline.TEntry')
        self.matnr_entry.pack(side = tk.LEFT, padx = 5, pady = 5)

        self.sm_label = ttk.Label(self.filter_frame, style = 'Comline.TLabel', text = 'SM Nummer:')
        self.sm_label.pack(side = tk.LEFT, padx = 5, pady = 5)
        self.sm_entry = ttk.Entry(self.filter_frame, style = 'Comline.TEntry')
        self.sm_entry.pack(side = tk.LEFT, padx = 5, pady = 5)

        if self.user == fc.ADMIN:
            self.posnr_label = ttk.Label(self.filter_frame, style = 'Comline.TLabel', text = 'Position / ID:')
            self.posnr_label.pack(side = tk.LEFT, padx = 5, pady = 10)
            self.posnr_entry.pack(side = tk.LEFT, padx = 5, pady = 10)
        

        # Textausgabefenster Frame und Textbox erstellen
        self.frame_bottom = ttk.Frame(self.window, style = 'Frame_grey.TFrame', padding = 10)
        self.entry_frame = ttk.Frame(self.window, style = 'Frame_grey.TFrame')
        self.bestellt_label = ttk.Label(self.entry_frame, style = 'Blue.TLabel', text = 'Materialnummer markieren um den Bestellstatus zu wechseln')
        self.bestellt_button = ttk.Button(self.entry_frame, 
                                            style = 'Blue2.TButton', 
                                            text = 'Bestellstatus ändern',
                                            command = lambda:fc.toggle_ordered_status(self))
        
        self.print_button = ttk.Button(self.entry_frame, 
                                            style = 'Comline2.TButton', 
                                            text = 'Drucken',
                                            command = lambda:fc.print_screen(self))
        
        
        
        self.delete_button = ttk.Button(self.entry_frame, 
                                            style = 'Red_Delete.TButton', 
                                            text = 'angezeigte Daten loeschen',
                                            command = lambda:fc.delete_selected_entries(self))
        
        if self.user in (fc.ADMIN, fc.EXPERT):
            self.bestellt_label.pack(side = tk.LEFT, padx = 50, pady = 5)
            self.bestellt_button.pack(side = tk.LEFT, padx = 50, pady = 5)
        if self.user == fc.STANDARD:
               self.print_button.pack(side = tk.TOP, pady = 5)
        else:
               self.print_button.pack(side = tk.LEFT, pady = 5)
        if self.user == fc.ADMIN:    
            self.delete_button.pack(side = tk.LEFT, pady = 5)
            
        self.entry_frame.pack(side = tk.BOTTOM, fill = 'x')
        
        self.vertical_scrollbar = ttk.Scrollbar(self.frame_bottom, orient = tk.VERTICAL)
        self.horizontal_scrollbar = ttk.Scrollbar(self.frame_bottom, orient = tk.HORIZONTAL)
        
        self.output_box = tk.Text(self.frame_bottom, 
                                  font = ('Courier', 11),
                                  wrap = 'none',
                                  foreground = 'white', 
                                  background = 'grey10', 
                                  yscrollcommand = self.vertical_scrollbar.set, 
                                  xscrollcommand = self.horizontal_scrollbar.set
                                  )
        
        self.vertical_scrollbar.config(command = self.output_box.yview)
        self.vertical_scrollbar.pack(side = tk.RIGHT,  fill = 'y')
        
        self.horizontal_scrollbar.config(command = self.output_box.xview)
        self.horizontal_scrollbar.pack(side = tk.BOTTOM,  fill = 'x')
        
        self.output_box.pack(side = tk.LEFT,  expand = True, fill = 'both')

        # die Frames im Fenster positionieren
        self.frame_top.pack(side = tk.TOP, fill = 'x')
        self.button_frame.pack(side = tk.TOP, fill = 'x')
        self.filter_frame.pack(side = tk.TOP, fill = 'x')
        self.frame_bottom.pack(side = tk.BOTTOM, expand = True, fill = 'both')

        # den ersten Knopf disablen und die Werte speichern 
        self.execution_dict = {'Kleinstmaterial' : ('SELECT * FROM Kleinstmaterial WHERE MatNr LIKE ?', '%matnr%,'),
                               'Standardmaterial' : ('SELECT * FROM Standardmaterial WHERE Matnr LIKE ?', '%matnr%,'),
                               'Warenausgabe_Comline' :('SELECT * FROM Warenausgabe_Comline WHERE SM_Nummer LIKE ? AND MatNr LIKE ?', r'%sm%,%matnr%'),
                               'Warenausgang' : ('SELECT * FROM Warenausgang WHERE SM_Nummer LIKE ? AND MatNr LIKE ?', r'%sm%,%matnr%'),
                               'Wareneingang' : ('SELECT * FROM Wareneingang WHERE ID = ?', 'posnr,')}
        self.filter_dict = {'Kleinstmaterial' : ([self.matnr_entry], [self.sm_entry, self.posnr_entry]),
                            'Standardmaterial' : ([self.matnr_entry], [self.sm_entry, self.posnr_entry]),
                            'Warenausgabe_Comline' : ([self.matnr_entry, self.sm_entry], [self.posnr_entry]),
                            'Warenausgang' : ([self.matnr_entry, self.sm_entry], [self.posnr_entry]),
                            'Wareneingang' : ([self.posnr_entry], [self.sm_entry, self.matnr_entry])
                            }

        self.disabled_button = self.button_krit_mat
        fc.show_critical_material(self)
        

    def open_booking_window(self, selection):
        self.cursor.execute(selection)
        materials = self.cursor.fetchall()
        values = [f"{row['MatNr']} {row['Bezeichnung']}" for row in materials]
        self.ingoing_mat_string_var = tk.StringVar()
        self.ingoing_menge_var = tk.IntVar()
        self.ingoing_window = tk.Toplevel(master = self.window)
        self.ingoing_window.protocol("WM_DELETE_WINDOW", lambda : [self.ingoing_window.quit(), self.ingoing_window.destroy()])
        
        self.ingoing_window.minsize(500,200)
        self.ingoing_window.configure(bg = 'black')
        
        self.ueberschrift_frame = ttk.Frame(self.ingoing_window, border = 1)
        self.ueberschrift_label = ttk.Label(self.ueberschrift_frame,
                                            style = 'Comline.TLabel',
                                            anchor = tk.CENTER,
                                            text = '\nBitte Material aus Dropdown Menü wählen und Menge eingeben.\n')
        self.mat_frame = ttk.Frame(self.ingoing_window)
        self.mat_label = ttk.Label(self.mat_frame,
                                       style = 'Comline.TLabel',
                                       anchor = tk.CENTER,
                                       text = 'Material')
        self.ueberschrift_label.pack(side = tk.TOP, fill = 'x')

        self.matnr_combobox = AutocompleteCombobox(self.mat_frame,
                                           style = 'Comline.TCombobox',
                                           width = 60,
                                           textvariable=self.ingoing_mat_string_var)
        self.matnr_combobox.set_completion_list(values)
        self.mat_label.pack(side = tk.TOP, fill = 'x')
        self.matnr_combobox.pack(side = tk.TOP, fill = 'x')
        self.menge_frame = ttk.Frame(self.ingoing_window)
        self.stck_label = ttk.Label(self.menge_frame,
                                       style = 'Comline.TLabel',
                                       anchor = tk.CENTER,
                                       text = 'Menge')
        self.stck_entry = ttk.Entry(self.menge_frame,
                                    style = 'Comline.TEntry',
                                    textvariable=self.ingoing_menge_var)
        self.ok_button = ttk.Button(self.ingoing_window,
                                    style = 'Blue.TButton',
                                    text = 'buchen',
                                    command = lambda : [self.ingoing_window.quit(), self.ingoing_window.destroy()])
        self.stck_label.pack(side = tk.TOP, fill = 'x')
        self.stck_entry.pack(side = tk.TOP, fill = 'x')

        self.ueberschrift_frame.pack(side = tk.TOP, padx = 10, pady = 5, fill = 'x')
        self.ok_button.pack(side = tk.BOTTOM, fill = 'x', padx = 10, pady = 10)
        self.mat_frame.pack(side = tk.LEFT, padx = 10)    
        self.menge_frame.pack(side = tk.LEFT, padx = 10)
        self.matnr_combobox.focus()
        self.ingoing_window.mainloop()


class AutocompleteCombobox(tk.ttk.Combobox):

        def set_completion_list(self, completion_list):
                """Use our completion list as our drop down selection menu, arrows move through menu."""
                self._completion_list = sorted(completion_list, key=str.lower) # Work with a sorted list
                self._hits = []
                self._hit_index = 0
                self.position = 0
                self.bind('<KeyRelease>', self.handle_keyrelease)
                self['values'] = self._completion_list  # Setup our popup menu

        def autocomplete(self, delta=0):
                """autocomplete the Combobox, delta may be 0/1/-1 to cycle through possible hits"""
                if delta: # need to delete selection otherwise we would fix the current position
                        self.delete(self.position, tk.END)
                else: # set position to end so selection starts where textentry ended
                        self.position = len(self.get())
                # collect hits
                _hits = []
                for element in self._completion_list:
                        if element.lower().startswith(self.get().lower()): # Match case insensitively
                                _hits.append(element)
                # if we have a new hit list, keep this in mind
                if _hits != self._hits:
                        self._hit_index = 0
                        self._hits=_hits
                # only allow cycling if we are in a known hit list
                if _hits == self._hits and self._hits:
                        self._hit_index = (self._hit_index + delta) % len(self._hits)
                # now finally perform the auto completion
                if self._hits:
                        self.delete(0,tk.END)
                        self.insert(0,self._hits[self._hit_index])
                        self.select_range(self.position,tk.END)

        def handle_keyrelease(self, event):
                """event handler for the keyrelease event on this widget"""
                if event.keysym == "BackSpace":
                        self.delete(self.index(tk.INSERT), tk.END)
                        self.position = self.index(tk.END)
                if event.keysym == "Left":
                        if self.position < self.index(tk.END): # delete the selection
                                self.delete(self.position, tk.END)
                        else:
                                self.position = self.position-1 # delete one character
                                self.delete(self.position, tk.END)
                if event.keysym == "Right":
                        self.position = self.index(tk.END) # go to end (no selection)
                if len(event.keysym) == 1:
                        self.autocomplete()
                # No need for up/down, we'll jump to the popup
                # list at the position of the autocompletion
        