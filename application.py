import sqlite3
import tkinter as tk
import tkinter.ttk as ttk
import functions as fc


ADMIN = 'admin'
EXPERT = 'expert'
STANDARD = 'standard'


class App:
    '''
    This class creates the front-end for a warehouse management sqlite3 database.

    top_buttons:
        - STANDARD (all users)- 
        Pfad ändern: changes the path to the underlying sqlite3 database
        Kritisches Material anzeigen: shows material, that is below a certain threshhold
        Bestand anzeigen: shows the stock of the material
        Wareneingang anzeigen: shows all ingoing material
        Warenausgang anzeigen: shows all outgoing material
        Material für SM-Auftrag anzeigen: shows ALL materials needed for a specific work order (including external bought material)
        - EXPERT -
        Wareneingang buchen (Anlieferung CTDI): book incoming material
        Warenausgang buchen (Excel aus PSL): book outgoing material bound to a specific work order by reading an excel file created in PSL (SAP app)
        Warenausgang buchen (nur Kleinstmaterial): book outgoing material independent of work order
        - ADMIN - 
        Einträge aus Datenbank löschen: select entries you want to delete from the underlying sqlite3 database
        Combobox: choose from which underlying table you want to delete entries
    filters:
        - dependent on which button was pressed, the filters are active / inactive -
        Materialnummer: material number
        SM Nummer: work order
        Position / ID: position or ID in the underlying sqlite3 database table
    bottom_buttons:
        - STANDARD - 
        Drucken: only available after 'Material für SM Auftrag anzeigen' button was pressed
                * exports the shown text into an Excel file, formats it and prints it on the standard printer
        - EXPERT - 
        Bestellstatus ändern: only available after 'Kritisches Material anzeigen' button was pressed
                * toggles the order status from the critical material (ordered / not ordered)
        - ADMIN -
        angezeigte Daten löschen: only available after 'Einträge aus Datenbank löschen' button was pressed
                * deletes all shown (filtered) entries from the underlying sqlite3 database 
    '''
    
    
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
        self._init_app()

    def _init_app(self) -> None:
        self.style:ttk.Style = self._create_styles()
        self.top_frame:ttk.Frame = self._get_top_frame()
        self.button_frame:ttk.Frame = self._get_button_frame()
        self.filter_frame:ttk.Frame = self._get_filter_frame()
        self.bottom_frame:ttk.Frame = self._get_bottom_frame()
        # place frames into the window
        self.top_frame.pack(side = tk.TOP, fill = 'x')
        self.button_frame.pack(side = tk.TOP, fill = 'x')
        self.filter_frame.pack(side = tk.TOP, fill = 'x')
        self.bottom_frame.pack(side = tk.BOTTOM, expand = True, fill = 'both')

        # these dicts control the output and the filter settings in ADMIN - Deletion Mode 
        # depending on the selection in the combobox
        self.execution_dict = {'Kleinstmaterial' : ('SELECT * FROM Kleinstmaterial WHERE MatNr LIKE ?', r'%matnr%,'),
                               'Standardmaterial' : ('SELECT * FROM Standardmaterial WHERE Matnr LIKE ?', r'%matnr%,'),
                               'Warenausgabe_Comline' :('SELECT * FROM Warenausgabe_Comline WHERE SM_Nummer LIKE ? AND MatNr LIKE ?', r'%sm%,%matnr%'),
                               'Warenausgang' : ('SELECT * FROM Warenausgang WHERE SM_Nummer LIKE ? AND MatNr LIKE ?', r'%sm%,%matnr%'),
                               'Wareneingang' : ('SELECT * FROM Wareneingang WHERE ID = ?', 'posnr,'),
                               'Warenausgang_Kleinstmaterial_ohne_SM_Bezug':('SELECT * FROM Warenausgang_Kleinstmaterial_ohne_SM_Bezug WHERE ID = ?', 'posnr,')}
        self.filter_dict = {'Kleinstmaterial' : ([self.matnr_entry], [self.sm_entry, self.posnr_entry]),
                            'Standardmaterial' : ([self.matnr_entry], [self.sm_entry, self.posnr_entry]),
                            'Warenausgabe_Comline' : ([self.matnr_entry, self.sm_entry], [self.posnr_entry]),
                            'Warenausgang' : ([self.matnr_entry, self.sm_entry], [self.posnr_entry]),
                            'Wareneingang' : ([self.posnr_entry], [self.sm_entry, self.matnr_entry])
                            }
        # start the app with the show critical material option
        self.disabled_button = self.button_krit_mat
        fc.show_critical_material(self)
    

    def _create_styles(self) -> ttk.Style:
        '''
        create the style for all widgets 
        style-maps are used to control the font, color etc. dependent on the status of the widget
        '''
        
        _style = ttk.Style()
        _style.theme_use('classic') # use theme 'classic' and alter it
        
        _style.map('Green.TButton', 
                    background=[('active','forest green'),('disabled', 'forest green')],
                    foreground=[('active', 'black'), ('disabled', 'black')],
                    relief=[('pressed', '!disabled', 'sunken'), ('disabled', 'sunken')],
                    )

        _style.map('Green2.TButton', 
                    width = [('disabled', 23)],
                    justify = [('disabled', tk.CENTER)],
                    background=[('active','forest green'), ('disabled', 'black')],
                    foreground=[('active', 'black'), ('disabled', 'black')],
                    borderwidth = [('disabled', 0)],
                    relief=[('pressed', '!disabled', 'sunken'),('disabled', 'flat')],
                    highlightcolor = [('disabled', 'black')]
                    )
        
        _style.map('Blue.TButton', 
                    background=[('active','deep sky blue'), ('disabled', 'deep sky blue')],
                    foreground=[('active', 'black'), ('disabled', 'black')],
                    relief=[('pressed', '!disabled', 'sunken'),('disabled', 'sunken')],
                    )
        
        _style.map('Blue2.TButton', 
                    background=[('active','deep sky blue'), ('disabled', 'black')],
                    foreground=[('active', 'black'), ('disabled', 'black')],
                    borderwidth = [('disabled', 0)],
                    relief=[('pressed', '!disabled', 'sunken'),('disabled', 'flat')],
                    highlightcolor = [('disabled', 'black')]
                    )

        _style.map('Red.TButton', 
                    background=[('active','orange red'), ('disabled', 'orange red')],
                    foreground=[('active', 'black'), ('disabled', 'black')],
                    relief=[('pressed', '!disabled', 'sunken'), ('disabled', 'sunken')],
                    )
        
        _style.map('Red_Delete.TButton', 
                    width = [('disabled', 23)],
                    font = [('disabled', 'Verdana 10')],
                    justify = [('disabled', tk.CENTER)],
                    background=[('active','orange red'), ('disabled', 'black')],
                    foreground=[('active', 'black'), ('disabled', 'black')],
                    relief=[('pressed', '!disabled', 'sunken'), ('disabled', 'flat')],
                    borderwidth = [('disabled', 0)],
                    highlightcolor = [('disabled', 'black')]
                    )
    
        _style.map('Blue.TLabel', 
                    foreground=[('disabled', 'black')],
                    background=[('disabled', 'black')],
                    )
        
        _style.map('Green.TEntry', 
                    fieldbackground = [('disabled', 'grey10')]
                    )

        _style.map('Red.TCombobox',
                    foreground = [('active', 'orange red'), ('disabled', 'black')],
                    background = [('active', 'black'), ('disabled', 'black')],
                    borderwidth = [('disabled', 0)],
                    highlightcolor = [('disabled', 'black')],
                    fieldbackground = [('disabled', 'black')]
                    )
        
        _style = self._configure_styles(_style)
        return _style

    def _configure_styles(self, _style) -> ttk.Style:
        '''
        defines the basic colors, fonts etc for the widgets.
        '''
        _style.configure('Blue.TButton', font = 'Verdana 10',foreground = 'deep sky blue', background = 'black', justify = tk.CENTER)
        _style.configure('Blue.TButton', font = 'Verdana 10',foreground = 'deep sky blue', background = 'black', justify = tk.CENTER)
        _style.configure('Green2.TButton',  font = 'Verdana 10',foreground = 'forest green', background = 'black', justify = tk.CENTER)
        _style.configure('Blue2.TButton', font = 'Verdana 10',foreground = 'deep sky blue', background = 'black', justify = tk.CENTER)
        _style.configure('Red.TButton', font = 'Verdana 10',foreground = 'orange red', background = 'black', justify = tk.CENTER)
        _style.configure('Red_Delete.TButton', font = 'Verdana 10',foreground = 'orange red', background = 'black', justify = tk.CENTER)
        _style.configure('Green.TLabel', font ='Verdana 10', foreground = 'forest green', background = 'black')
        _style.configure('Green2.TLabel', font ='Verdana 10 bold', foreground = 'forest green', background = 'black')
        _style.configure('Green.TButton',  font = 'Verdana 10',foreground = 'forest green', background = 'black', justify = tk.CENTER)
        _style.configure('Red.TLabel', font ='Verdana 10 bold', foreground = 'orange red', background = 'black')
        _style.configure('Blue.TLabel', font ='Verdana 10', foreground = 'deep sky blue', background = 'black')
        _style.configure('Green.TEntry', font ='Verdana 10 bold', foreground = 'black', fieldbackground = 'white')
        _style.configure('Frame_grey.TFrame', borderwidth = 2, bordercolor = 'grey', background = 'black', relief = 'groove')
        _style.configure('Frame_grey2.TFrame', borderwidth =2, background = 'black', relief = 'groove')
        _style.configure('Frame_filter.TFrame', borderwidth = 2, bordercolor = 'grey', background = 'black', relief = 'groove')
        _style.configure('Green.TCombobox', background = 'black', foreground = 'forest green', font = 'Verdana 10')
        _style.configure('Red.TCombobox', background = 'black', foreground = 'orange red', font = 'Verdana 10')
        #print(self.style.layout('Frame_grey2.TFrame'))
        #print(self.style.element_options('Frame.border'))
        return _style


    def _get_top_frame(self) -> ttk.Frame:
        '''
        the top frame contains: 
            - the headline
            - the path to the database (green font when working on the productive db, otherwise red font)
            - the button to change the path to the database
        '''
        _top_frame = ttk.Frame(self.window, style = 'Frame_grey2.TFrame')
        self.label_top = ttk.Label(_top_frame, 
                                   style = 'Green2.TLabel', 
                                   text = 'Lagerverwaltung Comline',
                                   anchor = 'e')
        self.label_top_path_to_db = ttk.Label(_top_frame, 
                                   style = 'Red.TLabel', 
                                   text = f'Pfad zur Datenbank: {self.path_to_db}',
                                   anchor = 'w')
        
        if 'Service-Center' in self.path_to_db:
            color = 'forest green'
        else:
            color = 'orange red'
        self.label_top_path_to_db.configure(foreground = color)
        
        self.button_change_db = ttk.Button(_top_frame, 
                                           style = 'Green.TButton',
                                           text = 'Pfad ändern',
                                           command = lambda : fc.change_db_path(self))
        
        self.button_change_db.pack(side = tk.RIGHT, padx = 50, pady = 5)
        self.label_top.pack(side = tk.LEFT,padx = 100, pady = 10, fill = 'x', expand = True)
        self.label_top_path_to_db.pack(side = tk.RIGHT, padx = 10, pady = 10, fill = 'x', expand = True)
        return _top_frame

        
    def _get_button_frame(self) -> ttk.Frame:    
        '''
        creates the frames and buttons at the top side of the app
        the left frame contains buttons for all users (STANDARD, EXPERT, ADMIN) (db read only)
        the middle frame contains buttons for EXPERT and ADMIN (db read and write)
        the right frame contains buttons only for ADMIN (db read, write and delete)
        '''
        _main_button_frame = ttk.Frame(self.window, style = 'Frame_grey.TFrame')
        _left_button_frame = self._get_left_button_frame(_main_button_frame)
        _right_button_frame = self._get_right_button_frame(_main_button_frame)
        _middle_button_frame = self._get_middle_button_frame(_main_button_frame)
        
        _left_button_frame.pack(side = tk.LEFT, padx = 5, pady = 5)
        if self.user == ADMIN:
            _right_button_frame.pack(side = tk.RIGHT, padx = 5, pady = 5)
        if self.user == ADMIN:
            _middle_button_frame.pack(side = tk.TOP, padx = 5, pady = 5)
        elif self.user == EXPERT:
            _middle_button_frame.pack(side = tk.LEFT, padx = 5, pady = 5)
        
        return _main_button_frame
    
    def _get_left_button_frame(self, _main_button_frame:ttk.Frame) -> ttk.Frame:
        _left_button_frame = ttk.Frame(_main_button_frame, style = 'Frame_grey.TFrame')
        self.button_krit_mat = ttk.Button(_left_button_frame, 
                                    style = 'Green.TButton', 
                                    text = 'Kritisches Material \nanzeigen', 
                                    command = lambda: fc.show_critical_material(self)
                                    )
        self.button_bestand = ttk.Button(_left_button_frame, 
                                    style = 'Green.TButton', 
                                    text = 'Bestand\nanzeigen', 
                                    command = lambda:fc.show_stock(self)
                                    )
        self.button_wareneingang = ttk.Button(_left_button_frame, 
                                              style = 'Green.TButton', 
                                              text = 'Wareneingang\nanzeigen',
                                              command = lambda:fc.show_ingoing_material(self)
                                              )
        self.button_warenausgang = ttk.Button(_left_button_frame, 
                                              style = 'Green.TButton', 
                                              text = 'Warenausgang\nanzeigen',
                                              command = lambda:fc.show_outgoing_material(self)
                                              )
        self.button_sm_auftrag = ttk.Button(_left_button_frame,
                                            style = 'Green.TButton',
                                            text = 'Material für\nSM-Auftrag anzeigen',
                                            command = lambda:fc.show_material_for_order(self)
                                            )
        
        self.button_krit_mat.pack(side = tk.LEFT, padx = 5, pady = 5)
        self.button_bestand.pack(side = tk.LEFT, padx = 5, pady = 5)
        self.button_wareneingang.pack(side = tk.LEFT, padx = 5, pady = 5)
        self.button_warenausgang.pack(side = tk.LEFT, padx = 5, pady = 5)
        self.button_sm_auftrag.pack(side = tk.LEFT, padx = 5, pady = 5)
        return _left_button_frame
        

    def _get_right_button_frame(self, _main_button_frame:ttk.Frame) -> ttk.Frame:
        # ADMIN Knopf und Combobox
        _right_button_frame = ttk.Frame(_main_button_frame, style = 'Frame_grey.TFrame')
        self.combobox_loeschen = ttk.Combobox(_right_button_frame, 
                                                style = 'Red.TCombobox',
                                                width = 50, 
                                                values = ['Kleinstmaterial', 'Standardmaterial', 'Warenausgabe_Comline',
                                                        'Warenausgang', 'Wareneingang'])
        self.combobox_loeschen.current(0)
        self.button_loeschen = ttk.Button(_right_button_frame, 
                                            style = 'Red.TButton', 
                                            text = 'Einträge aus \nDatenbank löschen',
                                            command  = lambda: fc.filter_entries_to_delete(self))
        self.button_loeschen.pack(side = tk.LEFT, padx = 5, pady = 5)
        self.combobox_loeschen.pack(side = tk.LEFT, padx = 5, pady = 5)
        return _right_button_frame
        
        
    def _get_middle_button_frame(self, _main_button_frame:ttk.Frame) -> ttk.Frame:
        _middle_button_frame = ttk.Frame(_main_button_frame, style = 'Frame_grey.TFrame')
        self.button_wareneingang_buchen = ttk.Button(_middle_button_frame, 
                                                        style = 'Blue.TButton', 
                                                        text = 'Wareneingang buchen\n(Anlieferung CTDI)',
                                                        command = lambda:fc.book_ingoing_position(self))
        self.button_warenausgang_buchen = ttk.Button(_middle_button_frame, 
                                                        style = 'Blue.TButton', 
                                                        text = 'Warenausgang buchen\n(Excel aus PSL)',
                                                        command = lambda: fc.book_outgoing_from_excel_file(self))
        self.button_warenausgang_buchen_Kleinstmaterial = ttk.Button(_middle_button_frame, 
                                                        style = 'Blue.TButton', 
                                                        text = 'Warenausgang buchen\n(nur Kleinstmaterial)',
                                                        command = lambda: fc.book_outgoing_kleinstmaterial(self))
        self.button_wareneingang_buchen.pack(side = tk.LEFT, padx = 5, pady = 5)
        self.button_warenausgang_buchen.pack(side = tk.LEFT, padx = 5, pady = 5)
        self.button_warenausgang_buchen_Kleinstmaterial.pack(side = tk.LEFT, padx = 5, pady = 5)
        return _middle_button_frame

        

    def _get_filter_frame(self) -> ttk.Frame:
        '''
        the filter frame contains:
            - entry for materialnumber
            - entry for work order number
            - entry for position / id number (only in ADMIN mode)
        '''
        _filter_frame = ttk.Frame(self.window, style = 'Frame_filter.TFrame')
        
        # these filters are always present
        self.matnr_label = ttk.Label(_filter_frame, style = 'Green.TLabel', text = ' Daten filtern:     Materialnummer:')
        self.matnr_entry = ttk.Entry(_filter_frame, style = 'Green.TEntry')
        self.sm_label = ttk.Label(_filter_frame, style = 'Green.TLabel', text = 'SM Nummer:')
        self.sm_entry = ttk.Entry(_filter_frame, style = 'Green.TEntry')
        
        self.matnr_label.pack(side = tk.LEFT, padx = 5, pady = 5)
        self.matnr_entry.pack(side = tk.LEFT, padx = 5, pady = 5)
        self.sm_label.pack(side = tk.LEFT, padx = 5, pady = 5)
        self.sm_entry.pack(side = tk.LEFT, padx = 5, pady = 5)
        
        # this button and label is only shown in ADMIN mode
        self.posnr_entry = ttk.Entry(_filter_frame, style = 'Green.TEntry')
        if self.user == ADMIN:
            self.posnr_label = ttk.Label(_filter_frame, style = 'Green.TLabel', text = 'Position / ID:')
            self.posnr_label.pack(side = tk.LEFT, padx = 5, pady = 10)
            self.posnr_entry.pack(side = tk.LEFT, padx = 5, pady = 10)
        
        return _filter_frame
        

    def _get_bottom_frame(self) -> ttk.Frame:
        '''
        bottom frame contains:
            - button to toggle order status
            - button to print on standard printer
            - button to delete entries from database
            - textbox to show the filtered data from the underlying sqlite3 database
        '''
        
        _bottom_frame = ttk.Frame(self.window, style = 'Frame_grey.TFrame', padding = 10)
        # creates buttons and label at the bottom of the bottom frame
        self.entry_frame = ttk.Frame(self.window, style = 'Frame_grey.TFrame')
        self.bestellt_label = ttk.Label(self.entry_frame, 
                                        style = 'Blue.TLabel', 
                                        text = 'Materialnummer markieren um den Bestellstatus zu wechseln'
                                        )
        self.bestellt_button = ttk.Button(self.entry_frame, 
                                            style = 'Blue2.TButton', 
                                            text = 'Bestellstatus ändern',
                                            command = lambda:fc.toggle_ordered_status(self)
                                            )
        self.print_button = ttk.Button(self.entry_frame, 
                                            style = 'Green2.TButton', 
                                            text = 'Drucken',
                                            command = lambda:fc.print_screen(self)
                                            )
        self.delete_button = ttk.Button(self.entry_frame, 
                                            style = 'Red_Delete.TButton', 
                                            text = 'angezeigte Daten loeschen',
                                            command = lambda:fc.delete_selected_entries(self)
                                            )
        
        if self.user in (ADMIN, EXPERT):
            self.bestellt_label.pack(side = tk.LEFT, padx = 50, pady = 5)
            self.bestellt_button.pack(side = tk.LEFT, padx = 50, pady = 5)
        if self.user == STANDARD:
               self.print_button.pack(side = tk.TOP, pady = 5)
        else:
               self.print_button.pack(side = tk.LEFT, pady = 5)
        if self.user == ADMIN:    
            self.delete_button.pack(side = tk.LEFT, pady = 5)
            
        self.entry_frame.pack(side = tk.BOTTOM, fill = 'x')
        
        #creates a textbox that fills the rest of the bottom fame
        self.vertical_scrollbar = ttk.Scrollbar(_bottom_frame, orient = tk.VERTICAL)
        self.horizontal_scrollbar = ttk.Scrollbar(_bottom_frame, orient = tk.HORIZONTAL)
        self.output_box = tk.Text(_bottom_frame, 
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
        return _bottom_frame
  

    def open_booking_window(self, title, selection):
        '''
        creates the frontend when booking ingoing or outgoing material

        conntains:
            - a simple label
            - combobox for material selection
            - entry for amount 
            - button to book the material
        '''
        self.user_closed_window = True
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
        self.title_label = ttk.Label(self.ueberschrift_frame,
                                            style = 'Green.TLabel',
                                            anchor = tk.CENTER,
                                            text = f'\n{title}',
                                            font = 'Verdana 12 bold',
                                            foreground = 'orange red'
                                            )
        self.ueberschrift_label = ttk.Label(self.ueberschrift_frame,
                                            style = 'Green.TLabel',
                                            anchor = tk.CENTER,
                                            text = '\nBitte Material aus Dropdown Menü wählen und Menge eingeben.\n'
                                            )
        self.mat_frame = ttk.Frame(self.ingoing_window)
        self.mat_label = ttk.Label(self.mat_frame,
                                       style = 'Green.TLabel',
                                       anchor = tk.CENTER,
                                       text = 'Material')
        self.title_label.pack(side = tk.TOP, fill = 'x')
        self.ueberschrift_label.pack(side = tk.TOP, fill = 'x')

        self.matnr_combobox = AutocompleteCombobox(self.mat_frame,
                                           style = 'Green.TCombobox',
                                           width = 60,
                                           textvariable=self.ingoing_mat_string_var)
        self.matnr_combobox.set_completion_list(values)
        self.mat_label.pack(side = tk.TOP, fill = 'x')
        self.matnr_combobox.pack(side = tk.TOP, fill = 'x')
        self.menge_frame = ttk.Frame(self.ingoing_window)
        self.stck_label = ttk.Label(self.menge_frame,
                                       style = 'Green.TLabel',
                                       anchor = tk.CENTER,
                                       text = 'Menge')
        self.stck_entry = ttk.Entry(self.menge_frame,
                                    style = 'Green.TEntry',
                                    textvariable=self.ingoing_menge_var)
        self.ok_button = ttk.Button(self.ingoing_window,
                                    style = 'Blue.TButton',
                                    text = 'buchen',
                                    command = lambda : fc.confirm_user_input(self))
        self.stck_label.pack(side = tk.TOP, fill = 'x')
        self.stck_entry.pack(side = tk.TOP, fill = 'x')

        self.stck_entry.bind('<Return>', lambda _ : fc.confirm_user_input(self))

        self.ueberschrift_frame.pack(side = tk.TOP, padx = 10, pady = 5, fill = 'x')
        self.ok_button.pack(side = tk.BOTTOM, fill = 'x', padx = 10, pady = 10)
        self.mat_frame.pack(side = tk.LEFT, padx = 10)    
        self.menge_frame.pack(side = tk.LEFT, padx = 10)
        self.matnr_combobox.focus()
        self.ingoing_window.mainloop()

    

class AutocompleteCombobox(tk.ttk.Combobox):
        '''
        stolen from 
        sloth (https://stackoverflow.com/users/142637/sloth)
        at 
        https://stackoverflow.com/questions/12298159/tkinter-how-to-create-a-combo-box-with-autocompletion

        thx for this :)

        creates a combobox with autocompletition
        '''
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
        