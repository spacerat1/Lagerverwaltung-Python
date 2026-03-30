import sqlite3
import functions as fc
from collections import defaultdict

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QFrame,
    QVBoxLayout, QHBoxLayout, 
    QPushButton, QLabel, QLineEdit, QComboBox,
    QTreeWidget, QTreeWidgetItem,
    QDialog, QAbstractItemView,
    QHeaderView, QCompleter
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QFont, QBrush

ADMIN = 'admin'
EXPERT = 'expert'
STANDARD = 'standard'

# ── Color constants ──────────────────────────────────────────────────
BLACK        = '#000000'
FOREST_GREEN = '#228B22'
ORANGE_RED   = '#FF4500'
DEEP_SKY     = '#00BFFF'
DARK_GREEN   = '#006400'
LIGHT_GREEN  = '#90EE90'
GREY65       = '#A6A6A6'
GREY80       = '#CCCCCC'
GREY10       = '#1A1A1A'
WHITE        = '#FFFFFF'

FRAME_COLOR = 'dimgrey'

# ── Stylesheet helpers ───────────────────────────────────────────────
def _btn(fg: str, bg: str = BLACK, disabled_fg: str | None = None,
         disabled_bg: str | None = None) -> str:
    dfg = disabled_fg or fg
    dbg = disabled_bg or bg
    return (
        f"QPushButton{{color:{fg};background:{bg};font:10pt Verdana;"
        f"border:1px solid {fg};padding:4px 8px;}}"
        f"QPushButton:hover{{background:{fg};color:{BLACK};}}"
        f"QPushButton:pressed{{background:{fg};color:{BLACK};border-style:inset;}}"
        f"QPushButton:disabled{{color:{dfg};background:{dbg};"
        f"border:2px inset turquoise;}}"
    )

GREEN_BTN  = _btn(FOREST_GREEN, BLACK, BLACK, FOREST_GREEN)
BLUE_BTN   = _btn(DEEP_SKY)
RED_BTN    = _btn(ORANGE_RED)

GREEN_LABEL_SS  = f"color:{FOREST_GREEN};background:{BLACK};font:10pt Verdana; border: 1px solid green"
GREEN2_LABEL_SS = f"color:{FOREST_GREEN};background:{BLACK};font:bold 10pt Verdana; border: 1px solid green"
RED_LABEL_SS    = f"color:{ORANGE_RED};background:{BLACK};font:10pt Verdana;"
BLUE_LABEL_SS   = f"color:{DEEP_SKY};background:{BLACK};font:10pt Verdana; border : 0px"

GREEN_ENTRY_SS  = (f"QLineEdit{{color:{BLACK};background:{WHITE};"
                   f"font:bold 10pt Verdana;border:1px solid {FOREST_GREEN};}}"
                   f"QLineEdit:disabled{{background:{GREY10};}}")

GREEN_COMBO_SS  = (f"QComboBox{{color:{FOREST_GREEN};background:{BLACK};"
                   f"font:10pt Verdana;border:1px solid {FOREST_GREEN};}}"
                   f"QComboBox QAbstractItemView{{color:{FOREST_GREEN};"
                   f"background:{BLACK};}}")

RED_COMBO_SS    = (f"QComboBox{{color:{ORANGE_RED};background:{BLACK};"
                   f"font:10pt Verdana;border:1px solid {ORANGE_RED};}}"
                   f"QComboBox QAbstractItemView{{color:{ORANGE_RED};"
                   f"background:{BLACK};}}")

TREE_SS = (
    f"QTreeWidget{{color:{FOREST_GREEN};background:{BLACK};"
    f"font:11pt Verdana;border:2px solid {FRAME_COLOR};" #WHITE
    f"alternate-background-color:{BLACK};}}"
    f"QTreeWidget::item:selected{{background:{LIGHT_GREEN};color:{BLACK};}}"
    f"QHeaderView::section{{background:{DARK_GREEN};color:{WHITE};"
    f"font:bold 8pt Verdana;border:1px solid {BLACK};}}"
)

FRAME_SS  = f"background:{BLACK};border:2px groove {FRAME_COLOR};"
FRAME2_SS = f"background:{BLACK};border:2px groove {FRAME_COLOR};"


# ── InputBox dialog ──────────────────────────────────────────────────
class InputBoxDialog(QDialog):
    """Replaces the Tkinter Toplevel InputBox / DialogResult pattern."""

    def __init__(self, parent, mat_number: str, mat_name: str):
        super().__init__(parent)
        self.setWindowTitle('Bestellmenge')
        self.setModal(True)
        self.result_text = ''

        self.setStyleSheet(f"background:{BLACK};")
        layout = QVBoxLayout(self)

        mat_label = QLabel(f"{mat_number} {mat_name}", self)
        mat_label.setStyleSheet(GREEN_LABEL_SS)
        layout.addWidget(mat_label)

        qty_label = QLabel('bestellte Menge?', self)
        qty_label.setStyleSheet(GREEN_LABEL_SS)
        layout.addWidget(qty_label)

        self.entry = QLineEdit(self)
        self.entry.setStyleSheet(GREEN_ENTRY_SS)
        layout.addWidget(self.entry)

        ok_btn = QPushButton('OK', self)
        ok_btn.setStyleSheet(BLUE_BTN)
        ok_btn.clicked.connect(self._accept)
        self.entry.returnPressed.connect(self._accept)
        layout.addWidget(ok_btn)

        self.entry.setFocus()

    def _accept(self):
        self.result_text = self.entry.text()
        self.accept()


# ── Main App window ──────────────────────────────────────────────────
class App(QMainWindow):
    '''
    PyQt6 port of the Tkinter warehouse management front-end.

    top_buttons:
        - STANDARD (all users)-
        Pfad ändern: changes the path to the underlying sqlite3 database
        Kritisches Material anzeigen: shows material below a threshold
        Bestand anzeigen: shows the stock
        Wareneingang anzeigen: shows ingoing material
        Warenausgang anzeigen: shows outgoing material
        Material für SM-Auftrag anzeigen: shows all materials for a work order
        - EXPERT -
        Wareneingang buchen (Anlieferung CTDI): book incoming material
        Warenausgang buchen (Excel aus PSL): book outgoing material via Excel
        Warenausgang buchen (nur Kleinstmaterial): book outgoing (no order)
        - ADMIN -
        Einträge aus Datenbank löschen: delete entries from the database
        Combobox: choose which table to delete from
    filters:
        Materialnummer / SM Nummer / Position / ID
    bottom_buttons:
        Drucken / Bestellstatus ändern / angezeigte Daten löschen
    '''

    def __init__(self, connection: sqlite3.Connection,
                 cursor: sqlite3.Cursor, user: str, path_to_db: str):
        super().__init__()
        self.user = user
        self.connection = connection
        self.cursor = cursor
        self.path_to_db = path_to_db

        self.strDialogResult = ''
        self.execution_string = ''
        self.execution_tuple = ''
        self.user_closed_window = False

        self.setWindowTitle('Lagerverwaltung Comline')
        self.setMinimumSize(1200, 768)
        self.setStyleSheet(f"background:{BLACK};")

        # context menu (right-click) – connect to fc helper
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(
            lambda pos: fc.show_context_menu(pos, self))

        self._init_app()

    # ── Initialisation ───────────────────────────────────────────────
    def _init_app(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QVBoxLayout(central)
        root_layout.setSpacing(0)
        root_layout.setContentsMargins(0, 0, 0, 0)

        self.top_frame    = self._get_top_frame()
        self.button_frame = self._get_button_frame()
        self.filter_frame = self._get_filter_frame()
        self.bottom_frame = self._get_bottom_frame()

        root_layout.addWidget(self.top_frame)
        root_layout.addWidget(self.button_frame)
        root_layout.addWidget(self.filter_frame)
        root_layout.addWidget(self.bottom_frame, stretch=1)

        self._create_standard_values()

        self.disabled_button = self.button_krit_mat
        fc.show_critical_material(self)

    def _create_standard_values(self) -> None:
        # Row-tag colours for the QTreeWidget are applied via item data roles
        # (see _set_row_tag / tag_configure wrappers below).
        self.tag_styles: dict[str, dict] = {
            'green':      {'bg': QColor('forest green'), 'fg': QColor('white'),
                           'font': QFont('Verdana', 11, QFont.Weight.Bold)},
            'bundle_head':{'bg': QColor('#A6A6A6'),       'fg': QColor('black'),
                           'font': QFont('Verdana', 11, QFont.Weight.Bold)},
            'bundle':     {'bg': QColor('#CCCCCC'),       'fg': QColor('black'),
                           'font': QFont('Verdana', 11)},
            'red_font':   {'bg': QColor('black'),         'fg': QColor('orange red'),
                           'font': QFont('Verdana', 11)},
            'green_font': {'bg': QColor('black'),         'fg': QColor('forest green'),
                           'font': QFont('Verdana', 11)},
        }

        self.threshhold_dict        = defaultdict(int)
        self.recommended_amount_dict= defaultdict(int)
        self.materialnames_dict     = defaultdict(str)
        self.units_dict             = defaultdict(str)
        self.deprecated_dict        = defaultdict(str)

        materials = self.cursor.execute(
            "SELECT * FROM Standardmaterial UNION SELECT * FROM Kleinstmaterial"
        ).fetchall()
        deprecated_materials = self.cursor.execute(
            "SELECT * FROM Veraltetes_Material"
        ).fetchall()

        self.correction_dict = defaultdict(int)
        for row in self.cursor.execute("SELECT * FROM Jahresinventur_Korrekturdaten"):
            self.correction_dict[row['MatNr']] += row['Menge']

        for row in materials:
            self.threshhold_dict[row['MatNr']]         = row['Grenzwert']
            self.recommended_amount_dict[row['MatNr']] = row['Auffüllen']
            self.materialnames_dict[row['MatNr']]      = row['Bezeichnung']
            self.units_dict[row['MatNr']]              = row['Einheit']
        for row in deprecated_materials:
            self.deprecated_dict[row['MatNr']] = row['Bezeichnung']

        standard_materials = self.cursor.execute(
            "SELECT * FROM Standardmaterial"
        ).fetchall()
        self.standard_materials = [m['MatNr'] for m in standard_materials]
        small_materials = self.cursor.execute(
            "SELECT * FROM Kleinstmaterial"
        ).fetchall()
        self.small_materials = [m['MatNr'] for m in small_materials]

        self.columns_dict = {
            'Adresse':                          (800, Qt.AlignmentFlag.AlignLeft),
            'Auffüllen':                        (80,  Qt.AlignmentFlag.AlignCenter),
            'Bedarfsmenge':                     (110, Qt.AlignmentFlag.AlignCenter),
            'Bemerkungen':                      (1000,Qt.AlignmentFlag.AlignLeft),
            'Bestand':                          (80,  Qt.AlignmentFlag.AlignRight),
            'bestellt':                         (80,  Qt.AlignmentFlag.AlignCenter),
            'Bezeichnung':                      (400, Qt.AlignmentFlag.AlignLeft),
            'Datum':                            (200, Qt.AlignmentFlag.AlignCenter),
            'Einheit':                          (60,  Qt.AlignmentFlag.AlignLeft),
            'empfohlene Menge':                 (130, Qt.AlignmentFlag.AlignCenter),
            'Grenzwert':                        (80,  Qt.AlignmentFlag.AlignCenter),
            'ID':                               (80,  Qt.AlignmentFlag.AlignRight),
            'LAST_COLUMN':                      (50,  Qt.AlignmentFlag.AlignCenter),
            'Lieferschein':                     (150, Qt.AlignmentFlag.AlignCenter),
            'Materialbeleg':                    (150, Qt.AlignmentFlag.AlignCenter),
            'MatNr':                            (110, Qt.AlignmentFlag.AlignCenter),
            'MatNr.':                           (110, Qt.AlignmentFlag.AlignCenter),
            'Menge':                            (80,  Qt.AlignmentFlag.AlignRight),
            'Menge ':                           (80,  Qt.AlignmentFlag.AlignCenter),
            'Nr.':                              (80,  Qt.AlignmentFlag.AlignRight),
            'Position':                         (80,  Qt.AlignmentFlag.AlignCenter),
            'Pos.Typ':                          (60,  Qt.AlignmentFlag.AlignCenter),
            'PosTyp':                           (60,  Qt.AlignmentFlag.AlignCenter),
            'SD Beleg':                         (150, Qt.AlignmentFlag.AlignCenter),
            'SD_Beleg':                         (150, Qt.AlignmentFlag.AlignCenter),
            'SM Nummer':                        (110, Qt.AlignmentFlag.AlignCenter),
            'SM_Nummer':                        (110, Qt.AlignmentFlag.AlignCenter),
            'SM Nummer / ID':                   (120, Qt.AlignmentFlag.AlignCenter),
            'Umbuchungsmenge':                  (140, Qt.AlignmentFlag.AlignCenter),
            'VPSZ':                             (200, Qt.AlignmentFlag.AlignRight),
            'Warenausgangsmenge':               (160, Qt.AlignmentFlag.AlignCenter),
        }

        self.execution_dict = {
            'Kleinstmaterial':
                ('SELECT * FROM Kleinstmaterial WHERE MatNr LIKE ?', r'%matnr%,'),
            'Standardmaterial':
                ('SELECT * FROM Standardmaterial WHERE Matnr LIKE ?', r'%matnr%,'),
            'Warenausgabe_Comline':
                ('SELECT * FROM Warenausgabe_Comline WHERE SM_Nummer LIKE ? AND MatNr LIKE ?', r'%sm%,%matnr%'),
            'Warenausgang':
                ('SELECT * FROM Warenausgang WHERE SM_Nummer LIKE ? AND MatNr LIKE ?', r'%sm%,%matnr%'),
            'Wareneingang':
                ('SELECT * FROM Wareneingang WHERE ID LIKE ?', r'%posnr%,'),
            'Warenausgang_Kleinstmaterial_ohne_SM_Bezug':
                ('SELECT * FROM Warenausgang_Kleinstmaterial_ohne_SM_Bezug WHERE ID LIKE ?', r'%posnr%,'),
            'Adresszuordnung':
                ('SELECT * FROM Adresszuordnung WHERE SM_Nummer LIKE ?', r'%sm%,'),
            'Veraltetes_Material':
                ('SELECT * FROM Veraltetes_Material WHERE MatNr LIKE ?', r'%matnr%,'),
            'Jahresinventur_Korrekturdaten':
                ('SELECT * FROM Jahresinventur_Korrekturdaten WHERE MatNr LIKE ?', r'%matnr%,'),
        }

        self.filter_dict = {
            'Kleinstmaterial':
                ([self.matnr_entry], [self.sm_entry, self.posnr_entry]),
            'Standardmaterial':
                ([self.matnr_entry], [self.sm_entry, self.posnr_entry]),
            'Warenausgabe_Comline':
                ([self.matnr_entry, self.sm_entry], [self.posnr_entry]),
            'Warenausgang':
                ([self.matnr_entry, self.sm_entry], [self.posnr_entry]),
            'Wareneingang':
                ([self.posnr_entry], [self.sm_entry, self.matnr_entry]),
            'Warenausgang_Kleinstmaterial_ohne_SM_Bezug':
                ([self.posnr_entry], [self.sm_entry, self.matnr_entry]),
            'Adresszuordnung':
                ([self.sm_entry], [self.posnr_entry, self.matnr_entry]),
            'Veraltetes_Material':
                ([self.matnr_entry], [self.sm_entry, self.posnr_entry]),
            'Jahresinventur_Korrekturdaten':
                ([self.matnr_entry], [self.sm_entry, self.posnr_entry]),
        }

        self.deletion_dict = {
            'Kleinstmaterial':
                ('DELETE FROM Kleinstmaterial WHERE MatNr = ?', 'matnr,'),
            'Standardmaterial':
                ('DELETE FROM Standardmaterial WHERE Matnr = ?', 'matnr,'),
            'Warenausgabe_Comline':
                ('DELETE FROM Warenausgabe_Comline WHERE SM_Nummer = ? AND Position = ?', 'sm,posnr'),
            'Warenausgang':
                ('DELETE FROM Warenausgang WHERE SM_Nummer = ? AND Position = ?', 'sm,posnr'),
            'Wareneingang':
                ('DELETE FROM Wareneingang WHERE ID = ?', 'posnr,'),
            'Warenausgang_Kleinstmaterial_ohne_SM_Bezug':
                ('DELETE FROM Warenausgang_Kleinstmaterial_ohne_SM_Bezug WHERE ID = ?', 'posnr,'),
            'Adresszuordnung':
                ('DELETE FROM Adresszuordnung WHERE SM_Nummer = ?', 'sm,'),
            'Veraltetes_Material':
                ('DELETE FROM Veraltetes_Material WHERE ID = ?', 'posnr,'),
            'Jahresinventur_Korrekturdaten':
                ('DELETE FROM Jahresinventur_Korrekturdaten WHERE ID = ?', 'posnr,'),
        }

    # ── Tag helpers (replaces ttk tag_configure / item tags) ─────────
    # def tag_configure(self, tag: str, **kwargs):
    #     """No-op – styles are pre-loaded in tag_styles dict."""
    #     pass  # styles already defined in _create_standard_values

    def apply_tag(self, item: QTreeWidgetItem, tag: str) -> None:
        """Apply a named tag style to a QTreeWidgetItem (all columns)."""
        style = self.tag_styles.get(tag)
        if not style:
            return
        col_count = self.output_listbox.columnCount()
        for col in range(col_count):
            if 'bg' in style:
                item.setBackground(col, QBrush(style['bg']))
            if 'fg' in style:
                item.setForeground(col, QBrush(style['fg']))
            if 'font' in style:
                item.setFont(col, style['font'])

    # ── Top frame ────────────────────────────────────────────────────
    def _get_top_frame(self) -> QFrame:
        frame = QFrame()
        frame.setStyleSheet(FRAME2_SS)
        layout = QHBoxLayout(frame)

        self.label_top = QLabel('Lagerverwaltung Comline')
        self.label_top.setStyleSheet(GREEN2_LABEL_SS)
        self.label_top.setAlignment(Qt.AlignmentFlag.AlignCenter)

        color = FOREST_GREEN if 'Service-Center' in self.path_to_db else ORANGE_RED
        self.label_top_path_to_db = QLabel(f'Pfad zur Datenbank: {self.path_to_db}')
        self.label_top_path_to_db.setStyleSheet(
            f"color:{color};background:{BLACK};font:10pt Verdana; border: 1px solid {color}")
        self.label_top_path_to_db.setAlignment(Qt.AlignmentFlag.AlignLeft |
                                                Qt.AlignmentFlag.AlignVCenter)

        self.button_change_db = QPushButton('Pfad ändern')
        self.button_change_db.setStyleSheet(GREEN_BTN)
        self.button_change_db.clicked.connect(lambda: fc.change_db_path(self))

        layout.addWidget(self.label_top, stretch=1)
        layout.addWidget(self.label_top_path_to_db, stretch=2)
        layout.addWidget(self.button_change_db)
        layout.setContentsMargins(10, 5, 50, 5)
        return frame

    # ── Button frame ─────────────────────────────────────────────────
    def _get_button_frame(self) -> QFrame:
        frame = QFrame()
        frame.setStyleSheet(FRAME_SS)
        layout = QHBoxLayout(frame)
        layout.setSpacing(5)
        layout.setContentsMargins(5, 5, 5, 5)

        left   = self._get_left_button_frame()
        middle = self._get_middle_button_frame()
        right  = self._get_right_button_frame()  # immer aufrufen, damit Widgets existieren

        layout.addWidget(left)

        if self.user in (EXPERT, ADMIN):
            layout.addWidget(middle)
        if self.user == ADMIN:
            layout.addWidget(right)

        layout.addStretch()
        return frame

    def _get_left_button_frame(self) -> QFrame:
        frame = QFrame()
        frame.setStyleSheet(FRAME_SS)
        layout = QHBoxLayout(frame)
        layout.setSpacing(5)

        self.button_krit_mat = QPushButton('Kritisches Material\nanzeigen')
        self.button_krit_mat.setStyleSheet(GREEN_BTN)
        self.button_krit_mat.clicked.connect(lambda: fc.show_critical_material(self))

        self.button_bestand = QPushButton('Bestand\nanzeigen')
        self.button_bestand.setStyleSheet(GREEN_BTN)
        self.button_bestand.clicked.connect(lambda: fc.show_stock(self))

        self.button_wareneingang = QPushButton('Wareneingang\nanzeigen')
        self.button_wareneingang.setStyleSheet(GREEN_BTN)
        self.button_wareneingang.clicked.connect(
            lambda: fc.show_ingoing_material(self))

        self.button_warenausgang = QPushButton('Warenausgang\nanzeigen')
        self.button_warenausgang.setStyleSheet(GREEN_BTN)
        self.button_warenausgang.clicked.connect(
            lambda: fc.show_outgoing_material(self))

        self.button_sm_auftrag = QPushButton('Material für\nSM-Auftrag anzeigen')
        self.button_sm_auftrag.setStyleSheet(GREEN_BTN)
        self.button_sm_auftrag.clicked.connect(
            lambda: fc.show_material_for_order(self))

        for btn in (self.button_krit_mat, self.button_bestand,
                    self.button_wareneingang, self.button_warenausgang,
                    self.button_sm_auftrag):
            layout.addWidget(btn)

        return frame

    def _get_middle_button_frame(self) -> QFrame:
        frame = QFrame()
        frame.setStyleSheet(FRAME_SS)
        layout = QHBoxLayout(frame)
        layout.setSpacing(5)

        self.button_wareneingang_buchen = QPushButton(
            'Wareneingang buchen\n(Anlieferung CTDI)')
        self.button_wareneingang_buchen.setStyleSheet(BLUE_BTN)
        self.button_wareneingang_buchen.clicked.connect(
            lambda: fc.book_ingoing_position(self))

        self.button_warenausgang_buchen = QPushButton(
            'Warenausgang buchen\n(Excel aus PSL)')
        self.button_warenausgang_buchen.setStyleSheet(BLUE_BTN)
        self.button_warenausgang_buchen.clicked.connect(
            lambda: fc.book_outgoing_from_excel_file(self))

        self.button_warenausgang_buchen_Kleinstmaterial = QPushButton(
            'Warenausgang buchen\n(nur Kleinstmaterial)')
        self.button_warenausgang_buchen_Kleinstmaterial.setStyleSheet(BLUE_BTN)
        self.button_warenausgang_buchen_Kleinstmaterial.clicked.connect(
            lambda: fc.book_outgoing_kleinstmaterial(self))

        for btn in (self.button_wareneingang_buchen,
                    self.button_warenausgang_buchen,
                    self.button_warenausgang_buchen_Kleinstmaterial):
            layout.addWidget(btn)

        return frame

    def _get_right_button_frame(self) -> QFrame:
        frame = QFrame()
        frame.setStyleSheet(FRAME_SS)
        layout = QHBoxLayout(frame)
        layout.setSpacing(5)

        tables = self.cursor.execute(
            "SELECT name FROM sqlite_master WHERE type='table';"
        ).fetchall()
        values = [t[0] for t in tables if t[0] != 'sqlite_sequence']

        # Immer erstellen (auch für STANDARD/EXPERT), damit filter_dict
        # und set_widget_status nie auf ein gelöschtes Objekt stoßen.
        # Nur für ADMIN wird das Widget ins Layout eingefügt und sichtbar.
        self.combobox_loeschen = QComboBox()
        self.combobox_loeschen.setStyleSheet(RED_COMBO_SS)
        self.combobox_loeschen.setMinimumWidth(300)
        self.combobox_loeschen.addItems(values)
        self.combobox_loeschen.setCurrentIndex(0)

        self.button_loeschen = QPushButton('Einträge aus\nDatenbank löschen')
        self.button_loeschen.setStyleSheet(RED_BTN)
        self.button_loeschen.clicked.connect(
            lambda: fc.filter_entries_to_delete(self))

        if self.user == ADMIN:
            layout.addWidget(self.button_loeschen)
            layout.addWidget(self.combobox_loeschen)

        return frame

    # ── Filter frame ─────────────────────────────────────────────────
    def _get_filter_frame(self) -> QFrame:
        frame = QFrame()
        frame.setStyleSheet(FRAME_SS)
        layout = QHBoxLayout(frame)
        layout.setSpacing(5)
        layout.setContentsMargins(5, 5, 5, 5)

        matnr_label = QLabel(' Daten filtern:     Materialnummer:')
        matnr_label.setStyleSheet(GREEN_LABEL_SS)
        self.matnr_entry = QLineEdit()
        self.matnr_entry.setStyleSheet(GREEN_ENTRY_SS)

        sm_label = QLabel('SM Nummer:')
        sm_label.setStyleSheet(GREEN_LABEL_SS)
        self.sm_entry = QLineEdit()
        self.sm_entry.setStyleSheet(GREEN_ENTRY_SS)

        # posnr_entry always created; only shown for ADMIN
        self.posnr_entry = QLineEdit()
        self.posnr_entry.setStyleSheet(GREEN_ENTRY_SS)

        layout.addWidget(matnr_label)
        layout.addWidget(self.matnr_entry)
        layout.addWidget(sm_label)
        layout.addWidget(self.sm_entry)

        if self.user == ADMIN:
            posnr_label = QLabel('Position / ID:')
            posnr_label.setStyleSheet(GREEN_LABEL_SS)
            layout.addWidget(posnr_label)
            layout.addWidget(self.posnr_entry)

        layout.addStretch()
        return frame

    # ── Bottom frame ─────────────────────────────────────────────────
    def _get_bottom_frame(self) -> QFrame:
        outer = QFrame()
        outer.setStyleSheet(FRAME_SS)
        outer_layout = QVBoxLayout(outer)
        outer_layout.setSpacing(0)
        outer_layout.setContentsMargins(0, 0, 0, 0)

        # ── treeview (output_listbox) ────────────────────────────────
        self.output_listbox = QTreeWidget()
        self.output_listbox.setStyleSheet(TREE_SS)
        self.output_listbox.setSelectionMode(
            QAbstractItemView.SelectionMode.ExtendedSelection)
        self.output_listbox.setAlternatingRowColors(False)
        self.output_listbox.header().setSectionResizeMode(
            QHeaderView.ResizeMode.Interactive)

        outer_layout.addWidget(self.output_listbox, stretch=1)

        # ── bottom button bar ────────────────────────────────────────
        self.entry_frame = QFrame()
        self.entry_frame.setStyleSheet(f"background:{BLACK};")
        btn_layout = QHBoxLayout(self.entry_frame)
        btn_layout.setSpacing(10)
        btn_layout.setContentsMargins(5, 5, 5, 5)

        self.bestellt_label = QLabel(
            'Materialnummer markieren um den Bestellstatus zu wechseln')
        self.bestellt_label.setStyleSheet(BLUE_LABEL_SS)

        self.bestellt_button = QPushButton('Bestellstatus ändern')
        self.bestellt_button.setStyleSheet(BLUE_BTN)
        self.bestellt_button.clicked.connect(lambda: fc.toggle_ordered_status(self))

        self.print_button = QPushButton('Drucken')
        self.print_button.setStyleSheet(GREEN_BTN)
        self.print_button.clicked.connect(lambda: fc.print_screen(self))

        self.delete_button = QPushButton('angezeigte Daten loeschen')
        self.delete_button.setStyleSheet(RED_BTN)
        self.delete_button.clicked.connect(lambda: fc.delete_selected_entries(self))

        if self.user in (ADMIN, EXPERT):
            btn_layout.addWidget(self.bestellt_label, alignment = Qt.AlignmentFlag.AlignLeft)
            btn_layout.addWidget(self.bestellt_button, alignment = Qt.AlignmentFlag.AlignHCenter)

        btn_layout.addWidget(self.print_button, stretch = 1, alignment = Qt.AlignmentFlag.AlignHCenter)

        if self.user == ADMIN:
            btn_layout.addWidget(self.delete_button)

        btn_layout.addStretch()
        outer_layout.addWidget(self.entry_frame)
        return outer

    # ── Scroll speed (replaces on_treeview_scroll) ───────────────────
    def wheelEvent(self, event):
        """Increase vertical scroll speed for the treeview."""
        scroll_speed = 10
        sb = self.output_listbox.verticalScrollBar()
        delta = -int(event.angleDelta().y() / 120) * scroll_speed
        sb.setValue(sb.value() + delta)

    # ── Booking window ───────────────────────────────────────────────
    def open_booking_window(self, title: str, selection: str) -> None:
        """
        Creates the booking window for ingoing / outgoing material.
        Replaces the Tkinter Toplevel + mainloop pattern with a modal QDialog.
        """
        self.user_closed_window = True
        self.cursor.execute(selection)
        materials = self.cursor.fetchall()
        values = [f"{row['MatNr']} {row['Bezeichnung']}" for row in materials]

        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setModal(True)
        dialog.setMinimumSize(500, 200)
        dialog.setStyleSheet(f"background:{BLACK};")

        layout = QVBoxLayout(dialog)

        # Title label
        self.title_label = QLabel(f'\n{title}')
        self.title_label.setStyleSheet(RED_LABEL_SS + 'font:bold 12pt Verdana;')
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.ueberschrift_label = QLabel(
            '\nBitte Material aus Dropdown Menü wählen und Menge eingeben.\n')
        self.ueberschrift_label.setStyleSheet(GREEN_LABEL_SS)
        self.ueberschrift_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        layout.addWidget(self.title_label)
        layout.addWidget(self.ueberschrift_label)

        content_layout = QHBoxLayout()

        # Material selection
        mat_frame = QFrame()
        mat_layout = QVBoxLayout(mat_frame)
        self.mat_label = QLabel('Material')
        self.mat_label.setStyleSheet(GREEN_LABEL_SS)
        self.mat_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        #self.matnr_combobox = AutocompleteComboBox()
        self.matnr_combobox = QComboBox()
        self.matnr_combobox.setEditable(True)
        self.matnr_combobox.setStyleSheet(GREEN_COMBO_SS)
        self.matnr_combobox.setMinimumWidth(400)
        #self.matnr_combobox.set_completion_list(values)
        self.matnr_combobox.addItems(values)

        # completer erstellen - dieser filtert automatisch die Einträge
        completer = QCompleter(values, self.matnr_combobox)
        completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        completer.setFilterMode(Qt.MatchFlag.MatchContains)  # Match anywhere in the string
        self.matnr_combobox.setCompleter(completer)

        mat_layout.addWidget(self.mat_label)
        mat_layout.addWidget(self.matnr_combobox)

        # Menge (quantity) entry
        menge_frame = QFrame()
        menge_layout = QVBoxLayout(menge_frame)
        self.stck_label = QLabel('Menge')
        self.stck_label.setStyleSheet(GREEN_LABEL_SS)
        self.stck_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.stck_entry = QLineEdit()
        self.stck_entry.setStyleSheet(GREEN_ENTRY_SS)

        menge_layout.addWidget(self.stck_label)
        menge_layout.addWidget(self.stck_entry)

        content_layout.addWidget(mat_frame)
        content_layout.addWidget(menge_frame)
        layout.addLayout(content_layout)

        self.ok_button = QPushButton('buchen')
        self.ok_button.setStyleSheet(BLUE_BTN)
        self.ok_button.clicked.connect(lambda: fc.confirm_user_input(self))
        self.stck_entry.returnPressed.connect(lambda: fc.confirm_user_input(self))
        layout.addWidget(self.ok_button)

        # Store dialog reference so fc helpers can close it
        self.ingoing_window = dialog
        self.matnr_combobox.setFocus()

        # Helper vars expected by fc module
        self.ingoing_mat_string_var = self.matnr_combobox
        self.ingoing_menge_var      = self.stck_entry

        dialog.exec()

    # ── InputBox / DialogResult ──────────────────────────────────────
    def InputBox(self, mat_number: str, mat_name: str) -> str:
        dlg = InputBoxDialog(self, mat_number, mat_name)
        dlg.exec()
        return dlg.result_text

    def DialogResult(self, result: str, dialog: QDialog) -> None:
        """Compatibility shim - kept in case fc module calls it directly."""
        self.strDialogResult = result
        dialog.accept()
