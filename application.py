import ctypes
import os
import sqlite3
import tempfile
from collections import defaultdict

from PyQt6.QtCore import QPointF, QRegularExpression, Qt
from PyQt6.QtGui import (
    QBrush,
    QColor,
    QFont,
    QKeySequence,
    QPainter,
    QPixmap,
    QPolygonF,
    QRegularExpressionValidator,
    QShortcut,
)
from PyQt6.QtPrintSupport import QPrinterInfo
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QCompleter,
    QDialog,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QPushButton,
    QSpinBox,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

import functions as fc

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

# Tabellen in "Einträge aus Datenbank löschen" (ADMIN) – müssen in execution_dict,
# filter_dict und deletion_dict stehen
LOESCHBARE_TABELLEN = (
    'Standardmaterial', 'Kleinstmaterial', 'Wareneingang', 'Warenausgang',
    'Warenausgabe_Comline', 'Warenausgang_Kleinstmaterial_ohne_SM_Bezug',
    'Adresszuordnung', 'Veraltetes_Material', 'Jahresinventur_Korrekturdaten', 'Bundles',
)


# ── Version (steht in version.txt, erscheint im Fenstertitel) ────────
def _read_version() -> str:
    pfad = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'version.txt')
    try:
        with open(pfad, encoding='utf-8') as datei:
            for zeile in datei:
                if zeile.strip().startswith('version'):
                    return zeile.split('=', 1)[1].strip()
    except (OSError, IndexError):
        pass
    return '?'


VERSION = _read_version()

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
    # Eingabefeld beim Bearbeiten einer Zelle (Bemerkung): schwarz auf weiß,
    # sonst erbt es Schwarz/Grün und der Text ist während der Eingabe unsichtbar
    f"QTreeWidget QLineEdit{{color:{BLACK};background:{WHITE};"
    f"font:11pt Verdana;border:1px solid {FOREST_GREEN};"
    f"selection-background-color:{FOREST_GREEN};selection-color:{WHITE};}}"
    f"QHeaderView::section{{background:{DARK_GREEN};color:{WHITE};"
    f"font:bold 8pt Verdana;border:1px solid {BLACK};}}"
)

FRAME_SS  = f"background:{BLACK};border:2px groove {FRAME_COLOR};"

# Einheitlicher Stil für die Eingabemasken (Dialoge)
DIALOG_SS = (
    f"QDialog{{background:{BLACK};}}"
    f"QLabel{{color:{FOREST_GREEN};background:{BLACK};font:10pt Verdana;border:0px;}}"
    f"QLineEdit,QSpinBox{{color:{BLACK};background:{WHITE};font:bold 10pt Verdana;"
    f"border:1px solid {FOREST_GREEN};padding:2px;}}"
    f"QComboBox{{color:{FOREST_GREEN};background:{BLACK};font:10pt Verdana;"
    f"border:1px solid {FOREST_GREEN};padding:2px;}}"
    f"QComboBox QAbstractItemView{{color:{FOREST_GREEN};background:{BLACK};}}"
)
ERROR_LABEL_SS = f"color:{ORANGE_RED};background:{BLACK};font:9pt Verdana;border:0px;"
MENU_SS = (f"QMenu{{background:{BLACK};color:{FOREST_GREEN};font:10pt Verdana;"
           f"border:1px solid {FOREST_GREEN};}}"
           f"QMenu::item{{padding:5px 20px;}}"
           f"QMenu::item:selected{{background:{FOREST_GREEN};color:{BLACK};}}")
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


# ── Weiße Sortierpfeile im Spaltenkopf ───────────────────────────────
_PFEIL_SS: str | None = None


def tree_ss() -> str:
    """TREE_SS plus weiße Sortierpfeile. Die Pfeile werden einmal als kleine PNG-Bilder
    in den TEMP-Ordner gezeichnet (Stylesheets können die Pfeilfarbe nicht direkt setzen)."""
    global _PFEIL_SS
    if _PFEIL_SS is None:
        ordner = os.path.join(tempfile.gettempdir(), 'Lagerverwaltung_Pfeile')
        os.makedirs(ordner, exist_ok=True)
        pfade = {}
        for name, punkte in (('auf', [(1, 7), (5, 2), (9, 7)]), ('ab', [(1, 2), (5, 7), (9, 2)])):
            bild = QPixmap(10, 9)
            bild.fill(Qt.GlobalColor.transparent)
            maler = QPainter(bild)
            maler.setRenderHint(QPainter.RenderHint.Antialiasing)
            maler.setPen(Qt.PenStyle.NoPen)
            maler.setBrush(QColor(WHITE))
            maler.drawPolygon(QPolygonF([QPointF(x, y) for x, y in punkte]))
            maler.end()
            pfad = os.path.join(ordner, f'pfeil_{name}.png')
            bild.save(pfad, 'PNG')
            pfade[name] = pfad.replace(os.sep, '/')
        _PFEIL_SS = (
            f"QHeaderView::up-arrow{{image:url({pfade['auf']});width:10px;height:9px;"
            f"subcontrol-position:center right;padding-right:4px;}}"
            f"QHeaderView::down-arrow{{image:url({pfade['ab']});width:10px;height:9px;"
            f"subcontrol-position:center right;padding-right:4px;}}"
        )
    return TREE_SS + _PFEIL_SS


# ── Druckerauswahl ───────────────────────────────────────────────────
class DruckDialog(QDialog):
    """Drucker (Windows-Standarddrucker vorausgewählt) und Anzahl Exemplare."""

    def __init__(self, parent, title: str):
        super().__init__(parent)
        self.setWindowTitle('Drucken')
        self.setModal(True)
        self.setMinimumWidth(420)
        self.setStyleSheet(DIALOG_SS)
        layout = QVBoxLayout(self)

        kopf = QLabel(title)
        kopf.setStyleSheet(GREEN2_LABEL_SS)
        layout.addWidget(kopf)

        layout.addWidget(QLabel('Drucker'))
        self.drucker_combo = QComboBox()
        standard = QPrinterInfo.defaultPrinterName()
        for name in QPrinterInfo.availablePrinterNames():
            self.drucker_combo.addItem(f'{name} (Standard)' if name == standard else name, name)
        index = self.drucker_combo.findData(standard)
        if index >= 0:
            self.drucker_combo.setCurrentIndex(index)
        layout.addWidget(self.drucker_combo)

        layout.addWidget(QLabel('Exemplare'))
        self.exemplare = QSpinBox()
        self.exemplare.setRange(1, 20)
        self.exemplare.setValue(1)
        self.exemplare.setMaximumWidth(80)
        layout.addWidget(self.exemplare)

        self.fehler = QLabel('')
        self.fehler.setStyleSheet(ERROR_LABEL_SS)
        layout.addWidget(self.fehler)

        knoepfe = QHBoxLayout()
        knoepfe.addStretch()
        abbrechen = QPushButton('Abbrechen')
        abbrechen.setStyleSheet(BLUE_BTN)
        abbrechen.clicked.connect(self.reject)
        drucken = QPushButton('Drucken')
        drucken.setStyleSheet(GREEN_BTN)
        drucken.setDefault(True)
        drucken.clicked.connect(self._ok)
        knoepfe.addWidget(abbrechen)
        knoepfe.addWidget(drucken)
        layout.addLayout(knoepfe)

    def _ok(self):
        if self.drucker_combo.count() == 0:
            self.fehler.setText('Kein Drucker gefunden.')
            return
        self.accept()

    def auswahl(self) -> tuple[str, int]:
        return self.drucker_combo.currentData(), self.exemplare.value()


# ── Material anlegen (Standard- oder Kleinstmaterial) ────────────────
class MaterialAnlegenDialog(QDialog):

    def __init__(self, parent: 'App', tabelle: str):
        super().__init__(parent)
        self.app = parent
        self.tabelle = tabelle   # 'Standardmaterial' oder 'Kleinstmaterial'
        self.setWindowTitle(f'{tabelle} anlegen')
        self.setModal(True)
        self.setMinimumWidth(520)
        self.setStyleSheet(DIALOG_SS)
        layout = QVBoxLayout(self)

        kopf = QLabel(f'{tabelle} anlegen')
        kopf.setStyleSheet(GREEN2_LABEL_SS)
        layout.addWidget(kopf)

        grid = QGridLayout()
        self.matnr = QLineEdit()
        self.matnr.setValidator(QRegularExpressionValidator(QRegularExpression(r'\d{0,8}'), self))
        self.matnr.setPlaceholderText('8-stellig')
        self.bezeichnung = QLineEdit()
        self.einheit = QComboBox()
        self.einheit.setEditable(True)
        einheiten = [r[0] for r in parent.cursor.execute(
            "SELECT Einheit FROM (SELECT Einheit FROM Standardmaterial "
            "UNION ALL SELECT Einheit FROM Kleinstmaterial) "
            "WHERE Einheit IS NOT NULL AND Einheit <> '' "
            "GROUP BY Einheit ORDER BY COUNT(*) DESC").fetchall()]
        self.einheit.addItems(einheiten or ['ST'])
        self.grenzwert = QSpinBox()
        self.grenzwert.setRange(0, 100000)
        self.auffuellen = QSpinBox()
        self.auffuellen.setRange(0, 100000)
        self.auffuellen.setValue(1)

        for zeile, (text, widget) in enumerate((
                ('Materialnummer', self.matnr), ('Bezeichnung', self.bezeichnung),
                ('Einheit', self.einheit), ('Grenzwert', self.grenzwert),
                ('Auffüllen auf', self.auffuellen))):
            grid.addWidget(QLabel(text), zeile, 0)
            grid.addWidget(widget, zeile, 1)
        layout.addLayout(grid)

        self.fehler = QLabel('')
        self.fehler.setStyleSheet(ERROR_LABEL_SS)
        self.fehler.setWordWrap(True)
        layout.addWidget(self.fehler)
        for feld in (self.matnr, self.bezeichnung):
            feld.textEdited.connect(lambda _: self.fehler.setText(''))

        knoepfe = QHBoxLayout()
        knoepfe.addStretch()
        abbrechen = QPushButton('Abbrechen')
        abbrechen.setStyleSheet(BLUE_BTN)
        abbrechen.clicked.connect(self.reject)
        anlegen = QPushButton('Anlegen')
        anlegen.setStyleSheet(GREEN_BTN)
        anlegen.setDefault(True)
        anlegen.clicked.connect(self._anlegen)
        knoepfe.addWidget(abbrechen)
        knoepfe.addWidget(anlegen)
        layout.addLayout(knoepfe)
        self.matnr.setFocus()

    def _anlegen(self):
        matnr       = self.matnr.text().strip()
        bezeichnung = self.bezeichnung.text().strip()
        einheit     = self.einheit.currentText().strip().upper()
        if len(matnr) != 8:
            self.fehler.setText('Die Materialnummer muss 8 Ziffern haben.')
            self.matnr.setFocus()
            return
        if not bezeichnung:
            self.fehler.setText('Bitte eine Bezeichnung eingeben.')
            self.bezeichnung.setFocus()
            return
        if not einheit:
            self.fehler.setText('Bitte eine Einheit angeben.')
            return
        vorhanden = self.app.cursor.execute(
            "SELECT 'Standardmaterial' FROM Standardmaterial WHERE MatNr = ? "
            "UNION SELECT 'Kleinstmaterial' FROM Kleinstmaterial WHERE MatNr = ?",
            (int(matnr), int(matnr))).fetchone()
        if vorhanden:
            self.fehler.setText(f'Materialnummer {matnr} gibt es schon ({vorhanden[0]}).')
            self.matnr.setFocus()
            return
        daten = (int(matnr), bezeichnung, einheit,
                 self.grenzwert.value(), self.auffuellen.value(), False)
        anlegen = (fc.add_standardmaterial if self.tabelle == 'Standardmaterial'
                   else fc.add_kleinstmaterial)
        try:
            anlegen(self.app.connection, self.app.cursor, daten)
        except fc.DatenbankBelegt:
            fc.meldung_belegt()
            return
        except sqlite3.IntegrityError:
            self.fehler.setText(f'Materialnummer {matnr} wurde inzwischen von jemand anderem angelegt.')
            return
        self.accept()


# ── Adressdaten aus Auftragslisten einlesen ──────────────────────────
class AdressenEinlesenDialog(QDialog):

    def __init__(self, parent: 'App'):
        super().__init__(parent)
        self.app = parent
        self.adressen: list = []
        self.gespeichert = 0
        self.setWindowTitle('Adressdaten einlesen')
        self.setModal(True)
        self.setMinimumSize(900, 500)
        self.setStyleSheet(DIALOG_SS)
        layout = QVBoxLayout(self)

        zeile = QHBoxLayout()
        self.dateien = QLineEdit()
        self.dateien.setReadOnly(True)
        self.dateien.setPlaceholderText('Auftragslisten (Excel) wählen')
        waehlen = QPushButton('Dateien wählen …')
        waehlen.setStyleSheet(BLUE_BTN)
        waehlen.clicked.connect(self._dateien_waehlen)
        zeile.addWidget(self.dateien, stretch=1)
        zeile.addWidget(waehlen)
        layout.addLayout(zeile)

        self.info = QLabel('Noch keine Dateien gewählt.')
        layout.addWidget(self.info)

        self.vorschau = QTreeWidget()
        self.vorschau.setStyleSheet(tree_ss())
        self.vorschau.setRootIsDecorated(False)
        self.vorschau.setHeaderLabels(['SM Nummer', 'VPSZ', 'Adresse'])
        self.vorschau.setColumnWidth(0, 120)
        self.vorschau.setColumnWidth(1, 120)
        fc.sortierung_einrichten(self.vorschau)
        layout.addWidget(self.vorschau, stretch=1)

        knoepfe = QHBoxLayout()
        knoepfe.addStretch()
        abbrechen = QPushButton('Abbrechen')
        abbrechen.setStyleSheet(BLUE_BTN)
        abbrechen.clicked.connect(self.reject)
        self.uebernehmen = QPushButton('Adressen übernehmen')
        self.uebernehmen.setStyleSheet(GREEN_BTN)
        self.uebernehmen.setEnabled(False)
        self.uebernehmen.clicked.connect(self._uebernehmen)
        knoepfe.addWidget(abbrechen)
        knoepfe.addWidget(self.uebernehmen)
        layout.addLayout(knoepfe)

    def _dateien_waehlen(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, 'Auftragslisten auswählen', '',
            'Excel (*.xlsx *.xlsm);;Alle Dateien (*.*)')
        if not paths:
            return
        self.dateien.setText('; '.join(os.path.basename(p) for p in paths))
        self.vorschau.clear()
        try:
            self.adressen, bekannt = fc.adressen_aus_dateien(self.app.cursor, paths)
        except KeyError as err:
            self.adressen = []
            self.info.setText(f'Spalte {err} fehlt – ist das die richtige Auftragsliste?')
            self.uebernehmen.setEnabled(False)
            return
        except Exception as err:
            self.adressen = []
            self.info.setText(f'Datei konnte nicht gelesen werden: {err}')
            self.uebernehmen.setEnabled(False)
            return
        for sm, vpsz, adresse in self.adressen:
            fc.SortItem(self.vorschau, [sm, vpsz, adresse])
        self.info.setText(f'{len(self.adressen)} neue Adressen gefunden '
                          f'({bekannt} schon bekannt, übersprungen).')
        fc._sortierung_wiederherstellen(self.vorschau)
        self.uebernehmen.setText(f'{len(self.adressen)} Adressen übernehmen')
        self.uebernehmen.setEnabled(bool(self.adressen))

    def _uebernehmen(self):
        try:
            self.gespeichert = fc.adressen_speichern(self.app.connection, self.adressen)
        except fc.DatenbankBelegt:
            fc.meldung_belegt()
            return
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

        self.setWindowTitle(f'Lagerverwaltung Comline {VERSION}')
        self.setMinimumSize(1200, 768)
        self.setStyleSheet(f"background:{BLACK};")

        # context menu (right-click) – connect to fc helper
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(
            lambda pos: fc.show_context_menu(pos, self))

        # neue Spalten (Bemerkung) bei Bedarf anlegen
        if not fc.ensure_schema(self.connection):
            fc.meldung_belegt()

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

    def load_material_data(self) -> None:
        """Materialstammdaten aus der Datenbank laden (beim Start, nach DB-Wechsel
        und nach dem Anlegen von Material)."""
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

        self.load_material_data()

        self.columns_dict = {
            'Bemerkung':                        (400, Qt.AlignmentFlag.AlignLeft),
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

        # Filterabfragen für "Einträge aus Datenbank löschen" (ADMIN):
        # (SQL, Filterfelder in Reihenfolge der ?) – Mat.Nr. / Bezeichnung zählt für beides
        mat = '(MatNr LIKE ? OR Bezeichnung LIKE ?)'
        self.execution_dict = {
            'Kleinstmaterial':
                (f'SELECT * FROM Kleinstmaterial WHERE {mat}', ('matnr', 'matnr')),
            'Standardmaterial':
                (f'SELECT * FROM Standardmaterial WHERE {mat}', ('matnr', 'matnr')),
            'Warenausgabe_Comline':
                (f'SELECT * FROM Warenausgabe_Comline WHERE SM_Nummer LIKE ? AND {mat}', ('sm', 'matnr', 'matnr')),
            'Warenausgang':
                (f'SELECT * FROM Warenausgang WHERE SM_Nummer LIKE ? AND {mat}', ('sm', 'matnr', 'matnr')),
            'Wareneingang':
                ('SELECT * FROM Wareneingang WHERE ID LIKE ?', ('posnr',)),
            'Warenausgang_Kleinstmaterial_ohne_SM_Bezug':
                ('SELECT * FROM Warenausgang_Kleinstmaterial_ohne_SM_Bezug WHERE ID LIKE ?', ('posnr',)),
            'Adresszuordnung':
                ('SELECT * FROM Adresszuordnung WHERE SM_Nummer LIKE ?', ('sm',)),
            'Veraltetes_Material':
                (f'SELECT * FROM Veraltetes_Material WHERE {mat}', ('matnr', 'matnr')),
            'Jahresinventur_Korrekturdaten':
                (f'SELECT * FROM Jahresinventur_Korrekturdaten WHERE {mat}', ('matnr', 'matnr')),
            'Bundles':
                (f'SELECT * FROM Bundles WHERE {mat}', ('matnr', 'matnr')),
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
            'Bundles':
                ([self.matnr_entry], [self.sm_entry, self.posnr_entry])
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
            'Bundles':
                ('DELETE FROM Bundles WHERE ID = ?', 'posnr,'),
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

        # Optionsmenü (nur EXPERT / ADMIN): Stammdaten pflegen
        if self.user in (EXPERT, ADMIN):
            self.button_optionen = QPushButton('Optionen  ▾')
            self.button_optionen.setStyleSheet(BLUE_BTN + "QPushButton::menu-indicator{width:0px;}")
            menu = QMenu(self.button_optionen)
            menu.setStyleSheet(MENU_SS)
            menu.addAction('Standardmaterial - neuer Eintrag in Datenbank',
                           lambda: self.material_anlegen('Standardmaterial'))
            menu.addAction('Kleinstmaterial - neuer Eintrag in Datenbank',
                           lambda: self.material_anlegen('Kleinstmaterial'))
            menu.addSeparator()
            menu.addAction('Adressen - neu in Datenbank einlesen', self.adressen_einlesen)
            self.button_optionen.setMenu(menu)
            layout.addWidget(self.button_optionen)
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

        self.button_noch_offen = QPushButton('gebucht - noch \nnicht zusammengestellt')
        self.button_noch_offen.setStyleSheet(GREEN_BTN)
        self.button_noch_offen.clicked.connect(lambda: fc.show_noch_offen(self))

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
                    self.button_noch_offen, self.button_wareneingang,
                    self.button_warenausgang, self.button_sm_auftrag):
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
        # Nur Tabellen anbieten, für die es Filter und Löschbefehl gibt
        # (nicht z. B. Noch_Offen_Status / Vollstaendig_Status / sqlite_sequence)
        values = [t[0] for t in tables if t[0] in LOESCHBARE_TABELLEN]

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

        matnr_label = QLabel(' Daten filtern:     Mat.Nr. / Bezeichnung:')
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
        self.output_listbox.setStyleSheet(tree_ss())
        self.output_listbox.setSelectionMode(
            QAbstractItemView.SelectionMode.ExtendedSelection)
        self.output_listbox.setAlternatingRowColors(False)
        self.output_listbox.header().setSectionResizeMode(
            QHeaderView.ResizeMode.Interactive)
        # Sortieren per Klick auf den Spaltenkopf
        fc.sortierung_einrichten(self.output_listbox)
        # Bemerkungen: Bearbeiten nur gezielt über Doppelklick / F2 (siehe fc.bemerkung_*)
        self.bem_col = None
        self.bem_speichern = None
        self.output_listbox.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.output_listbox.itemDoubleClicked.connect(
            lambda item, col: fc.bemerkung_bearbeiten(self, item, col))
        self.output_listbox.itemChanged.connect(
            lambda item, col: fc.bemerkung_geaendert(self, item, col))
        f2 = QShortcut(QKeySequence(Qt.Key.Key_F2), self.output_listbox)
        f2.setContext(Qt.ShortcutContext.WidgetShortcut)
        f2.activated.connect(
            lambda: fc.bemerkung_bearbeiten(self, self.output_listbox.currentItem()))

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

        self.button_aenderungen_bestaetigen = QPushButton('Änderungen bestätigen')
        self.button_aenderungen_bestaetigen.setStyleSheet(BLUE_BTN)
        self.button_aenderungen_bestaetigen.clicked.connect(
            lambda: fc.bestaetigen_noch_offen(self))
        self.button_aenderungen_bestaetigen.hide()  # standardmäßig versteckt

        if self.user in (ADMIN, EXPERT):
            btn_layout.addWidget(self.bestellt_label)
            btn_layout.addWidget(self.bestellt_button, stretch = 1)

        btn_layout.addWidget(self.print_button, stretch = 1)

        if self.user == ADMIN:
            btn_layout.addWidget(self.delete_button, stretch = 1)

        btn_layout.addWidget(self.button_aenderungen_bestaetigen, stretch = 1)
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
    def open_booking_window(self, title: str, selection: str,
                            mit_bemerkung: bool = False) -> None:
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

        # Bemerkung (optional, nur beim Wareneingang)
        self.bem_entry = QLineEdit()
        self.bem_entry.setStyleSheet(GREEN_ENTRY_SS)
        if mit_bemerkung:
            bem_label = QLabel('Bemerkung (optional)')
            bem_label.setStyleSheet(GREEN_LABEL_SS)
            layout.addWidget(bem_label)
            layout.addWidget(self.bem_entry)
            self.bem_entry.returnPressed.connect(lambda: fc.confirm_user_input(self))

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
        self.ingoing_bemerkung_var  = self.bem_entry

        dialog.exec()

    # ── Druckerauswahl ───────────────────────────────────────────────
    def choose_printer(self, title: str) -> tuple[str, int] | None:
        """Zeigt die Druckerauswahl. (Drucker, Exemplare) oder None bei Abbruch."""
        dlg = DruckDialog(self, title)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return None
        return dlg.auswahl()

    # ── Optionsmenü ──────────────────────────────────────────────────
    def material_anlegen(self, tabelle: str) -> None:
        dlg = MaterialAnlegenDialog(self, tabelle)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self.load_material_data()
            fc.refresh_view(self)

    def adressen_einlesen(self) -> None:
        dlg = AdressenEinlesenDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            # Windows-Meldung wie im übrigen Programm (QMessageBox erbt den schwarzen Hintergrund)
            ctypes.windll.user32.MessageBoxW(
                0, f'{dlg.gespeichert} neue Adressen aufgenommen.', 'Adressdaten einlesen', 64)
            fc.refresh_view(self)

    # ── InputBox / DialogResult ──────────────────────────────────────
    def InputBox(self, mat_number: str, mat_name: str) -> str:
        dlg = InputBoxDialog(self, mat_number, mat_name)
        dlg.exec()
        return dlg.result_text

    def DialogResult(self, result: str, dialog: QDialog) -> None:
        """Compatibility shim - kept in case fc module calls it directly."""
        self.strDialogResult = result
        dialog.accept()
