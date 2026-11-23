import ast
import ctypes
import datetime
import itertools
import os
import re
import sqlite3
import tempfile
from collections import defaultdict
from contextlib import contextmanager

import openpyxl as xl
import pandas as pd
import win32com.client
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QCursor
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QHBoxLayout,
    QLineEdit,
    QMenu,
    QTreeWidget,
    QTreeWidgetItem,
    QWidget,
)

import application

# ── Datenbankzugriff bei mehreren gleichzeitigen Nutzern ─────────────
# Die Datenbank liegt auf einem Netzlaufwerk und wird von mehreren Personen
# gleichzeitig benutzt. Deshalb:
#  - Journal-Modus bleibt DELETE (WAL ist auf Netzlaufwerken nicht sicher),
#  - jede Verbindung wartet bis zu DB_TIMEOUT Sekunden, wenn gerade jemand
#    anderes schreibt (statt sofort mit "database is locked" abzubrechen),
#  - jeder Schreibvorgang läuft als eigene Transaktion mit BEGIN IMMEDIATE:
#    Prüfen (z. B. "schon gebucht?") und Schreiben passieren in einem Zug,
#    niemand sonst kann dazwischen schreiben. Bei Fehler: alles zurück.
#  - Während eine Transaktion offen ist, wird nie auf Benutzereingaben gewartet.
DB_TIMEOUT = 20  # Sekunden


class DatenbankBelegt(Exception):
    """Die Datenbank war länger als DB_TIMEOUT durch einen anderen Nutzer gesperrt."""


def _ist_gesperrt(err: Exception) -> bool:
    text = str(err).lower()
    return 'locked' in text or 'busy' in text


def connect_db(path_to_db: str) -> sqlite3.Connection:
    connection = sqlite3.connect(path_to_db, timeout=DB_TIMEOUT)
    connection.row_factory = sqlite3.Row  # Standard-Rückgabeformat
    return connection


@contextmanager
def schreibvorgang(connection: sqlite3.Connection):
    """Schreibtransaktion: commit am Ende, rollback bei jedem Fehler.

    Löst DatenbankBelegt aus, wenn die Datenbank zu lange gesperrt war.
    """
    if connection.in_transaction:
        connection.commit()
    try:
        connection.execute('BEGIN IMMEDIATE')
    except sqlite3.OperationalError as err:
        if _ist_gesperrt(err):
            raise DatenbankBelegt() from err
        raise
    try:
        yield
        connection.commit()
    except sqlite3.OperationalError as err:
        connection.rollback()
        if _ist_gesperrt(err):
            raise DatenbankBelegt() from err
        raise
    except BaseException:
        connection.rollback()
        raise


def meldung_belegt() -> None:
    ctypes.windll.user32.MessageBoxW(
        0,
        "Die Datenbank wird gerade von jemand anderem bearbeitet.\n"
        "Es wurde nichts gespeichert – bitte in ein paar Sekunden erneut versuchen.",
        "Datenbank belegt", 48)


def _py(value):
    """numpy-/pandas-Werte in normale Python-Werte umwandeln (für SQL-Parameter)."""
    return value.item() if hasattr(value, 'item') else value


def _nr_text(value) -> str:
    """Nummer aus Excel als Text, ohne '.0' (z. B. SM-Nummer 210979478.0)."""
    value = _py(value)
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


# ── Filter "Mat.Nr. / Bezeichnung" (ab 1.1a) ─────────────────────────
def _material_filter(text: str, eigene_bezeichnung: bool = True) -> tuple[str, list]:
    """SQL-Bedingung und Parameter für das Filterfeld 'Mat.Nr. / Bezeichnung'.

    Trifft zu, wenn der Text in der Materialnummer oder in der Bezeichnung aus den
    Stammdaten (Standard-/Kleinstmaterial) vorkommt. Mit eigene_bezeichnung=True
    zusätzlich in der Spalte 'Bezeichnung' der abgefragten Tabelle (z. B. Text aus PSL).
    Groß-/Kleinschreibung egal (bei Umlauten nicht).
    """
    muster = f'%{text.strip()}%'
    bedingung = ("(MatNr LIKE ? OR MatNr IN ("
                 "SELECT MatNr FROM Standardmaterial WHERE Bezeichnung LIKE ? "
                 "UNION SELECT MatNr FROM Kleinstmaterial WHERE Bezeichnung LIKE ?)")
    params = [muster, muster, muster]
    if eigene_bezeichnung:
        bedingung += " OR Bezeichnung LIKE ?"
        params.append(muster)
    return bedingung + ")", params


# ── Datenbankschema ergänzen (ab 1.0b) ───────────────────────────────
# Neue Spalten werden beim Start automatisch angelegt. Ältere Programmversionen
# laufen damit weiter (Standard- und Kleinstmaterial bekommen die Spalte im
# selben Zug, weil sie per "SELECT * ... UNION SELECT * ..." gelesen werden).
BEMERKUNG_TABELLEN = ('Standardmaterial', 'Kleinstmaterial', 'Wareneingang', 'Noch_Offen_Status')


def ensure_schema(connection: sqlite3.Connection) -> bool:
    """Legt fehlende Tabellen/Spalten an. False, wenn die DB gerade belegt war."""
    def _spalten(tabelle):
        return {row[1] for row in connection.execute(f"PRAGMA table_info({tabelle})").fetchall()}
    try:
        _ensure_noch_offen_table(connection.cursor(), connection)
        fehlend = [t for t in BEMERKUNG_TABELLEN if 'Bemerkung' not in _spalten(t)]
        if fehlend:
            with schreibvorgang(connection):
                for tabelle in fehlend:
                    if 'Bemerkung' not in _spalten(tabelle):  # evtl. inzwischen von anderem angelegt
                        connection.execute(f"ALTER TABLE {tabelle} ADD COLUMN Bemerkung TEXT")
    except DatenbankBelegt:
        return False
    return True


def _bemerkung(row) -> str:
    """Bemerkung einer Zeile (leer, falls Spalte noch fehlt)."""
    try:
        return row['Bemerkung'] or ''
    except (IndexError, KeyError):
        return ''


def refresh_view(app: application.App) -> None:
    """Aktuelle Ansicht neu aufbauen (aktiver Knopf ist deaktiviert, daher erst freigeben)."""
    button = app.disabled_button
    button.setEnabled(True)
    button.click()


# ── Bemerkungen direkt in der Tabelle bearbeiten ─────────────────────
# Jede Ansicht mit Bemerkungsspalte meldet Spalte und Speicherfunktion an.
# Bearbeitbar sind nur Zeilen, die einen Schlüssel (UserRole) tragen.
# Gespeichert wird sofort nach dem Bearbeiten (Enter oder Zelle verlassen).
BEM_KEY  = Qt.ItemDataRole.UserRole
BEM_ORIG = Qt.ItemDataRole.UserRole + 1


def _bemerkung_anmelden(app: application.App, col: int, speichern) -> None:
    app.bem_col = col
    app.bem_speichern = speichern


def _bemerkung_setzen(item: QTreeWidgetItem, col: int, key, text: str) -> None:
    item.setData(col, BEM_KEY, key)
    item.setData(col, BEM_ORIG, text)
    item.setText(col, text)
    item.setToolTip(col, 'Doppelklick oder F2 zum Bearbeiten')


def bemerkung_bearbeiten(app: application.App, item: QTreeWidgetItem | None,
                         col: int | None = None) -> None:
    """Startet die Bearbeitung der Bemerkungszelle (Doppelklick / F2)."""
    bem_col = getattr(app, 'bem_col', None)
    if item is None or bem_col is None or app.user not in (application.EXPERT, application.ADMIN):
        return
    if col is not None and col != bem_col:
        return
    if item.data(bem_col, BEM_KEY) is None:
        return
    item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
    app.output_listbox.editItem(item, bem_col)


def bemerkung_geaendert(app: application.App, item: QTreeWidgetItem, col: int) -> None:
    """itemChanged: speichert die Bemerkung, wenn sie sich wirklich geändert hat."""
    if col != getattr(app, 'bem_col', None):
        return
    key = item.data(col, BEM_KEY)
    if key is None:
        return
    neu, alt = item.text(col).strip(), item.data(col, BEM_ORIG) or ''
    if neu == alt:
        return
    try:
        with schreibvorgang(app.connection):
            app.bem_speichern(key, neu or None)
    except DatenbankBelegt:
        meldung_belegt()
        neu = alt  # nicht gespeichert -> alten Text wieder anzeigen
    app.output_listbox.blockSignals(True)
    _bemerkung_setzen(item, col, key, neu)
    app.output_listbox.blockSignals(False)


# ── Sortieren per Klick auf den Spaltenkopf (ab 1.1c) ────────────────
# Klick 1: aufsteigend, Klick 2: absteigend, Klick 3: ursprüngliche Reihenfolge.
# Gruppenzeilen (Standardmaterial, Kleinstmaterial, …) behalten ihre Reihenfolge,
# sortiert wird innerhalb der Gruppen. Leere Zellen stehen immer am Ende.
# Zahlen werden als Zahlen sortiert (auch '5 (+3)', '-11', '1,5'), Datumsangaben
# 'TT.MM.JJJJ [hh:mm[:ss]]' chronologisch, Text natürlich ('2m' vor '10m').
_DATUM = re.compile(r'^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$')
_ZAHL  = re.compile(r'^(-?\d+(?:[.,]\d+)?)(?:\s+(.*))?$')
_ZIFFERN = re.compile(r'(\d+)')


def _natuerlich(text: str) -> list:
    teile = _ZIFFERN.split(text.casefold())
    return [int(t) if i % 2 else t for i, t in enumerate(teile)]


def _sortschluessel(text: str):
    t = text.strip()
    if not t:
        return None
    m = _DATUM.match(t)
    if m:
        return (0, tuple(int(x or 0) for x in (m[3], m[2], m[1], m[4], m[5], m[6])))
    m = _ZAHL.match(t)
    if m:
        return (1, float(m[1].replace(',', '.')), _natuerlich(m[2] or ''))
    return (2, _natuerlich(t))


class SortItem(QTreeWidgetItem):
    """Tabellenzeile mit sinnvoller Sortierung (Zahl, Datum, Text, Häkchen)."""
    _zaehler = itertools.count()

    def __init__(self, *args):
        super().__init__(*args)
        self.nr = next(SortItem._zaehler)   # Einfügereihenfolge = ursprüngliche Reihenfolge
        self.gruppe = False
        self._cache = {}                      # Spalte -> (Text, Sortierschlüssel)

    def _schluessel(self, col: int):
        tree = self.treeWidget()
        widget = tree.itemWidget(self, col) if tree else None
        if widget is not None:
            cb = widget.findChild(QCheckBox)
            if cb is not None:
                return (3, int(cb.isChecked()))
        text = self.text(col)
        gemerkt = self._cache.get(col)
        if gemerkt is None or gemerkt[0] != text:
            gemerkt = (text, _sortschluessel(text))
            self._cache[col] = gemerkt
        return gemerkt[1]

    def __lt__(self, other):
        tree = self.treeWidget()
        andere_nr = getattr(other, 'nr', 0)
        if tree is None or getattr(tree, 'orig_reihenfolge', False):
            return self.nr < andere_nr
        absteigend = tree.header().sortIndicatorOrder() == Qt.SortOrder.DescendingOrder
        if self.gruppe or getattr(other, 'gruppe', False):
            # Gruppen immer in ursprünglicher Reihenfolge (Qt dreht bei 'absteigend' um)
            return (self.nr < andere_nr) != absteigend
        col = tree.sortColumn()
        a = self._schluessel(col)
        b = other._schluessel(col) if isinstance(other, SortItem) else _sortschluessel(other.text(col))
        if a is None or b is None:
            if a is None and b is None:
                return self.nr < andere_nr
            # leere Zellen ans Ende – in beiden Richtungen
            return (a is None) == absteigend
        if a == b:
            return (self.nr < andere_nr) != absteigend   # gleiche Werte: ursprüngliche Reihenfolge
        if a[0] != b[0]:
            return a[0] < b[0]
        return a < b


def sortierung_einrichten(tree: QTreeWidget) -> None:
    """Spaltenköpfe anklickbar machen (einmal je Tabelle aufrufen)."""
    header = tree.header()
    header.setSectionsClickable(True)
    header.setSortIndicatorShown(True)
    header.setSortIndicator(-1, Qt.SortOrder.AscendingOrder)
    tree.sort_zustand = (None, None, None)   # (Spaltenköpfe, Spalte, Reihenfolge)
    header.sectionClicked.connect(lambda col: spalte_sortieren(tree, col))


def _kopf(tree: QTreeWidget) -> tuple:
    return tuple(tree.headerItem().text(c) for c in range(tree.columnCount()))


def _sortieren(tree: QTreeWidget, col: int | None, order) -> None:
    if col is None:
        tree.orig_reihenfolge = True
        tree.sortItems(0, Qt.SortOrder.AscendingOrder)
        tree.orig_reihenfolge = False
        tree.header().setSortIndicator(-1, Qt.SortOrder.AscendingOrder)
        tree.sort_zustand = (None, None, None)
    else:
        tree.header().setSortIndicator(col, order)
        tree.sortItems(col, order)
        tree.sort_zustand = (_kopf(tree), col, order)


def spalte_sortieren(tree: QTreeWidget, col: int) -> None:
    kopf, alt_col, alt_order = getattr(tree, 'sort_zustand', (None, None, None))
    if not tree.headerItem().text(col).strip():
        # Spalte ohne Überschrift (Baumspalte, Rest): nicht sortieren, Anzeige zurücksetzen
        tree.header().setSortIndicator(alt_col if alt_col is not None else -1,
                                       alt_order or Qt.SortOrder.AscendingOrder)
        return
    if alt_col != col or kopf != _kopf(tree):
        _sortieren(tree, col, Qt.SortOrder.AscendingOrder)
    elif alt_order == Qt.SortOrder.AscendingOrder:
        _sortieren(tree, col, Qt.SortOrder.DescendingOrder)
    else:
        _sortieren(tree, None, None)


def _sortierung_wiederherstellen(tree: QTreeWidget) -> None:
    """Nach dem Neuaufbau (z. B. beim Filtern) die gewählte Sortierung wieder anwenden –
    aber nur in derselben Ansicht (gleiche Spaltenköpfe)."""
    kopf, col, order = getattr(tree, 'sort_zustand', (None, None, None))
    if col is not None and kopf == _kopf(tree):
        _sortieren(tree, col, order)
    else:
        tree.header().setSortIndicator(-1, Qt.SortOrder.AscendingOrder)
        tree.sort_zustand = (None, None, None)


# ── Hilfsfunktion: Widget aktivieren / deaktivieren ─────────────────
def set_widget_status(enabled: list, disabled: list, show:list | None = None, hide:list | None = None) -> None:
    
    for widget in disabled:
        if widget is not None:
            if isinstance(widget, QLineEdit):
                widget.clear()
            widget.setEnabled(False)

    for widget in enabled:
        if widget is not None:
            widget.setEnabled(True)    
    
    if show:
        for widget in show:
            widget.show()
    if hide:
        for widget in hide:
            widget.hide()


# ── Hilfsfunktion: alle Signal-Verbindungen der Filter-Widgets lösen ─
def unbind_all_widgets(app: application.App) -> None:
    """Trennt alle zuvor verbundenen Signale der Filter-Widgets."""
    for widget in (app.matnr_entry, app.sm_entry, app.posnr_entry):
        try:
            widget.returnPressed.disconnect()
        except TypeError:
            pass
        try:
            widget.textChanged.disconnect()
        except TypeError:
            pass
    try:
        app.combobox_loeschen.currentIndexChanged.disconnect()
    except TypeError:
        pass
    # Beim Wechsel der Ansicht: Button verstecken und pending-Zustand verwerfen.
    # show_noch_offen zeigt den Button bei Bedarf selbst wieder an.
    try:
        app.button_aenderungen_bestaetigen.hide()
    except AttributeError:
        pass  # Button existiert noch nicht beim ersten Aufruf waehrend __init__
    if hasattr(app, '_noch_offen_pending'):
        app._noch_offen_pending = {}
    # Bemerkungsspalte gilt nur für die Ansicht, die sie anmeldet
    app.bem_col = None
    app.bem_speichern = None


# ── Datenbank öffnen / Pfad ändern ──────────────────────────────────
def open_db() -> str:
    settings = f"{os.environ['USERPROFILE']}\\Documents\\Lagerverwaltung_settings.txt"
    if not os.path.exists(settings):
        with open(settings, 'w'):
            pass
    with open(settings, 'r') as file:
        path_to_db = file.read()
    if not os.path.exists(path_to_db):
        ctypes.windll.user32.MessageBoxW(
            0,
            "Die Datenbank 'Lagerverwaltung_Datanbank.db' konnte nicht gefunden werden.\n"
            "Bitte im nächsten Fenster die Datenbank auswählen",
            "Pfad zur Datenbank suchen...",
            64
        )
        path_to_db = change_db_path()
    return path_to_db


def change_db_path(app: application.App = None) -> str:
    settings = f"{os.environ['USERPROFILE']}\\Documents\\Lagerverwaltung_settings.txt"
    path_to_db, _ = QFileDialog.getOpenFileName(
        None,
        'Datenbank auswählen',
        'Lagerverwaltung_Datenbank.db',
        'Datenbank (*.db);;Alle Dateien (*.*)'
    )
    if not path_to_db:
        return ''
    with open(settings, 'w') as file:
        file.write(path_to_db)
    if app:
        app.path_to_db = path_to_db
        color = '#228B22' if 'Service-Center' in path_to_db else '#FF4500'
        app.label_top_path_to_db.setText(f'Pfad zur Datenbank: {path_to_db}')
        app.label_top_path_to_db.setStyleSheet(
            f"color:{color};background:#000000;font:10pt Verdana; border: 1px solid {color}")
        app.connection.close()
        app.connection = connect_db(path_to_db)
        app.cursor = app.connection.cursor()
        if not ensure_schema(app.connection):
            meldung_belegt()
        # Materialstammdaten der neuen Datenbank laden (vorher blieben die alten stehen)
        app.load_material_data()
        # Zuletzt aktive Ansicht neu aufbauen (click() auf deaktivierten Knopf wirkte nicht)
        refresh_view(app)
    return path_to_db


# ── Kontextmenü (Rechtsklick) ────────────────────────────────────────
def show_context_menu(pos, app: application.App) -> None:
    selections = app.output_listbox.selectedItems()
    if not selections:
        return
    # Spaltenheader-Texte ermitteln
    col_count = app.output_listbox.columnCount()
    columns = [app.output_listbox.headerItem().text(c) for c in range(col_count)]

    menu = QMenu(app)
    menu.setStyleSheet(
        "QMenu{background:#000000;color:#228B22;font:10pt Verdana;}"
        "QMenu::item:selected{background:#228B22;color:#000000;}"
    )
    if 'MatNr.' in columns:
        menu.addAction('Kopieren: Materialnummer', lambda: copy_matnr(app))
    if 'SM Nummer / ID' in columns or 'SM Nummer' in columns:
        menu.addAction('Kopieren: SM Nummer', lambda: copy_sm_nr(app))
    if 'Bezeichnung' in columns:
        menu.addAction('Kopieren: Bezeichnung', lambda: copy_name(app))
    menu.addAction('Kopieren: Ganze Zeile', lambda: copy_line(app))
    menu.exec(QCursor.pos())


# ── Kopier-Hilfsfunktionen ───────────────────────────────────────────
def _col_index(app: application.App, col_name: str) -> int:
    """Gibt den Spaltenindex anhand des Header-Textes zurück, oder -1."""
    for c in range(app.output_listbox.columnCount()):
        if app.output_listbox.headerItem().text(c) == col_name:
            return c
    return -1


def copy_matnr(app: application.App) -> None:
    col = _col_index(app, 'MatNr.')
    if col < 0:
        return
    values = [item.text(col) for item in app.output_listbox.selectedItems()
              if item.text(col)]
    QApplication.clipboard().setText('\n'.join(values))


def copy_sm_nr(app: application.App) -> None:
    col = _col_index(app, 'SM Nummer / ID')
    if col < 0:
        col = _col_index(app, 'SM Nummer')
    if col < 0:
        return
    values = [item.text(col) for item in app.output_listbox.selectedItems()
              if item.text(col)]
    QApplication.clipboard().setText('\n'.join(values))


def copy_name(app: application.App) -> None:
    col = _col_index(app, 'Bezeichnung')
    if col < 0:
        return
    values = [item.text(col) for item in app.output_listbox.selectedItems()
              if item.text(col)]
    QApplication.clipboard().setText('\n'.join(values))


def copy_line(app: application.App) -> None:
    col_count = app.output_listbox.columnCount()
    lines = []
    for item in app.output_listbox.selectedItems():
        row_values = [item.text(c) for c in range(col_count)]
        lines.append('\t'.join(row_values))
    QApplication.clipboard().setText('\n'.join(lines))


# ── QTreeWidget-Hilfsfunktionen ──────────────────────────────────────
def _clear_tree(app: application.App) -> None:
    app.output_listbox.clear()
    # Sortierung erst anwenden, wenn die Ansicht fertig aufgebaut ist
    QTimer.singleShot(0, lambda: _sortierung_wiederherstellen(app.output_listbox))


def _setup_columns(app: application.App, columns: list[str]) -> None:
    """Setzt Spalten-Header und -Breiten im QTreeWidget."""
    visible = [c for c in columns if c != 'LAST_COLUMN']
    app.output_listbox.setColumnCount(len(visible) + 1)  # +1 für Baum-Spalte (index 0)
    # Spalte 0 = Baum/Parent-Beschriftung (kein Daten-Header)
    app.output_listbox.headerItem().setText(0, '')
    for i, col in enumerate(visible, start=1):
        app.output_listbox.headerItem().setText(i, col)
        width, _ = app.columns_dict.get(col, (100, Qt.AlignmentFlag.AlignLeft))
        app.output_listbox.setColumnWidth(i, width)
    app.output_listbox.setColumnWidth(0, 220)  # Baum-Spalte


def _insert_parent(app: application.App, text: str,
                   tag: str = 'green', expanded: bool = True) -> QTreeWidgetItem:
    """Fügt eine Parent-Zeile (Gruppe) ins QTreeWidget ein."""
    parent = SortItem(app.output_listbox)
    parent.gruppe = True   # Gruppen behalten ihre Reihenfolge
    parent.setText(0, text)
    parent.setExpanded(expanded)
    app.apply_tag(parent, tag)
    return parent


def _insert_child(parent: QTreeWidgetItem, values: tuple | list,
                  app: application.App, tag: str = '') -> QTreeWidgetItem:
    """Fügt eine Kind-Zeile unter einem Parent ein."""
    child = SortItem(parent)
    # Werte ab Spalte 1 eintragen (Spalte 0 = Baum-Text, bleibt leer)
    for i, val in enumerate(values, start=1):
        child.setText(i, str(val) if val is not None else '')
        child.setExpanded(True)
    if tag:
        app.apply_tag(child, tag)
    return child


# ── Kritisches Material anzeigen ─────────────────────────────────────
def show_critical_material(app: application.App) -> None:
    
    unbind_all_widgets(app)
    app.disabled_button.setEnabled(True)
    app.disabled_button = app.button_krit_mat
    app.button_krit_mat.setEnabled(False)

    enabled  = []
    disabled = [app.sm_entry, app.posnr_entry, app.matnr_entry]
    show     = [app.bestellt_button, app.bestellt_label]
    hide     = [app.delete_button, app.print_button, app.combobox_loeschen]
    
    set_widget_status(enabled, disabled, show, hide)
    

    cursor = app.cursor
    ingoing_dict           = defaultdict(int)
    outgoing_dict          = defaultdict(int)
    already_ordered_dict   = defaultdict(int)
    date_dict              = defaultdict(str)
    mengen_dict            = defaultdict(str)
    bemerkung_dict         = defaultdict(str)

    ingoing              = cursor.execute("SELECT * FROM Wareneingang").fetchall()
    outgoing             = cursor.execute("SELECT * FROM Warenausgang").fetchall()
    outgoing_small       = cursor.execute(
        "SELECT * FROM Warenausgang_Kleinstmaterial_ohne_SM_BEZUG").fetchall()
    all_materials        = cursor.execute(
        "SELECT * FROM Standardmaterial UNION SELECT * FROM Kleinstmaterial").fetchall()

    for row in ingoing:
        ingoing_dict[row['MatNr']] += row['Menge']
    for row in outgoing:
        menge = row['Warenausgangsmenge'] if row['Warenausgangsmenge'] != 0 else row['Bedarfsmenge']
        if row['PosTyp'] == 8:
            outgoing_dict[row['MatNr']] -= menge
        elif row['PosTyp'] == 9:
            outgoing_dict[row['MatNr']] += menge
    for row in outgoing_small:
        outgoing_dict[row['MatNr']] += row['Menge']
    for row in all_materials:
        already_ordered_dict[row['MatNr']] = 'nein' if row['bestellt'] == 0 else 'ja'
        date_dict[row['MatNr']]  = row['Datum'] if row['bestellt'] else ''
        menge = row['Menge'] if row['Menge'] else ''
        mengen_dict[row['MatNr']] = menge if row['bestellt'] else ''
        bemerkung_dict[row['MatNr']] = _bemerkung(row)

    small_material_output    = []
    standard_material_output = []
    for matnr, booked in sorted(outgoing_dict.items()):
        ingoing_mat    = ingoing_dict.get(matnr, 0)
        correction_sum = app.correction_dict.get(matnr, 0)
        bestand        = ingoing_mat - booked + correction_sum
        if bestand <= app.threshhold_dict[matnr] and matnr not in app.deprecated_dict:
            output = [matnr, app.materialnames_dict[matnr], bestand,
                      app.units_dict[matnr], app.recommended_amount_dict[matnr],
                      already_ordered_dict[matnr], mengen_dict[matnr], date_dict[matnr]]
            if matnr in app.standard_materials:
                standard_material_output.append(output)
            elif matnr in app.small_materials:
                small_material_output.append(output)

    columns = ['MatNr.', 'Bezeichnung', 'Bestand', 'Einheit',
               'empfohlene Menge', 'bestellt', 'Menge ', 'Datum', 'Bemerkung', 'LAST_COLUMN']
    bem_col = columns.index('Bemerkung') + 1   # +1 wegen Baum-Spalte 0
    _clear_tree(app)
    _setup_columns(app, columns)

    def _bem_speichern(matnr, text):
        # MatNr steht nur in einer der beiden Tabellen; UPDATE in beiden ist unschädlich
        for tabelle in ('Standardmaterial', 'Kleinstmaterial'):
            app.cursor.execute(f'UPDATE {tabelle} SET Bemerkung = ? WHERE MatNr = ?', (text, matnr))
    _bemerkung_anmelden(app, bem_col, _bem_speichern)

    if standard_material_output:
        parent = _insert_parent(app, 'Standardmaterial')
        for entry in standard_material_output:
            tag = 'red_font' if entry[2] <= 0 else 'green_font'
            child = _insert_child(parent, entry, app, tag)
            _bemerkung_setzen(child, bem_col, entry[0], bemerkung_dict[entry[0]])
    if small_material_output:
        parent = _insert_parent(app, 'Kleinstmaterial')
        for entry in small_material_output:
            child = _insert_child(parent, entry, app)
            _bemerkung_setzen(child, bem_col, entry[0], bemerkung_dict[entry[0]])
    if not standard_material_output and not small_material_output:
        item = SortItem(app.output_listbox)
        item.setText(2, 'Sieht gut aus, wir haben alles. :)')


# ── Bestellstatus umschalten ─────────────────────────────────────────
def toggle_ordered_status(app: application.App) -> None:
    connection = app.connection
    cursor     = app.cursor

    selections = app.output_listbox.selectedItems()
    mat_numbers, mat_names = [], []
    for item in selections:
        values = [item.text(c) for c in range(app.output_listbox.columnCount())]
        if not values[1]:   # leere Zeile oder Gruppenzeile (Standard-/Kleinstmaterial)
            continue
        mat_numbers.append(values[1])   # Spalte 1 = MatNr.
        mat_names.append(values[2])     # Spalte 2 = Bezeichnung

    # 1. Neuen Status ermitteln und ggf. Bestellmenge abfragen
    #    (keine offene Transaktion, solange auf Eingaben gewartet wird)
    aenderungen = []
    for mat_number, mat_name in zip(mat_numbers, mat_names):
        cursor.execute('''SELECT bestellt FROM Standardmaterial WHERE MatNr = ?
                          UNION SELECT bestellt FROM Kleinstmaterial WHERE MatNr = ?''',
                       (mat_number, mat_number))
        status = not cursor.fetchone()[0]
        if status:
            menge = app.InputBox(mat_number, mat_name)
            datum = datetime.datetime.strftime(datetime.datetime.now(), r'%d.%m.%Y %H:%M')
            aenderungen.append((mat_number, status, datum, menge))
        else:
            aenderungen.append((mat_number, status, None, None))

    # 2. Alles in einer Transaktion schreiben
    try:
        with schreibvorgang(connection):
            for mat_number, status, datum, menge in aenderungen:
                for tabelle in ('Standardmaterial', 'Kleinstmaterial'):
                    if status:
                        cursor.execute(f'UPDATE {tabelle} SET bestellt = ?, Datum = ?, Menge = ? '
                                       'WHERE MatNr = ?', (status, datum, menge, mat_number))
                    else:
                        cursor.execute(f'UPDATE {tabelle} SET bestellt = ? WHERE MatNr = ?',
                                       (status, mat_number))
    except DatenbankBelegt:
        meldung_belegt()
    show_critical_material(app)




# ── Noch offen anzeigen ───────────────────────────────────────────────
def _ensure_noch_offen_table(cursor: sqlite3.Cursor, connection: sqlite3.Connection) -> None:
    # Erstellt die Tabellen Noch_Offen_Status und Vollstaendig_Status falls nicht vorhanden.
    # Nur schreiben, wenn wirklich eine fehlt (wird bei jeder Ansicht aufgerufen).
    vorhanden = {row[0] for row in cursor.execute(
        "SELECT name FROM sqlite_master WHERE type = 'table' "
        "AND name IN ('Noch_Offen_Status', 'Vollstaendig_Status')").fetchall()}
    if len(vorhanden) == 2:
        return
    with schreibvorgang(connection):
        cursor.execute(
            "CREATE TABLE IF NOT EXISTS Noch_Offen_Status "
            "(SM_Nummer TEXT PRIMARY KEY, zusammengesucht INTEGER DEFAULT 0)"
        )
        cursor.execute(
            "CREATE TABLE IF NOT EXISTS Vollstaendig_Status "
            "(SM_Nummer TEXT PRIMARY KEY, vollstaendig INTEGER DEFAULT 0)"
        )


def get_open_sm_numbers(cursor: sqlite3.Cursor, connection: sqlite3.Connection) -> list:
    # Gibt alle SM-Nummern zurueck, die in 'noch offen' noch sichtbar sind.
    _ensure_noch_offen_table(cursor, connection)
    all_sm = [row['SM_Nummer'] for row in cursor.execute(
        "SELECT DISTINCT SM_Nummer FROM Warenausgabe_Comline ORDER BY SM_Nummer"
    ).fetchall()]
    done = {row['SM_Nummer'] for row in cursor.execute(
        "SELECT SM_Nummer FROM Noch_Offen_Status WHERE zusammengesucht = 1"
    ).fetchall()}
    return [sm for sm in all_sm if sm not in done]


def _make_cb_container(app: application.App, sm_nr: str,
                        field: str, checked: bool,
                        cb_zusammen: QCheckBox | None = None) -> tuple:
    # Erzeugt eine zentrierte Checkbox.
    # field: 'vollstaendig' oder 'zusammengesucht'
    # cb_zusammen: fuer 'vollstaendig'-Checkbox, Referenz auf die 'zusammengesucht'-Checkbox
    #              damit diese en/disabled werden kann.
    cb = QCheckBox()
    cb.setChecked(checked)

    def _on_toggled(state: bool):
        if not hasattr(app, '_noch_offen_pending'):
            app._noch_offen_pending = {}
        entry = app._noch_offen_pending.setdefault(sm_nr, {})
        entry[field] = state
        # Kopplung: vollstaendig steuert ob zusammengesucht aktivierbar ist
        if field == 'vollstaendig' and cb_zusammen is not None:
            cb_zusammen.setEnabled(state)
            if not state:
                # Haken bei zusammengesucht entfernen wenn vollstaendig deaktiviert
                cb_zusammen.setChecked(False)

    cb.toggled.connect(_on_toggled)
    cb.setStyleSheet(
        "QCheckBox { color: #228B22; background: #000000; }"
        "QCheckBox::indicator { width: 18px; height: 18px; }"
        "QCheckBox::indicator:unchecked { border: 2px solid #228B22; background: #000000; }"
        "QCheckBox::indicator:checked { border: 2px solid #FFFFFF; background: #228B22; }"
        "QCheckBox::indicator:disabled { border: 2px solid #555555; background: #222222; }"
    )
    container = QWidget()
    container.setStyleSheet("background: #000000; border-width: 0px;")
    lay = QHBoxLayout(container)
    lay.addWidget(cb)
    lay.setAlignment(cb, Qt.AlignmentFlag.AlignCenter)
    lay.setContentsMargins(0, 0, 0, 0)
    return container, cb


def show_noch_offen(app: application.App) -> None:
    unbind_all_widgets(app)
    app.disabled_button.setEnabled(True)
    app.disabled_button = app.button_noch_offen
    app.button_noch_offen.setEnabled(False)

    enabled  = [app.sm_entry]
    disabled = [app.posnr_entry, app.matnr_entry]
    hide     = [app.bestellt_button, app.bestellt_label,
                app.delete_button, app.print_button, app.combobox_loeschen]
    show     = [app.button_aenderungen_bestaetigen]
    set_widget_status(enabled, disabled, hide=hide, show=show)

    app.sm_entry.returnPressed.connect(lambda: show_noch_offen(app))
    app.sm_entry.textChanged.connect(lambda: show_noch_offen(app))
    app.sm_entry.setFocus()

    cursor     = app.cursor
    connection = app.connection
    _ensure_noch_offen_table(cursor, connection)

    smnr = app.sm_entry.text()
    all_sm_rows = cursor.execute(
        "SELECT DISTINCT SM_Nummer FROM Warenausgabe_Comline WHERE SM_Nummer LIKE ? ORDER BY SM_Nummer",
        (f'%{smnr}%',)
    ).fetchall()

    addr_rows = cursor.execute("SELECT SM_Nummer, Adresse FROM Adresszuordnung").fetchall()
    addr_dict = {row['SM_Nummer']: row['Adresse'] for row in addr_rows}

    # DB-Status zusammengesucht
    status_rows = cursor.execute(
        "SELECT * FROM Noch_Offen_Status"
    ).fetchall()
    status_dict = {row['SM_Nummer']: bool(row['zusammengesucht']) for row in status_rows}
    bem_dict    = {row['SM_Nummer']: _bemerkung(row) for row in status_rows}

    # DB-Status vollstaendig
    vollst_rows = cursor.execute(
        "SELECT SM_Nummer, vollstaendig FROM Vollstaendig_Status"
    ).fetchall()
    vollst_dict = {row['SM_Nummer']: bool(row['vollstaendig']) for row in vollst_rows}

    # Pending-Dict initialisieren
    if not hasattr(app, '_noch_offen_pending'):
        app._noch_offen_pending = {}
    # Bereits abgeschlossene Eintraege aus pending entfernen
    for sm_nr, done in status_dict.items():
        if done:
            app._noch_offen_pending.pop(sm_nr, None)

    _clear_tree(app)
    # Spalten: 0=tree, 1=Adresse, 2=SM Nummer, 3=vollständig, 4=zusammengesucht, 5=Bemerkung
    app.output_listbox.setColumnCount(7)
    app.output_listbox.headerItem().setText(0, '')
    app.output_listbox.headerItem().setText(1, 'Adresse')
    app.output_listbox.headerItem().setText(2, 'SM Nummer')
    app.output_listbox.headerItem().setText(3, 'vollständig')
    app.output_listbox.headerItem().setText(4, 'zusammengesucht')
    app.output_listbox.headerItem().setText(5, 'Bemerkung')
    app.output_listbox.headerItem().setText(6, '')
    app.output_listbox.setColumnWidth(0, 30)
    app.output_listbox.setColumnWidth(1, 620)
    app.output_listbox.setColumnWidth(2, 120)
    app.output_listbox.setColumnWidth(3, 110)
    app.output_listbox.setColumnWidth(4, 130)
    app.output_listbox.setColumnWidth(5, 400)

    def _bem_speichern(sm_nr, text):
        # sofort speichern, unabhängig von 'Änderungen bestätigen'
        app.cursor.execute(
            "INSERT INTO Noch_Offen_Status (SM_Nummer, zusammengesucht, Bemerkung) VALUES (?, 0, ?) "
            "ON CONFLICT(SM_Nummer) DO UPDATE SET Bemerkung = excluded.Bemerkung",
            (sm_nr, text))
    _bemerkung_anmelden(app, 5, _bem_speichern)

    any_visible = False
    for sm_row in all_sm_rows:
        sm_nr = sm_row['SM_Nummer']
        if status_dict.get(sm_nr, False):
            continue
        any_visible = True
        adresse = addr_dict.get(sm_nr, '')

        # Pending-Zustand ermitteln (Vorrang vor DB-Zustand)
        pending_entry   = app._noch_offen_pending.get(sm_nr, {})
        vollst_checked  = pending_entry.get('vollstaendig',    vollst_dict.get(sm_nr, False))
        zusammen_checked = pending_entry.get('zusammengesucht', False)

        item = SortItem(app.output_listbox)
        item.setText(1, adresse)
        item.setText(2, sm_nr)
        _bemerkung_setzen(item, 5, sm_nr, bem_dict.get(sm_nr, ''))

        # 'zusammengesucht'-Checkbox zuerst erstellen (wird an 'vollstaendig' übergeben)
        container_z, cb_z = _make_cb_container(app, sm_nr, 'zusammengesucht', zusammen_checked)
        cb_z.setEnabled(vollst_checked)  # nur aktivierbar wenn vollstaendig gesetzt

        # 'vollstaendig'-Checkbox mit Referenz auf cb_z
        container_v, _ = _make_cb_container(app, sm_nr, 'vollstaendig', vollst_checked, cb_zusammen=cb_z)

        app.output_listbox.setItemWidget(item, 3, container_v)
        app.output_listbox.setItemWidget(item, 4, container_z)

    if not any_visible:
        item = SortItem(app.output_listbox)
        item.setText(1, 'Alle Auftraege sind bereits zusammengesucht.')


def bestaetigen_noch_offen(app: application.App) -> None:
    # Speichert alle pending-Aenderungen in die Datenbank und aktualisiert die Ansicht.
    cursor     = app.cursor
    connection = app.connection
    _ensure_noch_offen_table(cursor, connection)

    pending = getattr(app, '_noch_offen_pending', {})
    try:
        with schreibvorgang(connection):
            for sm_nr, entry in pending.items():
                vollst     = entry.get('vollstaendig',    False)
                zusammen   = entry.get('zusammengesucht', False)

                if zusammen:
                    # Auftrag abgeschlossen: aus beiden Tabellen entfernen
                    # cursor.execute("DELETE FROM Noch_Offen_Status  WHERE SM_Nummer = ?", (sm_nr,))
                    # Upsert: auch wenn noch keine Zeile für die SM-Nummer existiert
                    # (vollständig + zusammengesucht in einem Zug angehakt)
                    cursor.execute(
                        "INSERT INTO Noch_Offen_Status (SM_Nummer, zusammengesucht) VALUES (?, 1) "
                        "ON CONFLICT(SM_Nummer) DO UPDATE SET zusammengesucht = 1",
                        (sm_nr,))
                    cursor.execute("DELETE FROM Vollstaendig_Status WHERE SM_Nummer = ?", (sm_nr,))
                else:
                    # vollstaendig-Zustand speichern
                    cursor.execute(
                        "INSERT INTO Vollstaendig_Status (SM_Nummer, vollstaendig) VALUES (?, ?) "
                        "ON CONFLICT(SM_Nummer) DO UPDATE SET vollstaendig = excluded.vollstaendig",
                        (sm_nr, 1 if vollst else 0)
                    )
                    # zusammengesucht in Noch_Offen_Status auf 0 sicherstellen
                    cursor.execute(
                        "INSERT INTO Noch_Offen_Status (SM_Nummer, zusammengesucht) VALUES (?, 0) "
                        "ON CONFLICT(SM_Nummer) DO NOTHING",
                        (sm_nr,)
                    )
    except DatenbankBelegt:
        # Häkchen bleiben als pending erhalten, Benutzer kann erneut bestätigen
        meldung_belegt()
        return

    # Pending-Dict leeren – DB ist jetzt der Stand
    app._noch_offen_pending = {}

    # Ansicht neu aufbauen
    show_noch_offen(app)


# ── Bestand anzeigen ─────────────────────────────────────────────────
def show_stock(app: application.App) -> None:
    unbind_all_widgets(app)
    app.disabled_button.setEnabled(True)
    app.disabled_button = app.button_bestand
    app.button_bestand.setEnabled(False)

    enabled  = [app.matnr_entry]
    disabled = [app.posnr_entry, app.sm_entry]
    hide     = [app.bestellt_button,
                app.bestellt_label, app.delete_button, app.print_button,
                app.combobox_loeschen]

    set_widget_status(enabled, disabled, hide = hide)
    

    app.matnr_entry.returnPressed.connect(lambda: show_stock(app))
    app.matnr_entry.textChanged.connect(lambda: show_stock(app))
    app.matnr_entry.setFocus()

    cursor = app.cursor
    connection = app.connection
    matnr  = app.matnr_entry.text()
    
    # Filter über MatNr / Stammdaten-Bezeichnung (nicht die PSL-Texte), damit Ein- und
    # Ausgang immer dieselben Materialnummern treffen und der Bestand stimmt
    bedingung, params = _material_filter(matnr, eigene_bezeichnung=False)
    ingoing  = cursor.execute(
        f"SELECT * FROM Wareneingang WHERE {bedingung}", params).fetchall()
    outgoing = cursor.execute(
        f"SELECT * FROM Warenausgang WHERE {bedingung}", params).fetchall()
    bundles = cursor.execute(
        "SELECT * FROM Bundles").fetchall()
    bundles_in = cursor.execute(
        "SELECT * FROM Wareneingang WHERE MatNr in (SELECT MatNr FROM Bundles)").fetchall()
    bundles_out = cursor.execute(
        "SELECT * FROM Warenausgang WHERE MatNr in (SELECT MatNr FROM Bundles)").fetchall()
    
    ingoing_dict  = defaultdict(int)
    outgoing_dict = defaultdict(int)
    bundles_in_dict = defaultdict(int)
    bundles_out_dict = defaultdict(int)
    bundles_total_dict = defaultdict(list)

    for row in ingoing:
        ingoing_dict[row['MatNr']] += row['Menge']
    for row in bundles_in:
        bundles_in_dict[row['MatNr']] += row['Menge']
    for row in outgoing:
        menge = row['Warenausgangsmenge'] if row['Warenausgangsmenge'] != 0 else row['Bedarfsmenge']
        if row['PosTyp'] == 8:
            outgoing_dict[row['MatNr']] -= menge
        elif row['PosTyp'] == 9:
            outgoing_dict[row['MatNr']] += menge
    for row in bundles_out:
        menge = row['Warenausgangsmenge'] if row['Warenausgangsmenge'] != 0 else row['Bedarfsmenge']
        if row['PosTyp'] == 8:
            bundles_out_dict[row['MatNr']] -= menge
        elif row['PosTyp'] == 9:
            bundles_out_dict[row['MatNr']] += menge        

    
    for row in bundles:
        menge = bundles_in_dict.get(row['MatNr'], 0) - bundles_out_dict.get(row['MatNr'], 0) + app.correction_dict.get(row['MatNr'], 0)
        bundles_total_dict[row['MatNr']] = [row['Bezeichnung'], ast.literal_eval(row['Packungsinhalt']), menge]
    
    for row in cursor.execute(
            "SELECT * FROM Warenausgang_Kleinstmaterial_ohne_SM_BEZUG").fetchall():
        outgoing_dict[row['MatNr']] += row['Menge']


    # ── Noch-offen-Mengen berechnen ────────────────────────────────────
    # Alle SM-Nummern die im Fenster 'noch offen' noch angezeigt werden
    open_sm_list = get_open_sm_numbers(cursor, connection)
    noch_offen_dict = defaultdict(int)
    if open_sm_list:
        placeholders = ','.join('?' for _ in open_sm_list)
        noch_offen_rows = cursor.execute(
            f"SELECT MatNr, Bedarfsmenge FROM Warenausgang WHERE SM_Nummer IN ({placeholders}) AND PosTyp = 9",
            open_sm_list
        ).fetchall()
        for row in noch_offen_rows:
            noch_offen_dict[row['MatNr']] += row['Bedarfsmenge']

    standardmaterial_list = []
    small_material_list   = []
    for matnr_key, menge in ingoing_dict.items():
        bezeichnung       = app.materialnames_dict[matnr_key]
        einheit           = app.units_dict[matnr_key]
        bemerkung         = ''
        bundle_correction = 0 # Das Material aus Bundles wird den Einzelpositionen hinzugerechnet
        for bundle, data in bundles_total_dict.items():
            _, bundle_content, bundle_menge = data
            if matnr_key in bundle_content:
                bundle_correction = bundle_menge
                if bundle_correction:
                    if not bemerkung:
                        bemerkung += f'davon {bundle_menge} {einheit} aus Bundle {bundle}'
                    else:
                        bemerkung += f", {bundle_menge} {einheit} aus {bundle}"
        bestand = menge - outgoing_dict.get(matnr_key, 0) + app.correction_dict.get(matnr_key, 0) + bundle_correction

        # Noch-offen-Menge in Klammern anzeigen
        noch_offen_menge = noch_offen_dict.get(matnr_key, 0)
        if noch_offen_menge:
            bestand_anzeige = f"{bestand} (+{noch_offen_menge})"
        else:
            bestand_anzeige = str(bestand)

        if matnr_key in app.standard_materials:
            standardmaterial_list.append((matnr_key, bezeichnung, bestand_anzeige, einheit, bemerkung))
        elif matnr_key in app.small_materials:
            small_material_list.append((matnr_key, bezeichnung, bestand_anzeige, einheit, bemerkung))

    columns = ['MatNr.', 'Bezeichnung', 'Bestand', 'Einheit', 'Bemerkungen', 'LAST_COLUMN']
    _clear_tree(app)
    _setup_columns(app, columns)

    if standardmaterial_list:
        parent = _insert_parent(app, 'Standardmaterial')
        for entry in sorted(standardmaterial_list, key=lambda x: int(x[0])):
            _insert_child(parent, entry, app)
    if small_material_list:
        parent = _insert_parent(app, 'Kleinstmaterial')
        for entry in sorted(small_material_list, key=lambda x: int(x[0])):
            _insert_child(parent, entry, app)
    if not standardmaterial_list and not small_material_list:
        item = SortItem(app.output_listbox)
        item.setText(2, 'Keine Daten gefunden. Bitte Filter prüfen')


# ── Wareneingang anzeigen ────────────────────────────────────────────
def show_ingoing_material(app: application.App) -> None:
    
    unbind_all_widgets(app)
    app.disabled_button.setEnabled(True)
    app.disabled_button = app.button_wareneingang
    app.button_wareneingang.setEnabled(False)

    enabled  = [app.matnr_entry]
    disabled = [app.sm_entry, app.posnr_entry, ]
    hide     = [app.bestellt_label,
                app.bestellt_button, app.delete_button, app.print_button,
                app.combobox_loeschen]

    set_widget_status(enabled, disabled, hide = hide)
    
    app.matnr_entry.setFocus()
    app.matnr_entry.returnPressed.connect(lambda: show_ingoing_material(app))
    app.matnr_entry.textChanged.connect(lambda: show_ingoing_material(app))

    cursor = app.cursor
    matnr  = app.matnr_entry.text()

    bedingung, params = _material_filter(matnr)
    selection  = cursor.execute(
        f"SELECT * FROM Wareneingang WHERE {bedingung}", params).fetchall()
    correction = cursor.execute(
        f"SELECT * FROM Jahresinventur_Korrekturdaten WHERE {bedingung} AND Menge > 0",
        params).fetchall()

    columns = ['ID', 'MatNr.', 'Bezeichnung', 'Menge', 'Einheit', 'Datum', 'Bemerkung', 'LAST_COLUMN']
    bem_col = columns.index('Bemerkung') + 1   # +1 wegen Baum-Spalte 0
    _clear_tree(app)
    _setup_columns(app, columns)

    def _bem_speichern(we_id, text):
        app.cursor.execute('UPDATE Wareneingang SET Bemerkung = ? WHERE ID = ?', (text, we_id))
    _bemerkung_anmelden(app, bem_col, _bem_speichern)

    if selection:
        parent = _insert_parent(app, 'Standardlieferung')
        for entry in reversed(selection):
            date = datetime.datetime.strftime(
                datetime.datetime.strptime(entry['Datum'], r"%Y-%m-%d %H:%M:%S"),
                r"%d.%m.%Y %H:%M:%S")
            child = _insert_child(parent, (entry['ID'], entry['MatNr'], entry['Bezeichnung'],
                                           entry['Menge'], app.units_dict[entry['MatNr']], date), app)
            _bemerkung_setzen(child, bem_col, entry['ID'], _bemerkung(entry))
    if correction:
        parent = _insert_parent(app, 'Inventurkorrektur', expanded = False)
        for entry in correction:
            date = datetime.datetime.strftime(
                datetime.datetime.strptime(entry['Datum'], r"%Y-%m-%d %H:%M:%S"),
                r"%d.%m.%Y %H:%M:%S")
            _insert_child(parent, (entry['ID'], entry['MatNr'], entry['Bezeichnung'],
                                   entry['Menge'], app.units_dict[entry['MatNr']], date), app)
    if not selection:
        item = SortItem(app.output_listbox)
        item.setText(3, 'Keine Daten gefunden. Bitte Filter prüfen.')


# ── Warenausgang anzeigen ────────────────────────────────────────────
def show_outgoing_material(app: application.App) -> None:
    unbind_all_widgets(app)
    app.disabled_button.setEnabled(True)
    app.disabled_button = app.button_warenausgang
    app.button_warenausgang.setEnabled(False)

    enabled  = [app.sm_entry, app.matnr_entry]
    disabled = [app.posnr_entry]
    hide     = [app.bestellt_label, app.bestellt_button,
                app.delete_button, app.print_button, app.combobox_loeschen]
    
    set_widget_status(enabled, disabled, hide = hide)
    

    app.sm_entry.returnPressed.connect(lambda: show_outgoing_material(app))
    app.sm_entry.textChanged.connect(lambda: show_outgoing_material(app))
    app.matnr_entry.returnPressed.connect(lambda: show_outgoing_material(app))
    app.matnr_entry.textChanged.connect(lambda: show_outgoing_material(app))
    focused = QApplication.focusWidget()
    if focused not in (app.sm_entry, app.matnr_entry):
        app.matnr_entry.setFocus()

    cursor = app.cursor
    matnr  = app.matnr_entry.text()
    smnr   = app.sm_entry.text()

    bedingung, params = _material_filter(matnr)
    selection_with_sm = cursor.execute(
        f"SELECT * FROM Warenausgang WHERE {bedingung} AND SM_Nummer LIKE ? AND PosTyp = 9",
        params + [f'%{smnr}%']).fetchall()
    selection_without_sm = cursor.execute(
        f"SELECT * FROM Warenausgang_Kleinstmaterial_ohne_SM_Bezug WHERE {bedingung}",
        params).fetchall()
    correction = cursor.execute(
        f"SELECT * FROM Jahresinventur_Korrekturdaten WHERE {bedingung} AND Menge < 0",
        params).fetchall()

    columns = ['Nr.', 'SM Nummer / ID', 'Position', 'Pos.Typ', 'MatNr.',
               'Bezeichnung', 'Menge', 'Einheit', 'Lieferschein',
               'SD Beleg', 'Materialbeleg', 'LAST_COLUMN']
    _clear_tree(app)
    _setup_columns(app, columns)
    app.output_listbox.setColumnWidth(0, 280)

    if selection_with_sm:
        selection_with_sm = sorted(selection_with_sm, key=lambda x: x['SM_Nummer'])
        parent = _insert_parent(app, 'Warenausgang mit SM Bezug')
        for number, entry in enumerate(selection_with_sm, start=1):
            _insert_child(parent, (number, entry['SM_Nummer'], entry['Position'],
                                   entry['PosTyp'], entry['MatNr'], entry['Bezeichnung'],
                                   entry['Bedarfsmenge'], app.units_dict[entry['MatNr']],
                                   entry['Lieferschein'][:-2], entry['SD_Beleg'][:-2],
                                   entry['Materialbeleg'][:-2]), app)
    if selection_without_sm and not smnr:
        parent = _insert_parent(app, 'Warenausgang ohne SM Bezug')
        for number, entry in enumerate(selection_without_sm,
                                       start=len(selection_with_sm) + 1):
            _insert_child(parent, (number, entry['ID'], '', '', entry['MatNr'],
                                   entry['Bezeichnung'], entry['Menge'],
                                   app.units_dict[entry['MatNr']]), app)
    if correction:
        parent = _insert_parent(app, 'Inventurkorrektur', expanded = False)
        start = len(selection_with_sm) + (len(selection_without_sm) if not smnr else 0)
        for number, entry in enumerate(correction, start=start + 1):
            date = datetime.datetime.strftime(
                datetime.datetime.strptime(entry['Datum'], r"%Y-%m-%d %H:%M:%S"),
                r"%d.%m.%Y %H:%M:%S")
            _insert_child(parent, (number, entry['ID'], '', '', entry['MatNr'],
                                   entry['Bezeichnung'], abs(entry['Menge']),
                                   app.units_dict[entry['MatNr']], date), app)
    if not selection_with_sm and not selection_without_sm and not correction:
        item = SortItem(app.output_listbox)
        item.setText(6, 'Keine Daten gefunden. Bitte Filter prüfen')


# ── Material für SM-Auftrag anzeigen ─────────────────────────────────
def show_material_for_order(app: application.App) -> None:
    
    unbind_all_widgets(app)
    
    app.disabled_button.setEnabled(True)
    app.disabled_button = app.button_sm_auftrag
    app.button_sm_auftrag.setEnabled(False)

    enabled  = [app.sm_entry]
    disabled = [app.posnr_entry, 
                app.matnr_entry
                ]
    show     = [app.print_button]
    hide     = [app.bestellt_label,
                app.bestellt_button, 
                app.delete_button, 
                app.combobox_loeschen
                ]
    set_widget_status(enabled, disabled, show = show, hide = hide)
    
    

    app.sm_entry.returnPressed.connect(lambda: show_material_for_order(app))
    app.sm_entry.textChanged.connect(lambda: show_material_for_order(app))
    app.sm_entry.setFocus()

    cursor = app.cursor
    smnr   = app.sm_entry.text()
    selection          = cursor.execute(
        "SELECT * FROM Warenausgabe_Comline WHERE SM_Nummer LIKE ?",
        (f'%{smnr}%',)).fetchall()
    selection_addresses= cursor.execute(
        "SELECT * FROM Adresszuordnung WHERE SM_Nummer LIKE ?",
        (f'%{smnr}%',)).fetchall()
    bundles      = cursor.execute("SELECT * FROM Bundles").fetchall()
    bundles_dict = {}
 # ---------------------- Hier bundles dict erzeugen, nicht jede Position einzeln ermitteln  
    for row in bundles:
        bundles_dict[row['MatNr']] = [row['Bezeichnung'], ast.literal_eval(row['Packungsinhalt'])]
 # --------------------------------------------------------------------------------------------   
    standard_material = []
    small_material    = []
    telekom_material  = []
    for row in selection:
        values = (row['SM_Nummer'], row['MatNr'], row['Bezeichnung'],
                  row['Bedarfsmenge'], app.units_dict[row['MatNr']] or 'n/a')
        if row['MatNr'] in app.standard_materials:
            standard_material.append(values)
        elif row['MatNr'] in app.small_materials:
            small_material.append(values)
        else:
            telekom_material.append(values)
    address_values = [(r['SM_Nummer'], r['VPSZ'], r['Adresse']) for r in selection_addresses]

    columns = ['SM Nummer', 'MatNr.', 'Bezeichnung', 'Menge', 'Einheit', 'LAST_COLUMN']
    _clear_tree(app)
    _setup_columns(app, columns)
    app.output_listbox.setColumnWidth(0, 250)

    if not selection:
        item = SortItem(app.output_listbox)
        item.setText(3, 'Keine Daten gefunden. Bitte Filter prüfen.')
        return

    if address_values:
        parent = _insert_parent(app, 'Adressen')
        for entry in address_values:
            _insert_child(parent, entry, app)

    if standard_material:
        parent = _insert_parent(app, 'Standardmaterial')
        for entry in standard_material:
            bundle_nr = entry[1]
            amount    = entry[3]
            if bundle_nr in bundles_dict:
                bundle_item = _insert_child(parent, entry, app, 'bundle_head')
                _, contents = bundles_dict[bundle_nr]
                for matnr in contents:
                    name = app.materialnames_dict[matnr]
                    unit = app.units_dict[matnr]
                    _insert_child(bundle_item, ('', matnr, name, amount, unit), app, 'bundle')
            else:
                _insert_child(parent, entry, app)

    if small_material:
        parent = _insert_parent(app, 'Kleinstmaterial')
        for entry in small_material:
            _insert_child(parent, entry, app)

    if telekom_material:
        parent = _insert_parent(app, 'Telekommaterial')
        for entry in telekom_material:
            _insert_child(parent, entry, app)


# ── Drucken ──────────────────────────────────────────────────────────
def _drucker_port(drucker: str) -> str | None:
    """Port des Druckers, wie Excel ihn erwartet (z. B. 'Ne03:'), aus der Registry."""
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r'Software\Microsoft\Windows NT\CurrentVersion\Devices') as key:
            wert, _ = winreg.QueryValueEx(key, drucker)   # z. B. 'winspool,Ne03:'
        return wert.split(',')[-1].strip()
    except OSError:
        return None


def _excel_drucker_setzen(excel_app, drucker: str) -> None:
    """Setzt den Drucker der (eigenen) Excel-Instanz.

    Excel erwartet 'Name auf Port' (deutsches Excel) bzw. 'Name on Port' (englisches).
    """
    if not drucker or excel_app.ActivePrinter.startswith(drucker + ' '):
        return  # schon aktiv (z. B. Standarddrucker)
    port = _drucker_port(drucker)
    kandidaten = [f'{drucker} {wort} {port}' for wort in ('auf', 'on')] if port else []
    kandidaten.append(drucker)
    for name in kandidaten:
        try:
            excel_app.ActivePrinter = name
            return
        except Exception:
            continue
    raise RuntimeError(f"Der Drucker '{drucker}' konnte in Excel nicht ausgewählt werden.")


def print_screen(app: application.App) -> None:
    # Temporäre Datei im TEMP-Ordner (Programmordner kann schreibgeschützt sein)
    file_path = os.path.join(tempfile.gettempdir(), 'Lagerverwaltung_Ausgabe.xlsx')
    if len(app.sm_entry.text()) < 4:
        ctypes.windll.user32.MessageBoxW(
            0,
            "Um Fehlausdrucke zu vermeiden müssen im Filter SM-Nummer \n"
            "mindestens 4 Zeichen eingetragen sein.",
            "Zu wenig Zeichen bei SM Nummer", 64)
        return
    # Druckerauswahl (Standarddrucker vorausgewählt) statt Ja/Nein-Frage
    auswahl = app.choose_printer(f"Drucken – SM {app.sm_entry.text()}")
    if auswahl is None:
        return
    drucker, exemplare = auswahl

    standard_material = []
    small_material    = []
    telekom_material  = []
    address           = []
    col_count         = app.output_listbox.columnCount()

    def _get_values(item: QTreeWidgetItem) -> list[str]:
        return [item.text(c) for c in range(col_count)]

    root = app.output_listbox.invisibleRootItem()
    for i in range(root.childCount()):
        parent     = root.child(i)
        group_text = parent.text(0)
        for j in range(parent.childCount()):
            child  = parent.child(j)
            values = _get_values(child)[1:]
            output = '\t'.join(values)
            if group_text == 'Standardmaterial':
                standard_material.append(output)
            elif group_text == 'Kleinstmaterial':
                small_material.append(output)
            elif group_text == 'Adressen':
                address.append(output)
            else:
                telekom_material.append(output)
            # Sub-Children (Bundles)
            for k in range(child.childCount()):
                sub     = child.child(k)
                subvals = _get_values(sub)[1:]
                subout  = '\t'.join(subvals)
                if group_text == 'Standardmaterial':
                    standard_material.append(subout)
                elif group_text == 'Kleinstmaterial':
                    small_material.append(subout)
                elif group_text == 'Adressen':
                    address.append(subout)
                else:
                    telekom_material.append(subout)

    wb    = xl.Workbook()
    sheet = wb.active
    sheet.page_setup.orientation = 'landscape'
    sheet.page_setup.fitToPage   = True
    sheet.column_dimensions['A'].width = 12
    sheet.column_dimensions['B'].width = 16
    sheet.column_dimensions['C'].width = 44
    sheet.column_dimensions['D'].width = 7
    sheet.column_dimensions['E'].width = 7
    sheet.column_dimensions['F'].width = 3
    for x in range(7, 37):
        letter = xl.utils.cell.get_column_letter(x)
        sheet.column_dimensions[letter].width = 1 if x % 2 == 1 else 3

    start = 1
    if address:
        sheet.cell(column=1, row=start, value='Adresse')
        sheet.cell(column=1, row=start).font = xl.styles.Font(size=12, bold=True)
        start += 1
        for row, line in enumerate(address, start=start):
            for col, word in enumerate(line.split('\t'), start=1):
                sheet.cell(column=col, row=row, value=word)
        start += len(address) + 1
    if standard_material:
        sheet.cell(column=1, row=start, value='Standardmaterial')
        sheet.cell(column=1, row=start).font = xl.styles.Font(size=12, bold=True)
        start += 1
        for row, line in enumerate(standard_material, start=start):
            for col, word in enumerate(line.split('\t'), start=1):
                sheet.cell(column=col, row=row, value=word)
        start += len(standard_material) + 1
    if small_material:
        sheet.cell(column=1, row=start, value='Kleinstmaterial')
        sheet.cell(column=1, row=start).font = xl.styles.Font(size=12, bold=True)
        start += 1
        for row, line in enumerate(small_material, start=start):
            for col, word in enumerate(line.split('\t'), start=1):
                sheet.cell(column=col, row=row, value=word)
        start += len(small_material) + 1
    if telekom_material:
        sheet.cell(column=1, row=start, value='Telekommaterial')
        sheet.cell(column=1, row=start).font = xl.styles.Font(size=12, bold=True)
        start += 1
        for row, line in enumerate(telekom_material, start=start):
            for col, word in enumerate(line.split('\t'), start=1):
                sheet.cell(column=col, row=row, value=word)

    border = xl.styles.Side(border_style='thin', color='000000')
    bundle = False
    for idx, row in enumerate(sheet, start=1):
        if not row[1].value:
            continue
        if not sheet.cell(row=idx, column=5).value:
            for col in range(1, 4):
                sheet.cell(row=idx, column=col).font = xl.styles.Font(color="3282F6")
            continue
        if row[0].value:
            start_col = 1
            bundle    = False
        else:
            start_col = 2
            for col in range(1, 6):
                if not bundle:
                    sheet.cell(row=idx - 1, column=col).fill = xl.styles.PatternFill(
                        start_color="B6B6B6", fill_type="solid")
                sheet.cell(row=idx, column=col).fill = xl.styles.PatternFill(
                    start_color="E6E6E6", fill_type="solid")
            bundle = True
        for col in range(start_col, 6):
            sheet.cell(row=idx, column=col).border = xl.styles.Border(bottom=border)
        try:
            if (int(sheet.cell(row=idx, column=4).value) > 1 and
                    sheet.cell(row=idx, column=5).value.strip() in ('M', 'n/a')):
                for col in range(1, 6):
                    sheet.cell(row=idx, column=col).font = xl.styles.Font(
                        color="FF0000", bold=True)
        except (TypeError, ValueError):
            pass
        if sheet.cell(row=idx, column=5).value.strip() in ('ST', 'SA', 'PAK'):
            amount = int(sheet.cell(row=idx, column=4).value) * 2
            if amount > 30:
                amount = 2
        else:
            amount = 2
        for col in range(amount):
            if col % 2 == 0:
                continue
            sheet.cell(row=idx, column=col + 7).border = xl.styles.Border(
                bottom=border, top=border, left=border, right=border)

    wb.save(file_path)
    # Eigene Excel-Instanz (DispatchEx): ein beim Benutzer offenes Excel wird
    # weder versteckt noch geschlossen.
    wb_com = None
    excel_app = None
    try:
        excel_app = win32com.client.DispatchEx('Excel.Application')
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        wb_com = excel_app.Workbooks.Open(file_path)
        # Drucker erst nach dem Öffnen setzen (ohne offene Mappe lehnt Excel das ab)
        _excel_drucker_setzen(excel_app, drucker)
        wb_com.PrintOut(Copies=exemplare)
    except Exception as err:
        ctypes.windll.user32.MessageBoxW(
            0, f"Drucken über Excel ist fehlgeschlagen:\n\n{err}",
            "Drucken...", 16)
    finally:
        if wb_com is not None:
            try:
                wb_com.Close(False)
            except Exception:
                pass
        if excel_app is not None:
            try:
                excel_app.Quit()
            except Exception:
                pass
        try:
            os.remove(file_path)
        except OSError:
            pass


# ── Warenausgang aus Excel buchen ────────────────────────────────────
def book_outgoing_from_excel_file(app: application.App) -> None:
    app.disabled_button.setEnabled(True)
    app.disabled_button = app.button_warenausgang_buchen

    enabled  = []
    disabled = [app.sm_entry, 
                app.posnr_entry, 
                app.matnr_entry,
                ]
    
    hide     = [app.delete_button, 
                app.print_button, 
                app.combobox_loeschen,
                app.bestellt_button, 
                app.bestellt_label]
    
    set_widget_status(enabled, disabled, hide = hide)

    cursor     = app.cursor
    connection = app.connection

    paths, _ = QFileDialog.getOpenFileNames(
        app, 'Bitte die EXPORT Datei aus PSL wählen – Mehrfachauswahl möglich...',
        '', 'Excel (*.xlsx);;Alle Dateien (*.*)')
    if not paths:
        show_critical_material(app)
        return

    all_rows = new_entries = no_booking = already_exists = 0
    belegt_dateien = []

    insert_spalten = ('(SM_Nummer, Position, PosTyp, MatNr, Bezeichnung, SD_Beleg, '
                      'Bedarfsmenge, Warenausgangsmenge, Umbuchungsmenge, '
                      'Lieferschein, Materialbeleg) VALUES (?,?,?,?,?,?,?,?,?,?,?)')

    def _insert_werte(line) -> tuple:
        entry = line.fillna(" ")
        return tuple(_py(entry[k]) for k in (
            'Auftrag', 'Stl.Position', 'Positionstyp', 'Material', 'Materialkurztext',
            'SD Beleg', 'Bedarfsmenge', 'Warenausgangsmenge', 'Umbuchungsmenge',
            'Lieferschein', 'Materialbeleg'))

    def _schon_vorhanden(tabelle: str, entry) -> bool:
        return cursor.execute(
            f"SELECT 1 FROM {tabelle} WHERE SM_Nummer = ? AND Position = ? LIMIT 1",
            (_py(entry['Auftrag']), _py(entry['Stl.Position']))).fetchone() is not None

    for file in paths:
        # Excel lesen, bevor die Datenbank gesperrt wird
        excel_file             = pd.read_excel(file)
        valid_entries          = []
        absolute_valid_entries = []
        ausgabe_entries        = []
        valid_ausgabe_entries  = []

        try:
            # Prüfen ("schon gebucht?") und Einfügen in einem Zug, damit dieselbe
            # Datei nicht gleichzeitig von zwei Personen doppelt gebucht wird.
            with schreibvorgang(connection):
                kleinstmaterial  = [r[0] for r in cursor.execute(
                    'SELECT MatNr FROM Kleinstmaterial').fetchall()]
                standardmaterial = [r[0] for r in cursor.execute(
                    'SELECT MatNr FROM Standardmaterial').fetchall()]

                for _, line in excel_file.iterrows():
                    if line['Positionstyp'] in (9, '9', 'N'):
                        ausgabe_entries.append(line)
                    if line['Material'] in kleinstmaterial:
                        valid_entries.append(line)
                        continue
                    if (line['Material'] in standardmaterial and
                            not pd.isnull(line['Lieferungen(ab/bis)'])): # or
                            #  line['Positionstyp'] in (8, '8'))):
                        valid_entries.append(line)

                for entry in valid_entries:
                    if not _schon_vorhanden('Warenausgang', entry):
                        absolute_valid_entries.append(entry)

                for entry in ausgabe_entries:
                    if not _schon_vorhanden('Warenausgabe_Comline', entry):
                        valid_ausgabe_entries.append(entry)

                for line in absolute_valid_entries:
                    cursor.execute('INSERT INTO Warenausgang ' + insert_spalten,
                                   _insert_werte(line))

                for line in valid_ausgabe_entries:
                    cursor.execute('INSERT INTO Warenausgabe_Comline ' + insert_spalten,
                                   _insert_werte(line))
        except DatenbankBelegt:
            belegt_dateien.append(os.path.basename(file))
            continue

        rows_in_excel       = excel_file.shape[0]
        valid_count          = len(valid_entries)
        absolute_valid_count = len(absolute_valid_entries)
        all_rows             += rows_in_excel
        new_entries          += absolute_valid_count
        no_booking           += rows_in_excel - valid_count
        already_exists       += valid_count - absolute_valid_count

    ctypes.windll.user32.MessageBoxW(
        0,
        f"{new_entries} von {all_rows} Datensätze in Datenbank aufgenommen.\n"
        f"{no_booking} Einträge enthielten keine Bestellung.\n"
        f"{already_exists} Einträge waren bereits in der Datenbank enthalten."
        + (f"\n\nNICHT gebucht (Datenbank war belegt, bitte erneut einlesen):\n"
           + "\n".join(belegt_dateien) if belegt_dateien else ""),
        'Ergebnis', 48 if belegt_dateien else 64)
    show_critical_material(app)


# ── Wareneingang buchen ──────────────────────────────────────────────
def book_ingoing_position(app: application.App) -> None:
    connection = app.connection
    cursor     = app.cursor

    app.disabled_button.setEnabled(True)
    app.disabled_button = app.button_wareneingang_buchen
    app.button_wareneingang_buchen.setEnabled(False)

    enabled  = []
    disabled = [app.sm_entry, 
                app.posnr_entry, 
                app.matnr_entry,
                ]
    
    hide     = [app.delete_button, 
                app.print_button, 
                app.combobox_loeschen,
                app.bestellt_button, 
                app.bestellt_label]
    
    set_widget_status(enabled, disabled, hide = hide)

    title     = 'Wareneingang buchen (Anlieferung CTDI)'
    selection = 'SELECT * FROM Kleinstmaterial UNION SELECT * FROM Standardmaterial'
    app.open_booking_window(title, selection, mit_bemerkung=True)

    if app.user_closed_window:
        show_ingoing_material(app)
        return

    try:
        matnr = re.findall(r'\d{8}', app.ingoing_mat_string_var.currentText())[0]
    except IndexError:
        show_ingoing_material(app)
        return

    try:
        menge = int(app.ingoing_menge_var.text())
    except ValueError:
        menge = 0
    if menge == 0:
        show_ingoing_material(app)
        return

    material_query = ("SELECT * FROM Standardmaterial WHERE MatNr = ? "
                      "UNION SELECT * FROM Kleinstmaterial WHERE MatNr = ?")
    cursor.execute(material_query, (matnr, matnr))
    material    = cursor.fetchall()[0]
    bezeichnung = material['Bezeichnung']
    einheit     = material['Einheit']
    bemerkung   = app.ingoing_bemerkung_var.text().strip()

    answer = ctypes.windll.user32.MessageBoxW(
        0,
        f"Soll folgendes Material gebucht werden?\n\n"
        f"{matnr}    '{bezeichnung}'       {menge} {einheit}"
        + (f"\n\nBemerkung: {bemerkung}" if bemerkung else ""),
        "Wareneingang buchen...", 68)
    if answer != 6:
        show_ingoing_material(app)
        return

    try:
        with schreibvorgang(connection):
            cursor.execute(
                'INSERT INTO Wareneingang (MatNr, Bezeichnung, Menge, Bemerkung) VALUES (?,?,?,?)',
                (matnr, bezeichnung, menge, bemerkung or None))
            # Material ist eingegangen: Bestellstatus zurücksetzen und die Bemerkung
            # aus 'Kritisches Material' leeren (gilt nur für die laufende Bestellung)
            for tabelle in ('Standardmaterial', 'Kleinstmaterial'):
                cursor.execute(
                    f'UPDATE {tabelle} SET bestellt = 0, Bemerkung = NULL WHERE MatNr = ?',
                    (matnr,))
    except DatenbankBelegt:
        meldung_belegt()
    app.button_wareneingang_buchen.setEnabled(True)
    show_ingoing_material(app)


# ── Kleinstmaterial ohne SM Bezug buchen ─────────────────────────────
def book_outgoing_kleinstmaterial(app: application.App) -> None:
    connection = app.connection
    cursor     = app.cursor

    app.disabled_button.setEnabled(True)
    app.disabled_button = app.button_warenausgang_buchen_Kleinstmaterial
    app.button_warenausgang_buchen_Kleinstmaterial.setEnabled(False)

    enabled  = []
    disabled = [app.sm_entry, 
                app.posnr_entry, 
                app.matnr_entry,
                ]
    
    hide     = [app.delete_button, 
                app.print_button, 
                app.combobox_loeschen,
                app.bestellt_button, 
                app.bestellt_label]
    
    set_widget_status(enabled, disabled, hide = hide)

    title     = 'Kleinstmaterial ohne SM Bezug herausgeben'
    selection = 'SELECT * FROM Kleinstmaterial'
    app.open_booking_window(title, selection)

    if app.user_closed_window:
        show_critical_material(app)
        return

    try:
        matnr = re.findall(r'\d{8}', app.ingoing_mat_string_var.currentText())[0]
    except IndexError:
        show_critical_material(app)
        return

    try:
        menge = int(app.ingoing_menge_var.text())
    except ValueError:
        menge = 0
    if menge == 0:
        show_critical_material(app)
        return

    cursor.execute("SELECT * FROM Kleinstmaterial WHERE MatNr = ?", (matnr,))
    material    = cursor.fetchall()[0]
    bezeichnung = material['Bezeichnung']
    einheit     = material['Einheit']

    answer = ctypes.windll.user32.MessageBoxW(
        0,
        f"Soll folgendes Material gebucht werden?\n\n"
        f"{matnr}    '{bezeichnung}'       {menge} {einheit}",
        "Warenausgang buchen...", 68)
    if answer != 6:
        show_critical_material(app)
        return

    try:
        with schreibvorgang(connection):
            cursor.execute(
                'INSERT INTO Warenausgang_Kleinstmaterial_ohne_SM_Bezug '
                '(MatNr, Bezeichnung, Menge) VALUES (?,?,?)',
                (matnr, bezeichnung, menge))
    except DatenbankBelegt:
        meldung_belegt()
    app.button_warenausgang_buchen_Kleinstmaterial.setEnabled(True)
    show_critical_material(app)


# ── Einträge aus Datenbank löschen ───────────────────────────────────
def filter_entries_to_delete(app: application.App) -> None:
    cursor = app.cursor

    # Signale der vorherigen Ansicht (und eigene vom letzten Aufruf) trennen –
    # sonst sprang das Tippen in die vorige Ansicht und die Verbindungen
    # vervielfachten sich mit jedem Tastendruck
    unbind_all_widgets(app)
    app.disabled_button.setEnabled(True)
    app.disabled_button = app.button_loeschen
    app.button_loeschen.setEnabled(False)

    if not app.combobox_loeschen.currentText():
        app.combobox_loeschen.setCurrentIndex(0)

    enabled  = [app.sm_entry, 
                app.posnr_entry, 
                app.matnr_entry
                ]
    disabled = []
    enabled.extend(app.filter_dict[app.combobox_loeschen.currentText()][0])
    disabled.extend(app.filter_dict[app.combobox_loeschen.currentText()][1])
    show = [app.delete_button, 
            app.combobox_loeschen
            ]
    hide = [app.bestellt_label, 
            app.bestellt_button, 
            app.print_button
            ]
    
    set_widget_status(enabled, disabled, show = show, hide = hide)

    app.sm_entry.textChanged.connect(lambda: filter_entries_to_delete(app))
    app.posnr_entry.textChanged.connect(lambda: filter_entries_to_delete(app))
    app.matnr_entry.textChanged.connect(lambda: filter_entries_to_delete(app))
    app.combobox_loeschen.currentIndexChanged.connect(
        lambda: filter_entries_to_delete(app))

    sm     = app.sm_entry.text()
    posnr  = app.posnr_entry.text()
    matnr  = app.matnr_entry.text()
    table  = app.combobox_loeschen.currentText()

    # Platzhalter je Filterfeld; Werte nicht per Textersetzung einsetzen
    # (Bezeichnungen enthalten Kommas, z. B. '1,5m')
    execution_string = app.execution_dict[table][0]
    werte = {'matnr': f'%{matnr.strip()}%', 'sm': f'%{sm}%', 'posnr': f'%{posnr}%'}
    execution_tuple  = tuple(werte[feld] for feld in app.execution_dict[table][1])

    columns_info = cursor.execute(f"PRAGMA table_info ({table})").fetchall()
    columns      = [info[1] for info in columns_info]
    columns.append('LAST_COLUMN')

    _clear_tree(app)
    _setup_columns(app, columns)

    selection = cursor.execute(execution_string, execution_tuple).fetchall()
    for row in selection:
        values = list(row)
        item   = SortItem(app.output_listbox)
        for i, val in enumerate(values, start=1):
            item.setText(i, str(val) if val is not None else '')


def delete_selected_entries(app: application.App) -> None:
    selections = app.output_listbox.selectedItems()
    if not selections:
        ctypes.windll.user32.MessageBoxW(
            0,
            "Es sind keine Einträge markiert.\nNur markierte Einträge werden gelöscht.",
            "Markier was..", 64)
        return
    answer = ctypes.windll.user32.MessageBoxW(
        0,
        "Achtung!!\nDie markierten Einträge werden unwiderruflich gelöscht.\n"
        "Bist du sicher, dass du die Einträge löschen willst?",
        "Einträge löschen", 68)
    if answer != 6:
        return

    col_count = app.output_listbox.columnCount()
    headers   = [app.output_listbox.headerItem().text(c) for c in range(col_count)]

    def _get_col(item: QTreeWidgetItem, name: str) -> str:
        try:
            idx = headers.index(name)
            return item.text(idx)
        except ValueError:
            return ''

    try:
        with schreibvorgang(app.connection):
            for item in selections:
                matnr = _get_col(item, 'MatNr')
                sm    = _get_col(item, 'SM_Nummer')
                posnr = _get_col(item, 'ID') or _get_col(item, 'Position')

                execution_string = app.deletion_dict[app.combobox_loeschen.currentText()][0]
                execute_values   = app.deletion_dict[app.combobox_loeschen.currentText()][1]
                execute_values   = execute_values.replace('matnr', matnr).replace('sm', sm).replace('posnr', posnr)
                execution_tuple  = tuple([e for e in execute_values.split(',') if e])
                if not execution_tuple:
                    execution_tuple = ('',)
                app.connection.execute(execution_string, execution_tuple)
    except DatenbankBelegt:
        meldung_belegt()
    filter_entries_to_delete(app)


# ── Buchungsfenster bestätigen ───────────────────────────────────────
def confirm_user_input(app: application.App) -> None:
    app.user_closed_window = False
    app.ingoing_window.accept()   # QDialog schließen (statt quit/destroy)


# ── Datenbankpflege (Aufruf von Hand aus Lagerverwaltung.py) ─────────
def add_kleinstmaterial(connection: sqlite3.Connection,
                        cursor: sqlite3.Cursor, data: tuple) -> None:
    matnr, bezeichnung, einheit, grenzwert, up_to, bestellt = data
    with schreibvorgang(connection):
        cursor.execute('''
            INSERT INTO Kleinstmaterial(MatNr,Bezeichnung,Einheit,Grenzwert,Auffüllen,bestellt)
            VALUES (?,?,?,?,?,?)''', (matnr, bezeichnung, einheit, grenzwert, up_to, bestellt))


def add_standardmaterial(connection: sqlite3.Connection,
                          cursor: sqlite3.Cursor, data: tuple) -> None:
    matnr, bezeichnung, einheit, grenzwert, up_to, bestellt = data
    with schreibvorgang(connection):
        cursor.execute('''
            INSERT INTO Standardmaterial(MatNr,Bezeichnung,Einheit,Grenzwert,Auffüllen,bestellt)
            VALUES (?,?,?,?,?,?)''', (matnr, bezeichnung, einheit, grenzwert, up_to, bestellt))


def add_bundle(connection: sqlite3.Connection, cursor: sqlite3.Cursor,
               matnr: int, bezeichnung: str, data: list) -> None:
    if not isinstance(data, list):
        print("Fehler: Datenformat muss Liste sein. z.B. [123,456]")
        return
    with schreibvorgang(connection):
        cursor.execute('''
            INSERT INTO Bundles(MatNr,Bezeichnung,Packungsinhalt) VALUES (?,?,?)''',
            (matnr, bezeichnung, str(data)))


def adressen_aus_dateien(cursor: sqlite3.Cursor, paths: list[str]) -> tuple[list, int]:
    """Liest neue Adressen (Zeilen 'OLT', ohne 'Auskundung') aus den Auftragslisten.

    Gibt (neue Adressen [(SM, VPSZ, Adresse)], Anzahl schon bekannter SM-Nummern) zurück.
    Schreibt nichts in die Datenbank.
    """
    bekannt_db = {r['SM_Nummer'] for r in cursor.execute(
        'SELECT SM_Nummer FROM Adresszuordnung').fetchall()}
    neu, gesehen, bekannt = [], set(), 0
    for file in paths:
        excel_file = pd.read_excel(file)
        start_row  = 0
        for idx, line in excel_file.iterrows():
            if line['ID'] == 'Start':
                start_row = idx
        for _, line in excel_file.iloc[start_row:].iterrows():
            if pd.isnull(line['Was']):
                continue
            was = str(line['Was'])
            if 'OLT' in was and 'Auskundung' not in was:
                sm = _nr_text(line['SM'])
                if sm in gesehen:
                    continue
                gesehen.add(sm)
                if sm in bekannt_db:
                    bekannt += 1
                    continue
                vpsz = '' if pd.isnull(line['VPSZ']) else str(_py(line['VPSZ']))
                ort  = '' if pd.isnull(line['ORT']) else str(_py(line['ORT']))
                neu.append((sm, vpsz, ort))
    return neu, bekannt


def adressen_speichern(connection: sqlite3.Connection, adressen: list) -> int:
    """Speichert Adressen; inzwischen (von anderen) angelegte SM-Nummern werden übersprungen."""
    with schreibvorgang(connection):
        vorher = connection.total_changes
        for smnr, vpsz, address in adressen:
            connection.execute(
                'INSERT OR IGNORE INTO Adresszuordnung (SM_Nummer, VPSZ, Adresse) VALUES (?,?,?)',
                (smnr, vpsz, address))
        return connection.total_changes - vorher


def read_adresses_from_workorder_list(connection: sqlite3.Connection,
                                      cursor: sqlite3.Cursor) -> None:
    """Konsolenvariante (ohne Maske), Aufruf von Hand aus Lagerverwaltung.py."""
    paths, _ = QFileDialog.getOpenFileNames(
        None, 'Excel-Datei auswählen', 'ExcelPYTHON.xlsm',
        'Excel (*.xlsx *.xlsm);;Alle Dateien (*.*)')
    if not paths:
        return
    adressen, _ = adressen_aus_dateien(cursor, paths)
    for smnr, vpsz, address in adressen:
        print(smnr, vpsz, address)
    print(f"Fertig. {adressen_speichern(connection, adressen)} neue Adressen aufgenommen.")


def write_correction_data_from_yearly_inspection(
        connection: sqlite3.Connection, cursor: sqlite3.Cursor) -> None:
    paths, _ = QFileDialog.getOpenFileNames(
        None, 'Inventurdaten auswählen', 'Inventurdaten.xlsx',
        'Excel (*.xlsx);;Alle Dateien (*.*)')
    if not paths:
        return
    tabellen = [pd.read_excel(file) for file in paths]
    with schreibvorgang(connection):
        for excel_file in tabellen:
            for _, line in excel_file.iterrows():
                cursor.execute('''
                    INSERT INTO Jahresinventur_Korrekturdaten (MatNr, Bezeichnung, Menge)
                    VALUES (?,?,?)''',
                    (_py(line['MatNr']), _py(line['Bezeichnung']), _py(line['Differenz'])))
