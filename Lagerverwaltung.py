import sys
import sqlite3
import functions as fc
import application
from PyQt6.QtWidgets import QApplication

# set user privileges
# STANDARD = database read only
# EXPERT = database read / write
# ADMIN = database read / write / delete

user = application.EXPERT
user = application.ADMIN

if __name__ == '__main__':
    # QApplication muss vor allen Widgets erstellt werden
    qt_app = QApplication(sys.argv)

    # connect to database
    path_to_db: str = fc.open_db()
    connection: sqlite3.Connection = sqlite3.connect(path_to_db)
    connection.row_factory = sqlite3.Row  # set the standard return format
    cursor: sqlite3.Cursor = connection.cursor()

    # open the frontend (maximiert statt "zoomed")
    app = application.App(connection, cursor, user, path_to_db)
    app.showMaximized()

    exit_code = qt_app.exec()

    # fc.read_adresses_from_workorder_list(connection, cursor)
    # fc.add_standardmaterial(connection, cursor, (40980058,'E&MMS-ZCH EMK Kassette für 96 S/P BG', 'SA', 2, 10, False))
    # fc.add_kleinstmaterial(connection, cursor,(40205824,'E&MMS-T EMK-Modul 8Rasteinheit 8Spleißk', 'ST', 2, 20, False))
    # fc.add_bundle(connection, cursor, 40935979,'PE E&MMS-T NGN-Spleiß-Patch-Set-96 li 8m', [40821750, 40821754])
    # fc.write_correction_data_from_yearly_inspection(connection, cursor)

    connection.close()
    sys.exit(exit_code)