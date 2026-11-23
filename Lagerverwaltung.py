import sqlite3
import sys

from PyQt6.QtWidgets import QApplication

import application
import functions as fc

# set user privileges
# STANDARD = database read only
# EXPERT = database read / write
# ADMIN = database read / write / delete

user = application.EXPERT
# user = application.ADMIN

if __name__ == '__main__':
    # QApplication muss vor allen Widgets erstellt werden
    qt_app = QApplication(sys.argv)

    # connect to database
    path_to_db: str = fc.open_db()
    if not path_to_db:
        sys.exit(0)  # Auswahl der Datenbank abgebrochen
    connection: sqlite3.Connection = fc.connect_db(path_to_db)
    cursor: sqlite3.Cursor = connection.cursor()

    # open the frontend (maximiert statt "zoomed")
    app = application.App(connection, cursor, user, path_to_db)
    app.showMaximized()

    exit_code = qt_app.exec()

    # fc.read_adresses_from_workorder_list(connection, cursor)
    # fc.add_standardmaterial(connection, cursor, (47194953,'Gf-Atk 1x12F SC/APC9 1,5m SC/APC9 4m 45m', 'ST', 0, 1, False))
    # fc.add_kleinstmaterial(connection, cursor,(40205824,'E&MMS-T EMK-Modul 8Rasteinheit 8Spleißk', 'ST', 2, 20, False))
    # fc.add_bundle(connection, cursor, 40935979,'PE E&MMS-T NGN-Spleiß-Patch-Set-96 li 8m', [40821750, 40821754])
    # fc.write_correction_data_from_yearly_inspection(connection, cursor)

    connection.close()
    sys.exit(exit_code)