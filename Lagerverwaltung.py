import sqlite3
import functions as fc
import application


# Nutzerrechhte vergeben 
# STANDARD = nur Leserechte
# EXPERT = Lese / Schreibrechte
# ADMIN = Lese/Schreibrechte und Daten löschen

user = application.ADMIN

# Verbindung zur Datenbank aufbauen
path_to_db= fc.open_db()
connection = sqlite3.connect(path_to_db)
connection.row_factory = sqlite3.Row
cursor = connection.cursor()
app = application.App(connection, cursor, user, path_to_db)
app.window.state('zoomed')
app.window.mainloop()
#fc.add_standardmaterial(connection, cursor,(47203868, 'Modulträger ETSI 6x Modul 16x GfK 1:2sym', 'ST', 5, 20, 1))
#fc.add_standardmaterial(connection, cursor,(40748416, 'nVent Kabelöse Stahl, 100 x 100', 'ST', 5, 20, 0))
#fc.add_standardmaterial(connection, cursor,(40945101, 'nVent-Kontaktscheibe 6mm 100St.', 'ST', 0, 1, 0))
connection.close()

