import sqlite3
import functions as fc
from application import App


# Nutzerrechhte vergeben 
# STANDARD = nur Leserechte
# EXPERT = Lese / Schreibrechte
# ADMIN = Lese/Schreibrechte und Daten löschen
user = fc.EXPERT

# Verbindung zur Datenbank aufbauen
path_to_db= fc.open_db()
connection = sqlite3.connect(path_to_db)
connection.row_factory = sqlite3.Row
cursor = connection.cursor()
app = App(connection, cursor, user, path_to_db)
app.window.state('zoomed')
app.window.mainloop()
#fc.add_standardmaterial(connection, cursor,(47203868, 'Modulträger ETSI 6x Modul 16x GfK 1:2sym', 'ST', 5, 20, 1))
#fc.add_kleinstmaterial(connection, cursor,(40945101, 'nVent-Kontaktscheibe 6mm 100St.', 'ST', 0, 1, 0))

