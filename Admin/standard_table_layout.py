tables = '''
CREATE TABLE iF NOT EXISTS Standardmaterial (
MatNr INTEGER PRIMARY KEY,
Bezeichnung TEXT,
Einheit TEXT,
Grenzwert INTEGER,
Auffüllen INTEGER,
bestellt BOOLEAN DEFAULT 0 NOT NULL CHECK (bestellt IN (0, 1))
);

CREATE TABLE IF NOT EXISTS Kleinstmaterial(
MatNr INTEGER PRIMARY KEY,
Bezeichnung TEXT,
Einheit TEXT,
Grenzwert INTEGER,
Auffüllen INTEGER,
bestellt BOOLEAN DEFAULT 0 NOT NULL CHECK (bestellt IN (0, 1)),
Datum DATETIME
);

CREATE TABLE IF NOT EXISTS Wareneingang(
ID INTEGER PRIMARY KEY AUTOINCREMENT,
MatNr INTEGER,
Bezeichnung TEXT,
Menge INTEGER,
Datum DATETIME NOT NULL DEFAULT (datetime(CURRENT_TIMESTAMP, 'localtime'))
); 

CREATE TABLE IF NOT EXISTS Warenausgang(
SM_Nummer TEXT,
Position TEXT,
PosTyp INTEGER,
MatNr INTEGER,
Bezeichnung TEXT,
SD_Beleg TEXT,
Bedarfsmenge INTEGER,
Warenausgangsmenge INTEGER,
Umbuchungsmenge INTEGER,
Lieferschein TEXT,
Materialbeleg TEXT
);

CREATE TABLE IF NOT EXISTS Warenausgabe_Comline(
SM_Nummer TEXT,
Position TEXT,
PosTyp INTEGER,
MatNr INTEGER,
Bezeichnung TEXT,
SD_Beleg TEXT,
Bedarfsmenge INTEGER,
Warenausgangsmenge INTEGER,
Umbuchungsmenge INTEGER,
Lieferschein TEXT,
Materialbeleg TEXT
);

CREATE TABLE IF NOT EXISTS Warenausgang_Kleinstmaterial_ohne_SM_Bezug(
ID INTEGER PRIMARY KEY AUTOINCREMENT,
MatNr INTEGER,
Bezeichnung TEXT,
Menge INTEGER,
Datum DATETIME NOT NULL DEFAULT (datetime(CURRENT_TIMESTAMP, 'localtime'))
); 

CREATE TABLE IF NOT EXISTS Adresszuordnung(
SM_Nummer TEXT PRIMARY KEY,
VPSZ TEXT,
Adresse TEXT
);

CREATE TABLE IF NOT EXISTS Veraltetes_Material(
ID INTEGER PRIMARY KEY AUTOINCREMENT,
MatNr INTEGER,
Bezeichnung TEXT,
Datum DATETIME NOT NULL DEFAULT (datetime(CURRENT_TIMESTAMP, 'localtime'))
);

CREATE TABLE IF NOT EXISTS Jahresinventur_Korrekturdaten(
ID INTEGER PRIMARY KEY AUTOINCREMENT,
MatNr INTEGER,
Bezeichnung TEXT,
Menge INTEGER,
Datum DATETIME NOT NULL DEFAULT (datetime(CURRENT_TIMESTAMP, 'localtime'))
); 

'''
