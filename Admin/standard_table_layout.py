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
bestellt BOOLEAN DEFAULT 0 NOT NULL CHECK (bestellt IN (0, 1))
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
'''
