import sqlite3
import functions as fc
import application


# set user privileges
# STANDARD = database read only
# EXPERT = database read / write
# ADMIN = database read / write / delete

user = application.EXPERT

if __name__ == '__main__':
    # connect to database
    path_to_db: str = fc.open_db()
    connection: sqlite3.Connection = sqlite3.connect(path_to_db)
    connection.row_factory = sqlite3.Row  # set the standard return format
    cursor: sqlite3.Cursor = connection.cursor()

    # open the frontend
    app = application.App(connection, cursor, user, path_to_db)
    app.window.state("zoomed")
    app.window.mainloop()
    connection.close()

