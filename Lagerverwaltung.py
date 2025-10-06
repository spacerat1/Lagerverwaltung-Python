import sqlite3
import functions as fc
import application


# set user privileges
# STANDARD = database read only
# EXPERT = database read / write
# ADMIN = database read / write / delete

user = application.EXPERT
#user = application.ADMIN

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
    #fc.read_adresses_from_workorder_list(connection, cursor)
    #fc.add_standardmaterial(connection, cursor, (47182895,'E&MMS-L Satz EMK Kassette V 2023 (LIC)', 'ST', 2, 10, False))
    #fc.add_kleinstmaterial(connection, cursor,(40205817,'E&MMS-Muffe-ZZ', 'ST', 2, 10, False))
    #fc.add_kleinstmaterial(connection, cursor,(40256524,'E&MMS HVt Einzel-KTU', 'PAK', 2, 10, False))
    #fc.write_correction_data_from_yearly_inspection(connection, cursor)
    connection.close()
    