import PySimpleGUI as pg
import os
from docxtpl import DocxTemplate
from pathlib import Path
import sqlite3 as sql


def is_valid_path(filepath):
    if filepath and Path(filepath).exists():
        return True
    pg.popup_error("Dateipfade nicht korrekt")
    return False


def validate(values):
    is_valid=True
    if len(values) != 0:
        pass
    else:
        is_valid = False
    return is_valid


def make_dictionary_from_db():
    # connection to the database
    database = 'database.db'
    conn = sql.connect(database)
    query = 'select * from data'

    # taking the dictionaries from the database
    conn.row_factory = sql.Row
    data = conn.execute(query).fetchall()
    conn.close()

    # Saving the dictionaries in a list
    dictionaries = [{k: item[k] for k in item.keys()} for item in data]
    keys_document = document_keys(dictionaries)
    context = dictionaries[0]
    return keys_document, context


def delete_column(name):
    try:
        db_path = 'database.db'
        conn = sql.connect(db_path)
        conn.execute("ALTER TABLE data DROP COLUMN " + name + "")
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Exception: {e}")
        return


def new_keys(name):
    db_path = 'database.db'
    conn = sql.connect(db_path)
    conn.execute("ALTER TABLE data ADD " + name + " text")
    conn.commit()
    conn.close()
    return True


def document_keys(context):
    keys_document = []
    for x in context[0]:
        keys_document.append(x)  # creating a list with all de keys to set the buttons
        
    try:
        with open("Eigenshaft.txt", "w+") as text:
            for key in keys_document:
                x = "{{"
                y= "}}"
                print(f"{x} {key} {y}", file=text)
            content=text.read()
        print(content)
    except:
        print("There was a problem reading the keys in the Word Document")
    return keys_document


def add_comment(comentario):
    db_path = 'database.db'
    conn = sql.connect(db_path)
    conn.execute("""INSERT INTO comments (sugestion) VALUES (?)""", (comentario,))
    conn.commit()
    conn.close()
    return True


def main_modify(context, project_folder, template_folder):
    """
    Generate new documents from templates with provided context data.

    Args:
    - context (dict): Data to be used for replacing placeholders in the template.
    - project_folder (str): Path to the folder where the generated documents will be saved.
    - template_folder (str): Path to the folder containing the DOCX template files.

    Returns:
    - List[Path]: A list of Paths to the generated documents.
    """

    try:
        project_path = Path(project_folder).resolve()
        template_path = Path(template_folder).resolve()
    except FileNotFoundError as e:
        print(f"Error in finding folder: {e}")
        return []

    if not project_path.is_dir() or not template_path.is_dir():
        print("Invalid project or template folder path.")
        return []

    generated_docs = []
    for template_file in template_path.glob('*.docx'):
        try:
            doc = DocxTemplate(template_file)
            output_file = project_path / f"{template_file.stem}_result.docx"
            doc.render(context)
            doc.save(output_file)
            generated_docs.append(output_file)
        except Exception as e:
            print(f"Error processing file {template_file}: {e}")
    
    return generated_docs


# ---------- GUI DEFINITION ------------- #
def create_new_key_window():
    """
    Creates a new GUI window for adding a new key (property).

    Returns:
    - bool: Indicates whether a new key was successfully added.
    """
    key_added = False

    # Define the layout of the window
    layout_new_key = [
        [pg.Text("Die neue Eigenschaft darf keine Leerzeichen haben")],
        [pg.Text("Name der neuen Eigenschaft: "), pg.Push(), pg.InputText(do_not_clear=False, key="-KEY_NAME-")],
        [pg.Exit(s=16, button_color="tomato"), pg.B("Neue Eigenschaft speichern.", s=20)]
    ]

    # Create and display the window
    window_new_key = pg.Window("Neue Eigenschaft.", layout_new_key, modal=True)

    # Event loop for handling user interactions
    while True:
        event, values = window_new_key.read()

        # Close the window on 'Exit' event or if window is closed
        if event in ("Exit", pg.WIN_CLOSED):
            break

        # Handle the event of saving a new key
        elif event == "Neue Eigenschaft speichern.":
            new_name = values["-KEY_NAME-"]
            if new_keys(new_name):  # Assuming new_keys is a function that adds a new key to the database
                pg.popup("Neue Eigenschaft erstellt :)")
                key_added = True

    # Close the window once done
    window_new_key.close()
    return key_added


def delete_key_window():
    cambio = False
    keys_document, context = make_dictionary_from_db()
    column = [[pg.Checkbox(key, default=False, key=key)] for key in keys_document]
    layout_delete_key = [
        [pg.Text("Wählen Sie die Eigenschaft aus, die Sie löschen möchten:")],
        [pg.Column(column)],
        [pg.Exit(s=16, button_color="tomato"), pg.B("Löschen", s=20)]
         ]

    window_delete_key = pg.Window("Delete key", layout_delete_key, modal=True)
    while True:  # Event loop
        event, values = window_delete_key.read()  # El programa para aquí para que el usuario ponga sus respuestas

        if event == "Exit" or event == pg.WIN_CLOSED:
            break

        elif event == "Löschen":
            for key in keys_document:
                if values[key]:
                    click = pg.popup_ok_cancel(f"Sind Sie sicher, dass Sie <{key}> löschen wollen?")
                    if click == 'OK':
                        delete_column(key)
                        pg.popup(f"Eigenschaft <{key}> gelöscht :(")
                        cambio = True

    window_delete_key.close()  # Close window
    return cambio


def logo_kunden_window():
    working_directory = os.getcwd()
    layout = [
        [pg.Text('Wählen Sie das Logo des Kunden:')],
        [pg.InputText(key="-LOGO_PATH-"),
         pg.FileBrowse(initial_folder=working_directory, file_types=[("JPG", "*.jpg"), ("PNG", "*.png")])],
        [pg.Submit("OK"), pg.Exit()]
    ]

    window = pg.Window('Logo Kunden', layout)
    logo_path = ""

    while True:
        event, values = window.read()
        if event == "Exit" or event == pg.WIN_CLOSED:  # exit from the app
            window.close()
            break

        elif event == 'OK':
            logo_path = values["-LOGO_PATH-"]
            print(logo_path)
            window.close()
            return True, logo_path
    return False, logo_path


def vorschlage_window():
    layout = [
        [pg.Text('Was soll das Programm Ihrer Meinung nach tun?: ')],
        [pg.Multiline(key="-COMMENT-", size=(50, 10), do_not_clear=False)],
        [pg.Submit("OK"), pg.Exit()]
    ]

    window = pg.Window('Vorschläge', layout)

    while True:
        event, values = window.read()
        if event == "Exit" or event == pg.WIN_CLOSED:  # exit from the app
            window.close()
            break

        elif event == 'OK':
            comment = values["-COMMENT-"]
            add_comment(comment)
            return True
    return


def make_main_window():  
    pg.theme("default1")  
    keys_document, context = make_dictionary_from_db()

    folders = [[
            pg.Text("Vorlagenordner"),
            pg.In(size=(100, 1), enable_events=True, key="-TEMPLATE_FOLDER-"),
            pg.FolderBrowse()], 
        [pg.Text("Projekt-Ordner"), pg.In(size=(100, 1), enable_events=True, key="-PROJECT_FOLDER-"),
         pg.FolderBrowse()]]  # choosing the project folder where you want the new files

    file_list_column = [  
        [
            pg.Text("Dokumente im Vorlagenordner, die geändert werden sollen: ")
        ],
        [
            pg.Listbox(  
                values=[],
                enable_events=True,
                size=(50, 20),
                key="-FILE_LIST-",
                select_mode="multiple"  # the view allows to pick several files
            )
        ]
    ]

    file_viewer_column = [[pg.Text(keys), pg.Push(), pg.InputText(do_not_clear=True, key=keys)] for keys in keys_document]
                    # input of all keys that exists in the dictionary

    instrucciones = [
        [pg.Text("1. Wählen Sie den Ordner, in dem sich die zu ändernden Vorlagen befinden.")],
        [pg.Text("2. Wählen Sie den Ordner, in dem Sie die erstellten Dokumente speichern möchten.")],
        [pg.Text("3. Geben Sie die Daten der neuen Dokumente ein. ")],
        [pg.Text("   3.1 Im Dokument muss die Eigenschaft zwischen {{ }} stehen. Beispiel: {{ Kunden_nr }}")],
        [pg.Text("   3.2 Jede Eigenschaft kann nur eine Zeile haben.")],
        [pg.Text("4. Klicken Sie auf die Schaltfläche Dokumente erstellen. ")],
    ]

    botones = [pg.Exit(s=16, button_color="tomato"),  # button of exit
            pg.B("Eigenschaft löschen", s=20),  # TASK: button to access to previous projects data
            pg.B("Neue Eigenschaft", s=20),  # button for creating the documents
            pg.Button("Dokumente erstellen.", s=20),
            pg.B("Vorschläge", s=20),
               ]


    layout = [  # putting both columns together
        [instrucciones],
        [folders],
        [pg.Column(file_list_column), pg.VSeparator(), pg.Column(file_viewer_column, size=(500, 300), scrollable=True, vertical_scroll_only=True)],
        [botones]
    ]

    window = pg.Window("Erstellung von document-Dokumenten", layout)  # Create Window
    return window


# ---------- BODY OF THE PROGRAM ------------- #

def main_window():
    keys_document, context = make_dictionary_from_db()
    folder_location = ""
    project_location = ""

    window = make_main_window()     # We make the window here

    while True:  # Event loop
        event, values = window.read()  
        if event == "Exit" or event == pg.WIN_CLOSED:  # exit from the app
            break

        elif event == "-TEMPLATE_FOLDER-":  # if you click to browse the template
            folder_location = values["-TEMPLATE_FOLDER-"]
            try:
                files = os.listdir(folder_location)  # List of files in the chosen folder
            except:
                files = []

            file_names = [
                file for file in files
                if os.path.isfile(os.path.join(folder_location, file))
                   and file.lower().endswith(".docx")
            ]
            window["-FILE_LIST-"].update(file_names)  # show the documents in the program

        elif event == "-PROJECT_FOLDER-":
            project_location = values["-PROJECT_FOLDER-"]

        elif event == "Eigenschaft löschen":
            cambio = delete_key_window()
            if cambio:
                keys_document, context = make_dictionary_from_db()
                window.close()
                window = make_main_window()

        elif event == "Neue Eigenschaft":
            cambio = create_new_key_window()
            if cambio:
                keys_document, context = make_dictionary_from_db()
                window.close()
                window = make_main_window()

        elif event == "Vorschläge":
            vorschlage_window()

        elif event == "Dokumente erstellen.":
            if (is_valid_path(folder_location)) and (is_valid_path(project_location)):
                for key in keys_document:
                    if validate(key):
                        context[key] = values[key]
                main_modify(context, project_location, folder_location)
                path = os.path.realpath(project_location)
                os.startfile(path)

    window.close()  


if __name__ == "__main__":
    main_window()
