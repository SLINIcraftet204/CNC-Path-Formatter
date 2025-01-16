import os
import csv
from datetime import datetime
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


def get_desktop_path():
    """
    Gibt den Desktop-Pfad des aktuellen Benutzers zurück.
    """
    return os.path.join(os.path.expanduser("~"), "Desktop")


import re

def replace_objfullpath(content, new_path):
    """
    Ersetzt den Inhalt des <ObjFullPath>-Tags mit dem neuen Pfad.

    :param content: Inhalt der Datei als String.
    :param new_path: Neuer Pfad, der in <ObjFullPath> gesetzt wird.
    :return: Der aktualisierte Inhalt der Datei.
    """
    pattern = r"<ObjFullPath>(.*?)</ObjFullPath>"
    updated_content = re.sub(pattern, f"<ObjFullPath>{new_path}</ObjFullPath>", content)
    return updated_content



def search_and_edit_files(base_path, file_types, new_path, log_entries, progress_var, progress_bar, total_files):
    """
    Sucht nach Dateien mit bestimmten Typen und bearbeitet die Zielvariable.

    :param base_path: Basisverzeichnis, das durchsucht wird.
    :param file_types: Liste der zu suchenden Dateiendungen.
    :param new_path: Neuer Pfad, der in der Variable eingesetzt wird.
    :param log_entries: Liste, um Log-Einträge für die CSV-Datei zu speichern.
    :param progress_var: Variable, die den Fortschritt aktualisiert.
    :param progress_bar: Fortschrittsanzeige.
    :param total_files: Gesamtanzahl der Dateien für die Fortschrittsberechnung.
    """
    current_file_count = 0

    for root, _, files in os.walk(base_path):
        for file in files:
            if any(file.endswith(ext) for ext in file_types):
                file_path = os.path.join(root, file)
                variable_found = False
                updated = False

                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()

                    # Suche nach gültigen Tags und ersetze deren Inhalt
                    matches = replace_objfullpath(content)
                    if matches:
                        variable_found = True
                        updated_content = replace_objfullpath(content, new_path)

                        # Datei aktualisieren
                        with open(file_path, 'w', encoding='utf-8') as f:
                            f.write(updated_content)
                        updated = True


                except Exception as e:
                    print(f"Fehler beim Verarbeiten von {file_path}: {e}")

                # Log-Eintrag erstellen
                log_entries.append({
                    "Dateipfad": file_path,
                    "Variable Gefunden": "Ja" if variable_found else "Nein",
                    "Bearbeitet": "Ja" if updated else "Nein"
                })

                # Fortschritt aktualisieren
                current_file_count += 1
                progress_var.set((current_file_count / total_files) * 100)
                progress_bar.update()

    messagebox.showinfo("Fertig", "Die Verarbeitung ist abgeschlossen.")


def write_log_to_csv(log_entries):
    """
    Schreibt die Log-Einträge in eine CSV-Datei auf dem Desktop.

    :param log_entries: Liste mit Log-Einträgen.
    """
    desktop_path = get_desktop_path()
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    csv_file_path = os.path.join(desktop_path, f"CNC_File_Log_{timestamp}.csv")

    try:
        with open(csv_file_path, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ["Dateipfad", "Variable Gefunden", "Bearbeitet"]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(log_entries)

        print(f"Log-Datei erstellt: {csv_file_path}")
        return csv_file_path
    except Exception as e:
        print(f"Fehler beim Schreiben der Log-Datei: {e}")
        return None


def gui_mode():
    """
    GUI-Modus des Programms mit Fortschrittsanzeige und zwei Schritten.
    """
    def validate_base_path():
        base_path = entry_base_path.get().strip()
        if not os.path.isdir(base_path):
            return messagebox.showwarning("Warnung", "Ungültiges Verzeichnis ausgewählt!")
        # Wenn Verzeichnis gültig ist, zur nächsten Seite wechseln
        app_data["base_path"] = base_path
        show_new_path_page()

    def process_files():
        new_path = entry_new_path.get().strip()
        if not new_path:
            return messagebox.showwarning("Warnung", "Kein neuer Pfad eingegeben!")

        app_data["new_path"] = new_path
        file_types = ['.hop', '.ganx']
        log_entries = []

        # Dateien zählen
        total_files = sum(len(files) for _, _, files in os.walk(app_data["base_path"]) if any(f.endswith(tuple(file_types)) for f in files))
        if total_files == 0:
            return messagebox.showinfo("Keine Dateien", "Keine passenden Dateien gefunden!")

        # Fortschrittsanzeige initialisieren
        progress_var.set(0)
        progress_bar.pack()

        # Suche und Bearbeitung starten
        search_and_edit_files(app_data["base_path"], file_types, new_path, log_entries, progress_var, progress_bar, total_files)

        # Log in CSV schreiben
        log_file = write_log_to_csv(log_entries)
        if log_file:
            messagebox.showinfo("Fertig", f"Der Prozess wurde abgeschlossen.\nLog-Datei: {log_file}")
        else:
            messagebox.showwarning("Fehler", "Es gab ein Problem beim Schreiben der Log-Datei.")

    def show_new_path_page():
        # Basisverzeichnis-Seite ausblenden
        frame_base_path.pack_forget()

        # Neue Pfad-Seite anzeigen
        frame_new_path.pack()

    # App-Daten (für Pfade speichern)
    app_data = {}

    # GUI erstellen
    root = tk.Tk()
    root.title("CNC-Dateien Prüfer und Bearbeiter")

    # Fortschrittsvariable nach dem Root-Fenster initialisieren
    progress_var = tk.DoubleVar()

    # Seite 1: Basisverzeichnis
    frame_base_path = tk.Frame(root)
    tk.Label(frame_base_path, text="Pfad des Basisverzeichnisses:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    entry_base_path = tk.Entry(frame_base_path, width=50)
    entry_base_path.grid(row=0, column=1, padx=10, pady=5)
    tk.Button(frame_base_path, text="Verzeichnis wählen", command=lambda: entry_base_path.insert(0, filedialog.askdirectory(title="Wählen Sie das Basisverzeichnis"))).grid(row=0, column=2, padx=10, pady=5)
    tk.Button(frame_base_path, text="Weiter", command=validate_base_path).grid(row=1, column=1, pady=20)
    frame_base_path.pack()

    # Seite 2: Neuer Pfad
    frame_new_path = tk.Frame(root)
    tk.Label(frame_new_path, text="Neuer CNC-Objekt-Pfad:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    entry_new_path = tk.Entry(frame_new_path, width=50)
    entry_new_path.grid(row=0, column=1, padx=10, pady=5)
    tk.Button(frame_new_path, text="Starten", command=process_files).grid(row=1, column=1, pady=20)

    # Fortschrittsanzeige
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, length=400)
    progress_bar.pack_forget()

    root.mainloop()


if __name__ == "__main__":
    gui_mode()
