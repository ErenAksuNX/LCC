import Geometries
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as filedialog
import start
import tkinter.messagebox as messagebox


class GUI:
    def __init__(self, master: tk.Tk):
        # Die Klassenobjekte werden initialisiert
        self.master: tk.Tk = master
        self.QumaPath = ""
        self.RechnungPath = ""
        self.Speicherort = ""
        #  Hier wird der GUI das Layout gegeben
        Geometries.geometries(self.master)

        # Es werden die Buttons erstellt
        self.buttons()

    def buttons(self):
        # Die für die verwendung nötige Buttons werden erstellt
        bt_select_quma = ttk.Button(text="Select Quma Matrix", master=self.master, command=self.select_Quma)
        bt_select_rechnung = ttk.Button(text="Select Rechnungsdatei", master=self.master, command=self.select_Rechnung)
        bt_select_save = ttk.Button(text="Select Save", master=self.master, command=self.select_save)
        bt_start = ttk.Button(text="Start", master=self.master, command=self.start)

        # Es werden die Buttons in der Funktion geplaced
        Geometries.place_Buttons([bt_select_quma, bt_select_rechnung, bt_select_save, bt_start])

    def select_Quma(self):
        # Der Pfad zur Quma Matrix wird gespeichert
        self.QumaPath = filedialog.askopenfilename(title="Bitte wähle die Quma Matrix aus")

    def select_Rechnung(self):
        # Der Pfad zu rechnungsdatei wird gespeichert
        self.RechnungPath = filedialog.askopenfilename(title="Bitte Wähle die Rechnungsdatei aus")

    def select_save(self):
        # Der Speicherort wird gespeichert
        self.Speicherort = filedialog.askdirectory(title="Bitte wähle den gewünschten Speicherort aus ")

    def start(self):
        if self.QumaPath == "":
            self.QumaPath = 'C:/Users/eaksu/OneDrive - National Express Rail GmbH/Desktop/Projekte/Python Projekte/LCC/quma.xlsx'
        if self.RechnungPath == "":
            self.RechnungPath = 'C:/Users/eaksu/OneDrive - National Express Rail GmbH/Desktop/Projekte/Python Projekte/LCC/Rechnung.xlsx'
        start.run(self.QumaPath, self.RechnungPath, self.Speicherort)

