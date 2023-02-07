import Geometries
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as filedialog
import start


class GUI:
    def __init__(self, master: tk.Tk):
        self.master: tk.Tk = master
        self.QumaPath = ""
        self.RechnungPath = ""
        self.Speicherort = ""

        #  Hier wird der GUI das Layout gegeben
        Geometries.geometries(self.master)

        self.buttons()

    def buttons(self):
        bt_select_quma = ttk.Button(text="Select Quma Matrix", master=self.master, command=self.select_Quma)
        bt_select_rechnung = ttk.Button(text="Select Rechnungsdatei", master=self.master, command=self.select_Rechnung)
        bt_select_save = ttk.Button(text="Select Save", master=self.master, command=self.select_save)
        bt_start = ttk.Button(text="Start", master=self.master, command=self.start)

        Geometries.place_Buttons([bt_select_quma, bt_select_rechnung, bt_select_save, bt_start])

    def select_Quma(self):
        self.QumaPath = filedialog.askopenfilename(title="Bitte wähle die Quma Matrix aus")

    def select_Rechnung(self):
        self.RechnungPath = filedialog.askopenfilename(title="Bitte Wähle die Rechnungsdatei aus")

    def select_save(self):
        self.Speicherort = filedialog.askdirectory(title="Bitte wähle den gewünschten Speicherort aus ")

    def start(self):
        start.run(self.QumaPath, self.RechnungPath, self.Speicherort)

