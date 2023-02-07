import tkinter


# Diese Funktion gibt einer GUI den Standard Layout
def geometries(master: tkinter.Tk):
    master.title("National-Express LCC Zuweisung")
    master.geometry("500x500")
    master.minsize(500, 500)
    master.configure(background="#8ec9e9")
    master.iconbitmap("data/national-express.ico")


# Diese Funktion placed die Buttons
# ! Diese Funktion funktioniert nur richtig wenn es 4 Buttons gibt
def place_Buttons(Buttons):
    for i, button in enumerate(Buttons):
        button.place(anchor="center", rely=(i + 1) / 5, relx=.5)
