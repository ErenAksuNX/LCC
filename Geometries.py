import tkinter


# Diese Funktion gibt einer GUI den Standard Layout
def geometries(master: tkinter.Tk):
    master.title("National-Express LCC Zuweisung")  # Der Title
    master.geometry("500x500")  # Die Start größe wird hier festgelegt
    master.minsize(500, 500)  # Die Mindestgröße wird festgelegt
    master.configure(background="#8ec9e9")  # Die Hintergrundfarbe wird hinzugefügt
    master.iconbitmap("data/national-express.ico")  # Das Icon wird erstellt


# Diese Funktion placed die Buttons
# ! Diese Funktion funktioniert nur richtig, wenn es 4 Buttons gibt
def place_Buttons(Buttons):
    # Es werden alle Buttons durch eine schleife in der mitte geplaced
    for i, button in enumerate(Buttons):
        button.place(anchor="center", rely=(i + 1) / 5, relx=.5)
        # Die berechnung funktioniert wie folgt:
        # i = nummer des Knopfes + 1, weil enumerate erst ab 0 anfängt zu zählen
        # / 5 damit der Prozent satz der relative Punt auf er y-Achse berechnet werden kann
