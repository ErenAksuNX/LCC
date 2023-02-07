import tkinter


def geometries(master: tkinter.Tk):
    master.title("National-Express LCC Zuweisung")
    master.geometry("500x500")
    master.minsize(500, 500)
    master.configure(background="#8ec9e9")
    master.iconbitmap("data/national-express.ico")


def place_Buttons(Buttons):
    for i, button in enumerate(Buttons):
        button.place(anchor="center", rely=(i + 1) / 5, relx=.5)
