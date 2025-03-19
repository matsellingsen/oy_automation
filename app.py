import tkinter as tk
from tkinter import TOP, filedialog, Button
from calc import readAndSummarize 

def run():
    def create_reports():
        filepath = filedialog.askopenfilename()
        readAndSummarize(filepath)
        app.destroy()

    app = tk.Tk()
    app.geometry("100x50")
    app.eval('tk::PlaceWindow . center')
    B = Button(app, text ="select report", command = create_reports)
    B.pack(side=TOP)

    app.mainloop()


if __name__ == "__main__":
    run()

