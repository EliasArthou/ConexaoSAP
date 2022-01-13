import tkinter as tk
from tkinter import ttk
from tkinter.ttk import Progressbar

class App(tk.Tk):
    def __init__(self, view, valorbarra, total):
        super().__init__()

        # configure the root window
        self.title('Andamento Jobs')
        self.geometry('300x100')

        # label
        self.labeljob = ttk.Label(self, text=view)
        self.labeljob.pack()

        # barra de progresso
        self.barra = Progressbar(self, orient=tk.HORIZONTAL, length=100, mode='determinate',
                                 variable=valorbarra, maximum=total)
        self.barra.place(x=300, y=10)
        self.barra.pack()

        # button
        self.button = ttk.Button(self, text='Fechar')
        self.button['command'] = self.button_clicked
        self.button.pack()

    def button_clicked(self):
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
