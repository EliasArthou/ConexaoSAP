import sys
import tkinter as tk
from tkinter import ttk
from tkinter.ttk import Progressbar


class App(tk.Tk):
    """
    Cria janela com retorno pro usuário
    """

    def __init__(self):
        super().__init__()

        w = 300
        h = 100

        ws = self.winfo_screenwidth()
        hs = self.winfo_screenheight()
        x = (ws / 2) - (w / 2)
        y = (hs / 2) - (h / 2)

        self.title('Andamento Jobs')
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))

        # configure the root window

        # self.geometry('300x100')
        self.attributes("-topmost", True)

        # label
        self.labeljob = ttk.Label(self, text='')
        self.labeljob.pack()

        # barra de progresso
        self.barra = Progressbar(self, orient=tk.HORIZONTAL, length=100, mode='determinate')
        self.barra.place(x=300, y=10)
        self.barra.pack()

        # label
        self.labelpassos = ttk.Label(self, text='')
        self.labelpassos.pack()

        # button
        self.button = ttk.Button(self, text='Fechar')
        self.button['command'] = self.button_clicked
        self.button.pack()

    def button_clicked(self):
        """
        Ação do botão
        """
        self.destroy()
        sys.exit()

    def mudartexto(self, nomelabel, texto):
        """
        :param nomelabel: nome do label a ter o texto alterado
        :param texto: texto a ser inserido
        """
        self.__getattribute__(nomelabel).config(text=texto)
        self.atualizatela()

    def configurarbarra(self, maximo, indicador):
        """
        :param maximo: limite máximo da barra de progresso.
        :param indicador: variável
        """
        self.barra.config(maximum=maximo, value=indicador)
        self.atualizatela()

    def atualizatela(self):
        """
        Dá um 'refresh' na tela para modificar com alterações realizadas
        """
        self.update()


if __name__ == "__main__":
    app = App()
    app.mainloop()
