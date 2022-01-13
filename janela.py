import sys
import tkinter as tk
from tkinter import ttk
from tkinter.ttk import Progressbar


class App(tk.Tk):
    """
    Cria janela com retorno para o usuário
    """

    def __init__(self):
        super().__init__()

        # Largura da Janela
        w = 300
        # Altura da Janela
        h = 200

        # Tamanho da tela total na horizontal (provavelmente resolução)
        ws = self.winfo_screenwidth()
        # Tamanho da tela total na vertical (provavelmente resolução)
        hs = self.winfo_screenheight()
        # Calculo do centro da tela
        x = (ws / 2) - (w / 2)
        y = (hs / 2) - (h / 2)

        # Define a janela como não exclusiva (outras janelas podem sobrepor ela)
        self.acertaconfjanela(False)
        # Adiciona o cabeçalho da janela
        self.title('Andamento Jobs')
        # Desenha a janela com a largura e altura definida e na posição calculada, ou seja, no centro da tela
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))

        # Label do nome do JOB
        self.labeljob = ttk.Label(self, text='', font="Arial 25 bold")
        self.labeljob.pack()

        # label contagem de transações (Views)
        self.statustrans = ttk.Label(self, text='', font="Arial 10")
        self.statustrans.pack()

        # ProgressBar de Quantidade de Transações (Views)
        self.barratrans = Progressbar(self, orient=tk.HORIZONTAL, length=200, mode='determinate')
        # self.barratrans.place(x=300, y=10)
        self.barratrans.pack()

        # Label de status do JOB
        self.statusjobs = ttk.Label(self, text='', font="Arial 10")
        self.statusjobs.pack()

        # ProgressBar
        self.barrajob = Progressbar(self, orient=tk.HORIZONTAL, length=200, mode='determinate')
        # self.barrajob.place(x=300, y=10)
        self.barrajob.pack()

        # Label Ação do Momento...
        self.labelpassos = ttk.Label(self, text='')
        self.labelpassos.pack()

        # button Fechar a Janela
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

    def configurarbarra(self, nomebarra, maximo, indicador):
        """
        :param nomebarra: nome da barra a ser atualizada.
        :param maximo: limite máximo da barra de progresso.
        :param indicador: variável
        """
        self.__getattribute__(nomebarra).config(maximum=maximo, value=indicador)
        self.atualizatela()

    def acertaconfjanela(self, exclusiva):
        """
        :param exclusiva: se a janela fica na frente das outras ou não
        """
        self.attributes("-topmost", exclusiva)
        self.atualizatela()

    def atualizatela(self):
        """
        Dá um 'refresh' na tela para modificar com alterações realizadas
        """
        self.update()


if __name__ == "__main__":
    app = App()
    app.mainloop()
