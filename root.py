import tkinter as tk
from tkinter import StringVar

class Janela(tk.Tk):
    
    def __init__(self):
        super().__init__()

        self.title('Data')
        self.geometry("200x60")

        self.var = StringVar(self)

        tk.Label(self, text='Data:', font=('calibre', 10, 'bold')).grid(row=0, column=0)
        self.entry = tk.Entry(self, textvariable=self.var,
                      font=('calibre', 10, 'normal'), width=20)
        self.entry.focus_set()
        self.entry.grid(row=0, column=1)

        tk.Button(self, text='OK', command=self.getInput, width=15).grid(row=5, column=1)

    def getInput(self) -> str:
        data = self.var.get()
        data = data.replace("/","")
        self.quit()
        return data

if __name__ == '__main__':
    app = Janela()
    app.mainloop()