import auditoria_lojas as au
from tkinter import messagebox

try:
    au.criarPasta()
    au.configuraFirefox()
    au.navegaSite()
    au.excel()
except:
    messagebox.showerror(title = "Erro", message="Ocorreu um erro, favor refazer o processo")

