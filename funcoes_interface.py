from pathlib import Path
from tkinter import Tk, filedialog, messagebox

def escolher_arquivos():
    janela = Tk()
    janela.withdraw()

    modo_atualizar = messagebox.askyesnocancel(
        title="Escolha uma ação",
        message=(
            "Deseja atualizar uma planilha Excel existente?\n\n"
            "Sim: atualizar uma planilha existente.\n"
            "Não: criar uma nova planilha."
        ),
    )

    if modo_atualizar is None:
        janela.destroy()
        print("Operação cancelada.")
        return None

    caminho_texto = filedialog.askopenfilename(
        title="Selecione o arquivo de texto",
        filetypes=[("Todos os arquivos", "*")],
    )

    if modo_atualizar:
        caminho_excel = filedialog.askopenfilename(
            title="Escolha a planilha Excel para atualizar",
            filetypes=[("Todos os arquivos", "*")],
        )
        acao = "atualizar"
    else:
        caminho_excel = filedialog.asksaveasfilename(
            title="Escolha onde salvar a planilha Excel",
            defaultextension=".xlsx",
            filetypes=[("Todos os arquivos", "*")],
        )
        acao = "criar"

    janela.destroy()

    if not caminho_texto or not caminho_excel:
        print("Operação cancelada: dois arquivos são necessários.")
        return None

    arquivo_texto = Path(caminho_texto)
    arquivo_excel = Path(caminho_excel)

    if arquivo_texto.suffix.lower() != ".txt":
        print("Erro: o arquivo de inventário precisa terminar em .txt.")
        return None

    if arquivo_excel.suffix.lower() != ".xlsx":
        print("Erro: a planilha precisa terminar em .xlsx.")
        return None

    return acao, arquivo_texto, arquivo_excel