import os
import time
import threading
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import win32print
import io
import sys

# Lista de impressoras
impressoras_lista = []

def listar_impressoras():
    # Obtém todas as impressoras disponíveis
    impressoras = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    for impressora in impressoras:
        x, y, nome, z = impressora[:4]
        print(nome)
        impressoras_lista.append(nome)

listar_impressoras()

def selecionar_pasta():
    caminho_pasta = filedialog.askdirectory()
    
    if caminho_pasta:
        entry_path.delete(0, tk.END)
        entry_path.insert(0, caminho_pasta)

def nome_impressora_():
    nome_impressora = impressora_escolhida.get()
    return nome_impressora

def caminho_pasta_():
    caminho_pasta_value = entry_path.get()
    return caminho_pasta_value


# Variável para controlar o loop de monitoramento
monitoring = threading.Event()

def list_files(path):
    """ Retorna um dicionário com o nome do arquivo e seu timestamp """
    return {f: os.path.getmtime(os.path.join(path, f)) for f in os.listdir(path)}

def enviar_zpl_para_impressora(zpl_command, printer_name):
    """ Envia comandos ZPL para a impressora """
    try:
        # Conecte-se à impressora
        conectar_impressora = win32print.OpenPrinter(printer_name)
        try:
            # Crie um objeto de trabalho
            hJob = win32print.StartDocPrinter(conectar_impressora, 1, ("ZPL Print Job", None, "RAW"))
            try:
                # Inicie a página
                win32print.StartPagePrinter(conectar_impressora)
                # Envie o comando ZPL
                win32print.WritePrinter(conectar_impressora, zpl_command.encode())
                # Termine a página
                win32print.EndPagePrinter(conectar_impressora)
                # Termine o trabalho
                win32print.EndDocPrinter(conectar_impressora)
            finally:
                # Feche o objeto de trabalho
                win32print.ClosePrinter(conectar_impressora)
        except Exception as e:
            print(f"Erro ao criar ou enviar o trabalho para a impressora: {e}")
    except Exception as e:
        print(f"Erro ao abrir a impressora: {e}")

def imprimir_testes(zpl_command, printer_name):
    print(f'Imprimindo na impressora: {printer_name}')
    x = zpl_command

def ler_codigo_zpl(file_path):
    """Lê o código ZPL do arquivo de texto"""
    with open(file_path, 'r') as file:
        return file.read()

def monitor_directory(caminho_pasta, interval=1):
    """ Monitora o diretório para alterações a cada 'interval' segundos """
    arquivos_anteriores = list_files(caminho_pasta)
    print(f"Monitorando mudanças na pasta: {caminho_pasta}")

    while not monitoring.is_set():
        nome_impressora = nome_impressora_()
        time.sleep(interval)
        arquivos_atuais = list_files(caminho_pasta)

        # Verifica arquivos criados ou modificados
        for arquivos in arquivos_atuais:
            if arquivos not in arquivos_anteriores:
                print(f"Arquivo criado: {arquivos}")
                if arquivos.endswith('.txt'):  # Verifica se é um arquivo .txt
                    caminho_do_arquivo = os.path.join(caminho_pasta, arquivos)
                    comando_zpl = ler_codigo_zpl(caminho_do_arquivo)
                    # Função de imprimir real
                    imprimir_testes(comando_zpl, nome_impressora)
                    enviar_zpl_para_impressora(comando_zpl, nome_impressora)

            elif arquivos_atuais[arquivos] != arquivos_anteriores[arquivos]:
                print(f"Arquivo modificado: {arquivos}")
                if arquivos.endswith('.txt'):  # Verifica se é um arquivo .txt
                    caminho_do_arquivo = os.path.join(caminho_pasta, arquivos)
                    comando_zpl = ler_codigo_zpl(caminho_do_arquivo)
                    # Função de imprimir real
                    imprimir_testes(comando_zpl, nome_impressora)
                    enviar_zpl_para_impressora(comando_zpl, nome_impressora)           

        # Verifica arquivos deletados
        for arquivos in arquivos_anteriores:
            if arquivos not in arquivos_atuais:
                print(f"Arquivo deletado: {arquivos}")

        # Atualiza o estado anterior
        arquivos_anteriores = arquivos_atuais

def iniciar_monitoramento():
    """ Inicia o monitoramento em uma nova thread """
    caminho_pasta = caminho_pasta_()
    if not caminho_pasta:
        messagebox.showwarning("Aviso", "Por favor, insira o caminho da pasta.")
        return
    # Desativa o botão de iniciar para evitar múltiplas instâncias
    btn_start.config(state=tk.DISABLED)
    # Ativa o botão de parar
    btn_stop.config(state=tk.NORMAL)
    # Limpa a área de texto
    text_console.delete(1.0, tk.END)
    # Sinaliza para iniciar o monitoramento
    monitoring.clear()
    # Inicia o monitoramento em uma nova thread
    threading.Thread(target=monitor_directory, args=(caminho_pasta,), daemon=True).start()

def parar_monitoramento():
    """ Para o monitoramento """
    # Sinaliza para parar o monitoramento
    monitoring.set()
    # Ativa o botão de iniciar
    btn_start.config(state=tk.NORMAL)
    # Desativa o botão de parar
    btn_stop.config(state=tk.DISABLED)

def redirect_output():
    """ Redireciona a saída padrão para o Text widget """
    class StdoutRedirector(io.StringIO):
        def write(self, message):
            text_console.insert(tk.END, message)
            text_console.see(tk.END)
    sys.stdout = StdoutRedirector()

# Configuração da interface gráfica
root = tk.Tk()
root.title("Impressão Automática ZPL")

# Variável para manter a seleção
impressora_escolhida = tk.StringVar()

# Widget de Label
label_nome_impressora = ttk.Label(root, text="Escolha um nome:")
label_nome_impressora.place(x=200, y=200)


# Caminho da pasta
label_path = tk.Label(root, text="Caminho da pasta:")
label_path.place(x=10, y=0)

entry_path = tk.Entry(root, width=50)
entry_path.place(x=10, y=20)

botao_selecionar_pasta = tk.Button(root, text="Selecionar Pasta", command=selecionar_pasta)
botao_selecionar_pasta.place(x=320, y=16, width=100)

# Área de texto para o console
text_console = tk.Text(root, height=10, width=50)
text_console.place(x=10, y=115)

texto_nome_impressora = tk.Label(root, text="Selecionar impressora:")
texto_nome_impressora.place(x=10, y=45)

# OptionMenu
option_menu = ttk.OptionMenu(root, impressora_escolhida, impressoras_lista[0], *impressoras_lista)
option_menu.place(x=135, y=45)


# Botão de iniciar
btn_start = tk.Button(root, text="Iniciar Monitoramento", command=iniciar_monitoramento)
btn_start.place(x=25, y=75)

# Botão de parar
btn_stop = tk.Button(root, text="Parar Monitoramento", command=parar_monitoramento, state=tk.DISABLED)
btn_stop.place(x=170, y=75)

# Redireciona a saída padrão para a área de texto
redirect_output()

# Configuração da janela
root.geometry("425x285")
root.mainloop()
