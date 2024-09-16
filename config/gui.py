import tkinter as tk
from tkinter import filedialog, Label, Button, Entry, StringVar
from config.logic import process_csv, save_to_excel
import pandas as pd
import os  # Importar o módulo os para manipulação de arquivos e diretórios

def select_file(label):
    # Função para abrir o explorador de arquivos e selecionar o CSV
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo CSV",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    if file_path:
        label.config(text=file_path)  # Atualiza o rótulo com o caminho do arquivo selecionado
    return file_path

def process_file(file_path, start_date, end_date, status_var, status_label):
    if not file_path:
        status_var.set("Nenhum arquivo foi selecionado.")
        status_label.config(fg="red")  # Define a cor do texto para vermelho
    else:
        try:
            # Se as datas estiverem em branco, definir os limites
            if not start_date:
                start_date = None
            else:
                start_date = pd.to_datetime(start_date, format='%d/%m/%Y')

            if not end_date:
                end_date = None
            else:
                end_date = pd.to_datetime(end_date, format='%d/%m/%Y')

            # Processar o CSV com as datas definidas ou None
            df_agrupado = process_csv(file_path, start_date, end_date)
            
            # Definir o caminho da pasta e do arquivo de saída
            output_dir = 'salvos'
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)  # Cria a pasta 'salvos' se não existir
            
            output_file = os.path.join(output_dir, 'resultado.xlsx')
            
            # Salvar os dados processados no Excel
            save_to_excel(df_agrupado, output_file)
            
            # Exibir mensagem de sucesso
            status_var.set("Dados filtrados e salvos no Excel com formatação.")
            status_label.config(fg="green")  # Define a cor do texto para verde
        
        except Exception as e:
            status_var.set(f'Erro ao processar o arquivo: {e}')
            status_label.config(fg="red")  # Define a cor do texto para vermelho

def format_date(event, entry):
    """ Formata a entrada de data enquanto o usuário digita e limita a 10 caracteres """
    text = entry.get().replace("/", "")  # Remove barras anteriores
    if len(text) > 8:  # Limita a 8 caracteres (sem contar as barras)
        text = text[:8]
    
    new_text = ""
    
    if len(text) > 2:
        new_text += text[:2] + "/"
    else:
        new_text += text
    
    if len(text) > 4:
        new_text += text[2:4] + "/"
    elif len(text) > 2:
        new_text += text[2:4]
    
    if len(text) > 4:
        new_text += text[4:]
    
    entry.delete(0, tk.END)
    entry.insert(0, new_text)

def create_gui():
    # Criando a janela principal
    root = tk.Tk()
    root.title("Automação de Processamento de CSV")

    # Rótulo para exibir o caminho do arquivo selecionado
    label = Label(root, text="Nenhum arquivo selecionado", width=50)
    label.pack(pady=10)

    # Botão para selecionar o arquivo CSV
    select_button = Button(root, text="Selecionar Arquivo CSV", command=lambda: select_file(label))
    select_button.pack(pady=10)

    # Input para data de início
    start_date_label = Label(root, text="Data de Início (dd/mm/aaaa):")
    start_date_label.pack(pady=5)
    start_date_entry = Entry(root)
    start_date_entry.pack(pady=5)
    
    # Ativar formatação automática e limitação no campo de data de início
    start_date_entry.bind('<KeyRelease>', lambda event: format_date(event, start_date_entry))

    # Input para data de fim
    end_date_label = Label(root, text="Data de Fim (dd/mm/aaaa):")
    end_date_label.pack(pady=5)
    end_date_entry = Entry(root)
    end_date_entry.pack(pady=5)

    # Ativar formatação automática e limitação no campo de data de fim
    end_date_entry.bind('<KeyRelease>', lambda event: format_date(event, end_date_entry))

    # Variável para armazenar a mensagem de status
    status_var = StringVar()
    status_var.set("")  # Mensagem inicial vazia

    # Rótulo para exibir a mensagem de status
    status_label = Label(root, textvariable=status_var, wraplength=400, justify="left")
    status_label.pack(pady=10)

    # Botão para processar o arquivo
    process_button = Button(
        root, 
        text="Processar Arquivo", 
        command=lambda: process_file(label.cget("text"), start_date_entry.get(), end_date_entry.get(), status_var, status_label)
    )
    process_button.pack(pady=10)

    # Loop da janela principal
    root.mainloop()

if __name__ == "__main__":
    create_gui()