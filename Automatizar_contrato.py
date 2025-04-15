import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from fpdf import FPDF
import re
import locale

# Função para selecionar o arquivo Excel
def selecionar_arquivo():
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx;*.xls")])
    entrada_arquivo.set(arquivo)

# Função para gerar os contratos
def gerar_contratos():
    # Carregar o arquivo Excel
    arquivo_excel = entrada_arquivo.get()
    if not arquivo_excel:
        messagebox.showerror("Erro", "Por favor, selecione um arquivo Excel primeiro.")
        return
    
    df = pd.read_excel(arquivo_excel)
   

    # Lista para armazenar os caminhos dos PDFs gerados
    arquivos_gerados = []

    # Define o formato brasileiro para moeda
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

    for index, row in df.iterrows():
        NOME = str(row['NOME'])
      
        CPF = str(row['CPF'])  # Converte para string
        CPF = re.sub(r"\D", "", CPF)  # Remove tudo que não for número (pontos, traços, espaços, etc.)
        CPF = CPF.zfill(11)  # Garante que tenha 11 dígitos, preenchendo com zeros à esquerda
        CPF = re.sub(r"(\d{3})(\d{3})(\d{3})(\d{2})", r"\1.\2.\3-\4", CPF)  # Formata corretamente


        """
        Aqui você pode criar novas variáveis, de acordo com o que a sua empresa utiliza no dia a dia
        dos seus contratos. Utilizei apenas NOME e CPF para  o exemplo ficar simples e objetivo.
        """
       
      
      
        print(f"""
                ===============================
                Índice: {index}
                CPF: {CPF}
                Nome: {NOME}
             
                
               
                ===============================
                """)
        


        # Gerar o PDF
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()

        pdf.set_font("Arial", "B", 14)
        pdf.cell(200, 10, txt="", ln=True, align="C")
        pdf.ln(10)

        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0, 10, txt=f""" 

                       
Aqui irão os dados do seu contrato. É só jogar no texto e escolher as variáveis que deseja.

____________________________________________
{NOME}
{CPF}





""".encode('latin-1', 'replace').decode('latin-1'))




        # Salvar o contrato em PDF
        # Defina o caminho da pasta de downloads, dependendo do seu sistema operacional
        pasta_downloads = os.path.join(os.path.expanduser('~'), 'Downloads')

        caminho_pdf = os.path.join(pasta_downloads, f"Minuta_{NOME}.pdf")
        pdf.output(caminho_pdf)

        # Adicionar o arquivo à lista de gerados
        arquivos_gerados.append(caminho_pdf)
    # Exibir mensagem ao final do processo
    messagebox.showinfo("Finalizado", "Todos os contratos foram gerados com sucesso!")

    


# Configuração da interface gráfica
root = tk.Tk()
root.title("Sua empresa")
root.configure(bg="#E0F7FA")  # Fundo azul claro

entrada_arquivo = tk.StringVar()




# Frame principal com borda arredondada e sombra
frame = tk.Frame(root, padx=20, pady=20, bg="#64B5F6", bd=5, relief="solid", borderwidth=2)
frame.pack(padx=40, pady=40)

# Botão para selecionar o arquivo Excel com bordas arredondadas e fonte sofisticada
btn_selecionar = tk.Button(frame, text="Selecionar Arquivo", 
                           command=selecionar_arquivo,  # Substitua pela função real
                           bg="#1E88E5", fg="white", font=("Georgia", 12, "bold"), 
                           padx=8, pady=8, relief="flat", bd=0)
btn_selecionar.pack(pady=10)

# Campo de exibição do caminho do arquivo com bordas arredondadas
entrada_arquivo = tk.StringVar()
campo_arquivo = tk.Entry(frame, textvariable=entrada_arquivo, width=50, font=("Georgia", 12), bd=2, relief="sunken")
campo_arquivo.pack(pady=10)

# Frame para organizar os botões lado a lado com um visual limpo
frame_botoes = tk.Frame(frame, bg="#64B5F6")
frame_botoes.pack(pady=20)

# Botão para gerar os contratos
btn_gerar = tk.Button(frame_botoes, text="Gerar Minutas", 
                      command=gerar_contratos,  # Substitua pela função real
                      bg="#1E88E5", fg="white", font=("Georgia", 12, "bold"), 
                      padx=8, pady=8, relief="flat", bd=0)
btn_gerar.pack(side=tk.LEFT, padx=15)



# Texto FEBRAPO@2025 estilizado na parte inferior
label_febrapo = tk.Label(root, text="Sua empresa@2025", font=("Georgia", 16, "italic"), fg="#1E88E5", bg="#E0F7FA")
label_febrapo.pack(side=tk.BOTTOM, pady=15)

# Rodar a interface gráfica
root.mainloop()
