import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, Menu
import time
import pyautogui
import os
import pyperclip
import webbrowser
from datetime import datetime

# Variáveis globais
df = None
caminho_arquivo = ""
log_envios = []  # Lista para armazenar os envios

def carregar_planilha():
    """Carrega a planilha e verifica se a aba correta existe."""
    global df, caminho_arquivo
    caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx;*.xls")])
    
    if not caminho_arquivo:
        messagebox.showerror("Erro", "Nenhum arquivo selecionado.")
        return
    
    try:
        xls = pd.ExcelFile(caminho_arquivo, engine="openpyxl" if caminho_arquivo.endswith('.xlsx') else "xlrd")
        if "CHAMADOS LASA" not in xls.sheet_names:
            raise ValueError("Aba 'CHAMADOS LASA' não encontrada na planilha.")
        
        df = pd.read_excel(xls, sheet_name="CHAMADOS LASA")

        # Converter a coluna de contato para string
        df["CONTATO"] = df["CONTATO"].astype(str)

        messagebox.showinfo("Sucesso", "Planilha carregada com sucesso!")
        atualizar_tabela()
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar a planilha: {str(e)}")

def atualizar_tabela():
    """Atualiza a exibição da tabela com os dados carregados."""
    for i in tree.get_children():
        tree.delete(i)
    
    if df is not None:
        for _, row in df.iterrows():
            tree.insert("", "end", values=(row["OS"], row["CHAMADO"], row["LOJA"], row["EQUIPAMENTO"], row["RESUMO"], row["CONTATO"], row["STATUS FRESH"]))

def registrar_envio(loja):
    """Registra um envio no log com data e hora."""
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_envios.append({"Loja": loja, "Data e Hora": agora})

def enviar_mensagem():
    """Envia uma mensagem via WhatsApp Desktop."""
    selecionado = tree.focus()
    if not selecionado:
        messagebox.showerror("Erro", "Nenhum chamado selecionado.")
        return
    
    dados = tree.item(selecionado, "values")
    if len(dados) < 6:
        messagebox.showerror("Erro", "Dados incompletos no chamado.")
        return
    
    os, chamado, loja, equipamento, resumo, contato, status_fresh = dados

    if not contato:
        messagebox.showerror("Erro", "Número de contato inválido.")
        return

    status_fresh = status_fresh if status_fresh.strip() else "Sem informação"

    mensagem = f"""
Olá, aqui é da *Zhaz*! 👋  
Temos esse chamado abaixo aberto no sistema.  
Pode confirmar para dar início ao suporte?  

📌 *Detalhes do Chamado*
-------------------------------
🔹 *OS:* {os}
🔹 *Chamado:* {chamado}
🏬 *Loja:* {loja}
🔧 *Equipamento:* {equipamento}
💬 *Resumo:* {resumo}
📞 *Contato:* {contato}
📌 *Status Fresh:* {status_fresh}
-------------------------------
Aguardamos seu retorno! Obrigado. ✅

🔍 *Checklist de Diagnóstico*
1️⃣ Confirma o número da loja?  
2️⃣ Qual o problema do equipamento?  
3️⃣ Quantos equipamentos tem na loja?  
4️⃣ Quantos estão funcionando?  
5️⃣ Pode enviar uma foto ou vídeo do problema?  
6️⃣ A loja tem quedas de energia frequentes?  
7️⃣ Há quanto tempo o terminal está parado?  
8️⃣ O equipamento veio transferido de outra loja?  
9️⃣ A loja possui alguma das redes: MOBILASA ou 0jL4Jr6?  
🔄 *Se voltou de conserto:*  
   🔹 Quando chegou na loja?  
   🔹 Quando foi enviado para ZHAZ para conserto?  
📸 Envie um print das redes Wi-Fi disponíveis na loja.
"""

    try:
        pyperclip.copy(mensagem)
        link_whatsapp = f"whatsapp://send?phone=55{contato}"
        webbrowser.open(link_whatsapp)

        time.sleep(5)

        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.5)
        pyautogui.press("enter")

        registrar_envio(loja)

        messagebox.showinfo("Sucesso", f"Mensagem enviada para a Loja {loja}!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao enviar mensagem: {str(e)}")

def gerar_relatorio():
    """Exibe um relatório de envios na tela."""
    top = tk.Toplevel(root)
    top.title("Relatório de Envios")
    top.geometry("500x300")

    colunas = ["Loja", "Data e Hora"]
    tree_rel = ttk.Treeview(top, columns=colunas, show="headings")

    for col in colunas:
        tree_rel.heading(col, text=col)
        tree_rel.column(col, width=200)

    tree_rel.pack(expand=True, fill="both")

    for envio in log_envios:
        tree_rel.insert("", "end", values=(envio["Loja"], envio["Data e Hora"]))

def baixar_relatorio():
    """Exporta os dados do relatório para um arquivo Excel."""
    if not log_envios:
        messagebox.showerror("Erro", "Nenhum envio registrado.")
        return

    caminho = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])
    
    if not caminho:
        return

    try:
        df_relatorio = pd.DataFrame(log_envios)
        df_relatorio.to_excel(caminho, index=False, engine="openpyxl")
        messagebox.showinfo("Sucesso", f"Relatório salvo em:\n{caminho}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar o relatório: {str(e)}")

def sobre():
    """Exibe informações sobre o programa."""
    messagebox.showinfo("Sobre", "Desenvolvido por Gabriel Menezes\nData: 01/2025")

# Criando interface gráfica
root = tk.Tk()
root.title("Gerenciador de Chamados LASA")
root.geometry("900x500")

# Criando menu
menu_bar = Menu(root)
menu_bar.add_command(label="Sobre", command=sobre)
root.config(menu=menu_bar)

frame_top = tk.Frame(root)
frame_top.pack(pady=10)

btn_carregar = tk.Button(frame_top, text="Carregar Planilha", command=carregar_planilha)
btn_carregar.pack(side=tk.LEFT, padx=10)

btn_enviar = tk.Button(frame_top, text="Enviar Mensagem", command=enviar_mensagem)
btn_enviar.pack(side=tk.LEFT, padx=10)

btn_relatorio = tk.Button(frame_top, text="Gerar Relatório", command=gerar_relatorio)
btn_relatorio.pack(side=tk.LEFT, padx=10)

btn_salvar_relatorio = tk.Button(frame_top, text="Baixar Relatório", command=baixar_relatorio)
btn_salvar_relatorio.pack(side=tk.LEFT, padx=10)

colunas = ["OS", "CHAMADO", "LOJA", "EQUIPAMENTO", "RESUMO", "CONTATO", "STATUS FRESH"]
tree = ttk.Treeview(root, columns=colunas, show="headings")

for col in colunas:
    tree.heading(col, text=col)
    tree.column(col, width=150)

tree.pack(expand=True, fill="both")

root.mainloop()
