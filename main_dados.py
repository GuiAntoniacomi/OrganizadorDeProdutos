from tkinter import *
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
import os
import sys

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def clean_and_map_categories(df):
    category_mapping = {
        'camiseta': 'Camisetas',
        'Boné': 'Bonés',
        'Polos': 'Polos',
        'Calças': 'Calças',
        'Bermuda': 'Bermudas',
        'Tênis': 'Tênis',
        'Blusas': 'Moda Inverno',
        'Camisa': 'Camisas Sociais',
        'Carteira': 'Carteiras',
        'Cueca': 'Cuecas',
        'Chinelo': 'Chinelos',
        'Cinto': 'Cintos',
        'Meia': 'Meias',
        'Pijama': 'Pijamas',
        'Mochila': 'Mochilas e Malas',
        'Sunga': 'Sungas',
        'Calçado': 'Calçados',
        'Aramis': 'Acessórios',
        'King & Joe': 'Chapéus',
        'Marcas': 'Acessórios',
        'Polo Ralph Lauren': 'Polos',
        'U.S. Polo Assn.': 'Polos',
        'Reserva': 'Acessórios',
        'Sergio K': 'Acessórios'
    }

    # Aplicando as substituições
    for keyword, new_category in category_mapping.items():
        df.loc[df['Categoria'].str.contains(keyword, case=False, na=False), 'Categoria'] = new_category

    # Removendo as linhas com valores nulos na coluna 'Categoria'
    df.dropna(subset=['Categoria'], inplace=True)

    return df

def process_files(json_file, excel_file):
    # Carregar o arquivo JSON
    arquivo = pd.read_json(json_file)
    df_bagy = pd.DataFrame(arquivo)
    df_bagy = df_bagy[['Brands → Name', 'Categories - Category Default → Name', 'Price', 'Cost', 'Slug', 'Name', 'Stocks → Balance', 'External ID', 'Variations → Sku']]
    df_bagy = df_bagy.rename(columns={'Brands → Name': 'Marca', 'Categories - Category Default → Name': 'Categoria', 'Price': 'Preço', 'Cost': 'Custo', 'Slug': 'Link', 'Name': 'Produto', 'Stocks → Balance': 'Estoque', 'External ID': 'SKU Pai', 'Variations → Sku': 'Código'})
    ordem_colunas = ['SKU Pai', 'Código', 'Produto', 'Link', 'Estoque', 'Marca', 'Categoria', 'Preço', 'Custo']
    df_bagy = df_bagy[ordem_colunas]
    df_bagy['Total Preço'] = df_bagy['Preço'] * df_bagy['Estoque']
    df_bagy['Total Custo'] = df_bagy['Estoque'] * df_bagy['Custo']

    df_bagy = df_bagy.dropna(subset=['Código'])
    df_bagy['Código'] = df_bagy['Código'].astype(int)
    df_agrupada = df_bagy.sort_values(by='Estoque', ascending=False).reset_index(drop=True)
    df_agrupada['Marca'].dropna()

    df_categoria = clean_and_map_categories(df_agrupada)

    # Alterar categoria para "Camisas Sociais MC" quando o produto tiver "manga curta"
    df_categoria.loc[df_categoria['Produto'].str.contains('manga curta', case=False, na=False) & (df_categoria['Categoria'] == "Camisas Sociais"), 'Categoria'] = "Camisas Sociais MC"

    # Remover linhas onde a coluna "Categoria" está vazia
    df_categoria = df_categoria.dropna(subset=['Categoria'])

    # Nome da nova aba
    new_sheet_name = "Base - Ordem de Estoque"

    # Conferir se existe tabela do Bling
    bling = entry_bling.get()
    if bling:
        df_bling = pd.read_excel(bling)
        df_categoria['Código'] = df_categoria['Código'].astype(int)
        df_bling = pd.read_excel(bling)
        df_bling = df_bling.iloc[:, [0, 5, 6]]  # Seleciona as colunas 0, 5 e 6
        df_categoria = df_categoria.merge(df_bling, how='left', on='Código')
    else:
        pass

    # Salvar o DataFrame na aba desejada
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_categoria.to_excel(writer, sheet_name=new_sheet_name, index=False)

    messagebox.showinfo("Sucesso", f"Dados salvos na aba '{new_sheet_name}' no arquivo Excel existente.")

def select_json_file():
    file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
    if file_path:
        entry_json.delete(0, tk.END)
        entry_json.insert(0, file_path)

def select_bling_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry_bling.delete(0, tk.END)
        entry_bling.insert(0, file_path)

def select_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry_excel.delete(0, tk.END)
        entry_excel.insert(0, file_path)

def run_process():
    json_file = entry_json.get()
    excel_file = entry_excel.get()
    if json_file and excel_file:
        try:
            process_files(json_file, excel_file)
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
    else:
        messagebox.showwarning("Aviso", "Por favor, selecione ambos os arquivos JSON e Excel.")

window = Tk()
window.title("Ferramenta de Processamento de Dados")
window.geometry("1368x760")
window.configure(bg = "#ffffff")
canvas = Canvas(
    window,
    bg = "#ffffff",
    height = 760,
    width = 1368,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge")
canvas.place(x = 0, y = 0)

background_img = PhotoImage(file = f"FrontEnd/background.png")
background = canvas.create_image(
    710.5, 380.5,
    image=background_img)

img0 = PhotoImage(file = f"FrontEnd/img0.png")
btn_gerar = Button(
    image = img0,
    borderwidth = 0,
    highlightthickness = 0,
    bg = '#2E2E2E',
    command = run_process,
    relief = "flat")

btn_gerar.place(
    x = 978, y = 644,
    width = 128,
    height = 53)

# Arquivo Json
img_json = PhotoImage(file = f"FrontEnd/img3.png")
b_json = Button(
    image = img_json,
    borderwidth = 0,
    highlightthickness = 0,
    bg = '#2E2E2E',
    command = select_json_file,
    relief = "flat")

b_json.place(
    x = 1181, y = 297,
    width = 116,
    height = 35)

entry_json_img = PhotoImage(file = f"FrontEnd/img_textBox2.png")
entry2_json_bg = canvas.create_image(
    974.0, 314.5,
    image = entry_json_img)

entry_json = Entry(
    bd = 0,
    bg = "#f0f0f0",
    highlightthickness = 0)

entry_json.place(
    x = 786, y = 297,
    width = 376,
    height = 33)

# Arquvo Bling
img_bling = PhotoImage(file = f"FrontEnd/img2.png")
b_bling = Button(
    image = img_bling,
    borderwidth = 0,
    highlightthickness = 0,
    bg = '#2E2E2E',
    command = select_bling_file,
    relief = "flat")

b_bling.place(
    x = 1181, y = 405,
    width = 117,
    height = 33)

entry_bling_img = PhotoImage(file = f"FrontEnd/img_textBox1.png")
entry1_bling_bg = canvas.create_image(
    974.5, 422.0,
    image = entry_bling_img)

entry_bling = Entry(
    bd = 0,
    bg = "#f0f0f0",
    highlightthickness = 0)

entry_bling.place(
    x = 786, y = 405,
    width = 377,
    height = 32)

# Arquivo Excel
img_excel = PhotoImage(file = f"FrontEnd/img1.png")
b_excel = Button(
    image = img_excel,
    borderwidth = 0,
    highlightthickness = 0,
    bg = '#2E2E2E',
    command = select_excel_file,
    relief = "flat")

b_excel.place(
    x = 1182, y = 512,
    width = 117,
    height = 35)

entry_excel_img = PhotoImage(file = f"FrontEnd/img_textBox0.png")
entry_excel_bg = canvas.create_image(
    974.5, 529.5,
    image = entry_excel_img)

entry_excel = Entry(
    bd = 0,
    bg = "#f0f0f0",
    highlightthickness = 0)

entry_excel.place(
    x = 786, y = 512,
    width = 377,
    height = 33)

window.resizable(False, False)
window.mainloop()
