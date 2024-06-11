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
    bling = bling_file_entry.get()
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
        json_file_entry.delete(0, tk.END)
        json_file_entry.insert(0, file_path)

def select_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        excel_file_entry.delete(0, tk.END)
        excel_file_entry.insert(0, file_path)

def select_bling_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        bling_file_entry.delete(0, tk.END)
        bling_file_entry.insert(0, file_path)

def run_process():
    json_file = json_file_entry.get()
    excel_file = excel_file_entry.get()
    if json_file and excel_file:
        try:
            process_files(json_file, excel_file)
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
    else:
        messagebox.showwarning("Aviso", "Por favor, selecione ambos os arquivos JSON e Excel.")

window = Tk()
window.title("Ferramenta de Processamento de Dados")
window.geometry("700x400")
window.configure(bg = "#ffffff")
canvas = Canvas(
    window,
    bg = "#ffffff",
    height = 400,
    width = 700,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge")
canvas.place(x = 0, y = 0)

background_img = PhotoImage(file=resource_path("FrontEnd/background.png"))
background = canvas.create_image(
    363.0, 200.0,
    image=background_img)

img0 = PhotoImage(file=resource_path("FrontEnd/img0.png"))
b0 = Button(
    image = img0,
    borderwidth = 0,
    highlightthickness = 0,
    bg = '#2E2E2E',
    command = select_json_file,
    relief = "groove")

b0.place(
    x = 613, y = 143,
    width = 66,
    height = 20)

entry0_img = PhotoImage(file=resource_path("FrontEnd/img_textBox0.png"))
entry0_bg = canvas.create_image(
    496.0, 153.0,
    image = entry0_img)

entry0 = Entry(
    bd = 0,
    bg = "#f0f0f0",
    highlightthickness = 0)

entry0.place(
    x = 390, y = 143,
    width = 212,
    height = 18)

json_file_entry = tk.Entry(window, width=50)

json_file_entry.place(
    x = 390, y = 143,
    width = 212,
    height = 18)

img1 = PhotoImage(file=resource_path("FrontEnd/img1.png"))
b1 = Button(
    image = img1,
    borderwidth = 0,
    highlightthickness = 0,
    bg = '#2E2E2E',
    command = select_bling_file,
    relief = "flat")

b1.place(
    x = 613, y = 200,
    width = 66,
    height = 20)

entry1_img = PhotoImage(file=resource_path("FrontEnd/img_textBox1.png"))
entry1_bg = canvas.create_image(
    496.0, 210.0,
    image = entry1_img)

entry1 = Entry(
    bd = 0,
    bg = "#f0f0f0",
    highlightthickness = 0)

entry1.place(
    x = 390, y = 200,
    width = 212,
    height = 18)

bling_file_entry = tk.Entry(window, width=50)

bling_file_entry.place(
    x = 390, y = 200,
    width = 212,
    height = 18)

img2 = PhotoImage(file=resource_path("FrontEnd/img2.png"))
b2 = Button(
    image = img2,
    borderwidth = 0,
    highlightthickness = 0,
    bg = '#2E2E2E',
    command = select_excel_file,
    relief = "flat")

b2.place(
    x = 613, y = 257,
    width = 66,
    height = 20)

entry2_img = PhotoImage(file=resource_path("FrontEnd/img_textBox2.png"))
entry2_bg = canvas.create_image(
    496.0, 267.0,
    image = entry2_img)

entry2 = Entry(
    bd = 0,
    bg = "#f0f0f0",
    highlightthickness = 0)

entry2.place(
    x = 390, y = 257,
    width = 212,
    height = 18)

excel_file_entry = tk.Entry(window, width=50)

excel_file_entry.place(
    x = 390, y = 257,
    width = 212,
    height = 18)

img3 = PhotoImage(file=resource_path("FrontEnd/img3.png"))
b3 = Button(
    image = img3,
    borderwidth = 0,
    highlightthickness = 0,
    bg = '#2E2E2E',
    command = run_process,
    relief = "flat")

b3.place(
    x = 436, y = 316,
    width = 98,
    height = 37)

window.resizable(False, False)
window.mainloop()
