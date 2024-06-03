import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook

def remover_marca(categoria, marca):
    if pd.isna(categoria) or pd.isna(marca):
        return categoria
    if marca in categoria:
        return categoria.replace(marca, '').strip()
    return categoria

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
    df_resumida = df_bagy[['Brands → Name', 'Categories - Category Default → Name', 'Price', 'Cost', 'Slug', 'Name', 'Stocks → Balance', 'External ID']]
    df_resumida = df_resumida.rename(columns={'Brands → Name': 'Marca', 'Categories - Category Default → Name': 'Categoria', 'Price': 'Preço', 'Cost': 'Custo', 'Slug': 'Link', 'Name': 'Produto', 'Stocks → Balance': 'Estoque', 'External ID': 'SKU Pai'})
    ordem_colunas = ['SKU Pai', 'Produto', 'Link', 'Estoque', 'Marca', 'Categoria', 'Preço', 'Custo']
    df_resumida = df_resumida[ordem_colunas]
    df_resumida['Total Preço'] = df_resumida['Preço'] * df_resumida['Estoque']
    df_resumida['Total Custo'] = df_resumida['Estoque'] * df_resumida['Custo']

    df_agrupada = df_resumida.groupby('SKU Pai').agg({
        'Produto': 'first',
        'Link': 'first',
        'Estoque': 'sum',
        'Marca': 'first',
        'Categoria': 'first',
        'Preço': 'sum',
        'Custo': 'sum',
        'Total Preço': 'sum',
        'Total Custo': 'sum'
    }).reset_index()
    df_agrupada = df_agrupada.sort_values(by='Estoque', ascending=False).reset_index(drop=True)
    df_agrupada['Marca'].dropna()

    # Aplicar a função para remover a marca da coluna Categoria
    df_agrupada['Categoria'] = df_agrupada.apply(lambda row: remover_marca(row['Categoria'], row['Marca']), axis=1)

    df_categoria = clean_and_map_categories(df_agrupada)

    # Alterar categoria para "Camisas Sociais MC" quando o produto tiver "manga curta"
    df_categoria.loc[df_categoria['Produto'].str.contains('manga curta', case=False, na=False) & (df_categoria['Categoria'] == "Camisas Sociais"), 'Categoria'] = "Camisas Sociais MC"

    # Remover linhas onde a coluna "Categoria" está vazia
    df_categoria = df_categoria.dropna(subset=['Categoria'])

    # Exibir resultado filtrado
    camisas_sociais = df_categoria.loc[df_categoria['Categoria'] == "Camisas Sociais"]

    # Nome da nova aba
    new_sheet_name = "Base - Ordem de Estoque"

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

# Criar a interface gráfica
root = tk.Tk()
root.title("Processador de Arquivos")

tk.Label(root, text="Selecione o arquivo JSON:").grid(row=0, column=0, padx=10, pady=10)
json_file_entry = tk.Entry(root, width=50)
json_file_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Procurar", command=select_json_file).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Selecione o arquivo Excel:").grid(row=1, column=0, padx=10, pady=10)
excel_file_entry = tk.Entry(root, width=50)
excel_file_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Procurar", command=select_excel_file).grid(row=1, column=2, padx=10, pady=10)

tk.Button(root, text="Processar", command=run_process).grid(row=2, column=0, columnspan=3, pady=20)

root.mainloop()
