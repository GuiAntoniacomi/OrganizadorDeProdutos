import pandas as pd

# Carregar o arquivo JSON
arquivo = pd.read_json(r"C:\Users\anton\Downloads\todos_os_produtos_2024-05-29T15_05_06.99827Z.json")

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

# Função para remover o nome da marca da categoria
def remover_marca(categoria, marca):
    if pd.isna(categoria) or pd.isna(marca):
        return categoria
    if marca in categoria:
        return categoria.replace(marca, '').strip()
    return categoria

# Aplicar a função para remover a marca da coluna Categoria
df_agrupada['Categoria'] = df_agrupada.apply(lambda row: remover_marca(row['Categoria'], row['Marca']), axis=1)

# Exibir o DataFrame resultante
print(df_agrupada)

df_marca = df_resumida.groupby('Marca').agg({
    'SKU Pai': 'first',
    'Produto': 'first',
    'Link': 'first',
    'Estoque':'sum',
    'Categoria': 'first',
    'Total Preço':'sum',
    'Total Custo':'sum',
    'Preço':'sum',
    'Custo':'sum'
}).reset_index()

df_marca = df_marca[['Marca', 'Estoque', 'Total Preço', 'Total Custo']]
df_marca = df_marca.sort_values(by='Estoque', ascending=False).reset_index(drop=True)

#import ace_tools as tools; tools.display_dataframe_to_user(name="df_marca Ordenado", dataframe=df_marca)

# Calculando os totais
total_geral = pd.DataFrame({
    'Marca': ['Total Geral'],
    'Estoque': [df_marca['Estoque'].sum()],
    'Total Preço': [df_marca['Total Preço'].sum()],
    'Total Custo': [df_marca['Total Custo'].sum()]
})

# Adicionando a linha de total geral ao DataFrame
df_marca = pd.concat([df_marca, total_geral], ignore_index=True)

print(df_marca)