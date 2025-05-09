import pandas as pd
import datetime
 
# Nome do arquivo de entrada
nome_arquivo_entrada = "ME5A.xlsx"
 
# Carregar o DataFrame a partir do arquivo Excel processado
df_processado = pd.read_excel(nome_arquivo_entrada)
 
# Colunas
coluna_contagem_Centro = 'Centro'
coluna_contagem_Liberacao = 'StatusLiberacao'
coluna_GC = 'Grupo de compradores'
 
# Adicionar a coluna "StatusLiberacao" com base na nova lógica
df_processado['StatusLiberacao'] = df_processado['Código de liberação'].apply(lambda x: 'Liberada' if x == "2" or pd.isnull(x) else ('Bloqueada' if x == 'X' else 'Bloqueada'))
 
# Obter valores distintos na coluna "0PUR_GROUP"
valores_distintos_GC = df_processado[coluna_GC].unique()
 
# Contagem de valores Centro
contagem_valores_Centro = df_processado[coluna_contagem_Centro].value_counts()
contagem_valores_Liberacao = df_processado[coluna_contagem_Liberacao].value_counts()
 
# Exibir DataFrame atualizado
#print("\nDataFrame com a coluna 'StatusLiberacao' adicionada:")
#print(df_processado)
 
# Exibir a contagem de valores Centro
#print(f"\nQuantidade de Valores Centro: {contagem_valores_Centro}")
#print(f"\nQuantidade de Liberacao: {contagem_valores_Liberacao}")
#print(f"\nQuantidade de GC distintos:")
#print(f"\n{valores_distintos_GC}")
 
# Tabela de prazos
dados = {
    'Grupo de compradores': ['G11','G13','G14','G15','G16','G26','G44','G27','G22','G24','G41','G42','G45','G21','G28','G31','G40','G43','G17','G34','G23','G25','G32','G33','G46','G47','G48','G49','G16'],
    'Prazo': [10, 10, 10, 10, 15, 10, 10, 15, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 30, 10, 30, 90, 45, 30, 45, 10, 45, 25, 10]
}
 
# Criar DataFrame
Tabela_prazos = pd.DataFrame(dados)
 
# Mesclar os DataFrames usando a coluna 'Grupo de compradores'
df_processado = pd.merge(df_processado, Tabela_prazos, on='Grupo de compradores', how='left')
 
# Adicionar a coluna 'Data Prazo' somando 'Data da liberação' e 'Prazo'
df_processado['Data Prazo'] = df_processado['Modificado em'] + pd.to_timedelta(df_processado['Prazo'], unit='D')
 
# Adicionar a coluna 'SomaPrazo' calculando a diferença entre 'Data Prazo' e a data atual
df_processado['SomaPrazo'] = df_processado['Data Prazo'] - pd.to_datetime(datetime.date.today())
 
# Adicionar a coluna 'No prazo?' seguindo a lógica especificada
df_processado['No prazo?'] = df_processado.apply(lambda row: 'Bloqueada' if row['StatusLiberacao'] == 'Bloqueada' else ('Não' if row['StatusLiberacao'] == 'Liberada' and row['SomaPrazo'] < pd.to_timedelta(0, unit='D') else 'Sim'), axis=1)
 
# Adicionar a coluna 'Emergencial' pegando os 3 primeiros dígitos de 'Nº acompanhamento'
df_processado['Emergencial'] = df_processado['Nº acompanhamento'].astype(str).str[:3]
 
# Exibir DataFrame atualizado com os prazos
#print("\nDataFrame atualizado com os prazos:")
#print(df_processado)
 
df_processado.to_excel("teste.xlsx", index=False)
 
# Exibir valores únicos no campo 'Grupo de compradores' e suas contagens
valores_contagem_GC = df_processado['Grupo de compradores'].value_counts()
#print("\nValores únicos no campo 'Grupo de compradores' e suas contagens:")
#print(valores_contagem_GC)
 
# Criar DataFrame com a informação 'valores_contagem_GC'
df_contagem_GC = pd.DataFrame({
    'Grupo de comprador': valores_contagem_GC.index,
    'Quantidade de RCs': valores_contagem_GC.values
})
 
# Adicionar a coluna 'No prazo' seguindo a lógica especificada
df_contagem_GC['Bloqueadas'] = df_contagem_GC['Grupo de comprador'].apply(
    lambda grupo: df_processado[df_processado['Grupo de compradores'] == grupo]['StatusLiberacao'].value_counts().get('Bloqueada', 0)
)
 
# Adicionar a coluna 'No prazo' seguindo a lógica especificada
df_contagem_GC['Liberadas'] = df_contagem_GC['Grupo de comprador'].apply(
    lambda grupo: df_processado[df_processado['Grupo de compradores'] == grupo]['StatusLiberacao'].value_counts().get('Liberada', 0)
)
 
# Adicionar a coluna 'No prazo' seguindo a lógica especificada
df_contagem_GC['No prazo'] = df_contagem_GC['Grupo de comprador'].apply(
    lambda grupo: df_processado[df_processado['Grupo de compradores'] == grupo]['No prazo?'].value_counts().get('Sim', 0)
)
 
# Adicionar a coluna 'Fora prazo' seguindo a lógica especificada
df_contagem_GC['Fora prazo'] = df_contagem_GC['Grupo de comprador'].apply(
    lambda grupo: df_processado[df_processado['Grupo de compradores'] == grupo]['No prazo?'].value_counts().get('Não', 0)
)
 
# Adicionar a coluna 'RCs Emergenciais' seguindo a lógica especificada
df_contagem_GC['RCs Emergenciais'] = df_contagem_GC.apply(
    lambda row: df_processado[
        (df_processado['Grupo de compradores'] == row['Grupo de comprador']) & 
        (df_processado['Emergencial'].str.contains('RE/|RE\\|RE-|RE_', case=False, na=False)) &
        (df_processado['StatusLiberacao'] == 'Liberada')
    ].shape[0],
    axis=1
)
 
# Adicionar a coluna 'Update' com a data atual no formato DD/MM/AAAA
df_contagem_GC['Update'] = datetime.datetime.now().strftime('%d/%m/%Y')
 
# Exibir o novo DataFrame
#print("\nNovo DataFrame com a contagem de 'Grupo de compradores' e 'Quantidade de RCs':")
print(df_contagem_GC)
 
df_contagem_GC.to_excel("Gerenciamento.xlsx", index=False)
 
