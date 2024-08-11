import pandas as pd

planilha = r'C:\Users\moesios\Desktop\DOWNLOAD DE DADOS\BD.xlsx'

# Carregar as planilhas
df_morador = pd.read_excel(planilha, sheet_name='MORADOR')
df_domicilio = pd.read_excel(planilha, sheet_name='DOMICÍLIO')
df_menu = pd.read_excel(planilha, sheet_name='MENU')

# Contagem de moradores por domicílio
contagem_morador = df_morador['ID_DOMICILIO'].value_counts().reset_index()
contagem_morador.columns = ['ID_DOMICILIO', 'QTD_PESSOAS_MORADOR']

# Merge com a tabela de domicílio
df_domicilio = df_domicilio.merge(contagem_morador, on='ID_DOMICILIO', how='left')

# Validação e ajuste da quantidade de pessoas por domicílio
df_domicilio['VALIDAÇÃO'] = df_domicilio['QUANTAS PESSOAS MORAM NESSE DOMICÍLIO ?'] == df_domicilio['QTD_PESSOAS_MORADOR']
ajustes_realizados = df_domicilio['VALIDAÇÃO'] == False
df_domicilio.loc[ajustes_realizados, 'QUANTAS PESSOAS MORAM NESSE DOMICÍLIO ?'] = df_domicilio.loc[ajustes_realizados, 'QTD_PESSOAS_MORADOR']
df_domicilio['AJUSTADO'] = ajustes_realizados.replace({True: 'SIM', False: 'NÃO'})

# Verificação dos itens de conforto
colunas_para_verificar = [
    'GELADEIRA', 'FREEZER', 'LAVA-LOUÇAS', 'MÁQUINA DE LAVAR ROUPA',
    'MÁQUINA DE SECAR ROUPA', 'MICRO-ONDAS', 'COMPUTADOR/NOTEBOOK',
    'DVD/VIDEOGAME', 'BANHEIROS', 'QUANTOS EMPREGADOS DOMÉSTICOS QUE TRABALHAM PELO MENOS 5 DIAS POR SEMANA',
    'BICICLETAS', 'MOTOCICLETAS', 'AUTOMÓVEIS', 'TEM ÁGUA ENCANADA NESTE DOMICÍLIO ?'
]
df_domicilio['VERIFICAÇÃO DOS ITENS DE CONFORTO PREENCHIDOS'] = df_domicilio.apply(
    lambda row: '/'.join(
        [col for col in colunas_para_verificar if pd.isna(row[col])]
    ) if any(pd.isna(row[col]) for col in colunas_para_verificar) else 'OK',
    axis=1
)

# Merge com a tabela de menu
df_domicilio = df_domicilio.merge(df_menu[['ID_MENU', 'DATA']], on='ID_MENU', how='left')

# Salvando de volta no arquivo Excel
with pd.ExcelWriter(planilha, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    # Escrever todas as abas menos DOMICÍLIO
    for sheet_name in pd.ExcelFile(planilha).sheet_names:
        if sheet_name != 'DOMICÍLIO':
            pd.read_excel(planilha, sheet_name=sheet_name).to_excel(writer, sheet_name=sheet_name, index=False)
    
    # Escrever a aba DOMICÍLIO atualizada
    df_domicilio.to_excel(writer, sheet_name='DOMICÍLIO', index=False)


print("DEU CERTO")