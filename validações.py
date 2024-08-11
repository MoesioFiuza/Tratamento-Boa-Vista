import pandas as pd

planilha = r'C:\Users\moesios\Desktop\DOWNLOAD DE DADOS\BD.xlsx'
zonas = r'C:\Users\moesios\Desktop\DOWNLOAD DE DADOS\EnderecosDomiciliar_2401_01.xlsx'
logins = r'C:\Users\moesios\Desktop\DOWNLOAD DE DADOS\0_2_Logins.csv'

df_zonas = pd.read_excel(zonas, sheet_name='Domicilios')
df_menu = pd.read_excel(planilha, sheet_name='MENU')
df_logins = pd.read_csv(logins)

df_menu = df_menu.merge(df_zonas[['IDENTIFICADOR DO DOMICÍLIO', 'ZONA']], 
                        how='left', 
                        left_on='IDENTIFICADOR DO DOMICÍLIO', 
                        right_on='IDENTIFICADOR DO DOMICÍLIO')

df_menu['NÚMERO DO DOMICÍLIO'] = df_menu['NÚMERO DO DOMICÍLIO VIZINHO'].combine_first(df_menu['NÚMERO DO DOMICÍLIO'])
df_menu['NÚMERO DO DOMICÍLIO'] = df_menu['NÚMERO DO DOMICÍLIO'].fillna(0).astype(int).astype(str)

df_menu['ENDEREÇO CONCATENADO'] = (
    df_menu['NOME DO LOGRADOURO'].astype(str) + ", " +
    df_menu['NÚMERO DO DOMICÍLIO'] + ", " + 
    df_menu['BAIRRO DO DOMICÍLIO'].astype(str) + " - " +
    "BOA VISTA / RR"
).str.upper()


df_menu = df_menu.merge(df_logins[['ID_LOGIN', 'Nome do App']], 
                        how='left', 
                        left_on='ID_LOGIN', 
                        right_on='ID_LOGIN')

df_menu.rename(columns={'Nome do App': 'Versão do APP'}, inplace=True)


with pd.ExcelWriter(planilha, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    for sheet_name in pd.ExcelFile(planilha).sheet_names:
        if sheet_name != 'MENU':
            pd.read_excel(planilha, sheet_name=sheet_name).to_excel(writer, sheet_name=sheet_name, index=False)
    
    df_menu.to_excel(writer, sheet_name='MENU', index=False)


print("DEU CERTO")    