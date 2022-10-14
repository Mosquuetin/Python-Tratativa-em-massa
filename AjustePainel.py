import pandas as pd
from datetime import datetime
import numpy as np
import glob
import os
import io
import msoffcrypto
import xlwings as xw
import time

#dia = '03/10/2022'
#hoje = datetime.strptime(dia, '%d/%m/%Y').date()
hoje = datetime.today()
#print(hoje, type(hoje))
if hoje.day == 1 or hoje.day == 3 and hoje.weekday() == 0:
    mes = str(hoje.today().month-1)
    ano = str(hoje.today().year)
else:
    mes = str(hoje.today().month)
    ano = str(hoje.today().year)
#print(mes)

if int(mes) < 10:
    mes = '0' + mes

usuario = os.getcwd().split('\\')[2]
caminho_idpainel = "C:\\Users\\"+usuario+"\\"
caminho_it = "C:\\Users\\"+usuario+"\\"
caminho_final = "C:\\Users\\"+usuario+"\\"
caminho_personal = "C:\\Users\\"+usuario+"\\"

df_duplicadas = pd.DataFrame()
print(caminho_idpainel)
print(caminho_it)
print(caminho_final)
print(caminho_personal)
files = glob.glob(caminho_idpainel+'*.csv')
files_final = glob.glob(caminho_final+'*.csv')
files_it = glob.glob(caminho_it+'*.csv')
if len(files_final) ==1 and len(files_it) ==1:
    for file in files:
        filename = file.split('\\')[9].replace('.csv','')
        print('Lendo arquivo '+filename)
        df = pd.read_csv(file, encoding='ISO-8859-1', low_memory=False)
        df.insert(0, 'Arquivo', value = 0)
        if df_duplicadas.empty:
            df_duplicadas = df
        else:
            df_duplicadas = pd.concat([df_duplicadas,df])#
    files_final = glob.glob(caminho_final+'*.csv')
    for file in files_final:
        filename = file.split('\\')[10].replace('.csv','')
        print('Lendo arquivo final '+filename)
        df = pd.read_csv(file, encoding='ISO-8859-1', low_memory=False)
        df.insert(0, 'Arquivo', value = 1)
        df_final = pd.concat([df,df_duplicadas])
        df_duplicadas = pd.DataFrame()
        df_final = df_final.drop_duplicates(subset=['IdPayment'], keep=False)
        #df_teste = df_final
        #df_teste = df_teste[df_final['Arquivo'] == 0]
        #df_teste = df_teste.drop(['Arquivo'], axis = 1)
        #df_teste.to_csv('C:\\Users\\ELDAN\Directa24 Dropbox\\Auditoria\\2022\\10.2022\\- Arquivos Sistema\\- Painel Admin\\ArquivoPython\\'+'Teste.txt', index=False)
        #Teste Aqui com 0, voltar para 1 para o Padrão
        df_final = df_final[df_final['Arquivo'] == 1]
        df_final = df_final.drop(['Arquivo'], axis = 1)
        df_final.insert(0, 'IT', value = "")
        df_final.insert(0, 'IT2', value = "")
        df_final.insert(0, 'IT3', value = "")
        df_final.insert(0, 'IT4', value = "")
        df_final.insert(0, 'IT5', value = "")
        df_final.insert(0, 'Mes (Original)', value = "")
        df_final.insert(0, 'Data (Original)', value = "")
        df_final.insert(0, 'Mes', value = "")
        df_final.insert(0, 'Data', value = "")
        df_final.insert(0, 'PainelAPD', value = "APD")
        #df_final['Payment Amount'] = df_final['Payment Amount'].apply(lambda x: str(x).replace(".", ""))
        abertura_excel = io.BytesIO()
        with open(caminho_idpainel+"PARAMETROS RODAR PAINEL.xlsb", 'rb') as abertura:
            excel = msoffcrypto.OfficeFile(abertura)
            excel.load_key(('mini2019'))
            excel.decrypt(abertura_excel)
        dfgateway = pd.read_excel(abertura_excel, sheet_name='ID Gateway')
        dfpainel = pd.read_excel(abertura_excel, sheet_name='Tipo Painel')
        dfgateway = dfgateway.rename(columns={'id_gateway': 'Id Gateway'})
        dfgateway = dfgateway.rename(columns={'Código': 'Codigo'})
        dfgateway = dfgateway[['Id Gateway','Codigo']]
        dfpainel = dfpainel.rename(columns={'Código': 'Agente/Pais'})
        dfpainel = dfpainel[['Agente/Pais','Tipo Painel']]
        df_final.insert(0, 'Agente/Pais', value = "")
        #df_final['Country'] = df_final['Country'].apply(lambda x: str(x).replace('BR', 'CO'))
        df_final['Agente/Pais'] = df_final['Retain Agent'] + ' ' + '[' + df_final['Country'] +']'
        df_final = pd.merge(df_final, dfgateway, on='Id Gateway', how='left')
        df_final = pd.merge(df_final, dfpainel, on='Agente/Pais', how='left')
        df_final['Codigo'] = df_final['Codigo'].apply(lambda x: str(x).replace('nan', 'ID Não Localizada - Metodo Novo'))
        df_final['Codigo'] = df_final['Codigo'].apply(lambda x: str(x).replace('FerID Não Localizada - Metodo Novo', 'Fernan'))
        df_final['Tipo Painel'] = df_final['Tipo Painel'].apply(lambda x: str(x).replace('nan', 'Sem Painel'))
        df_final['Tipo Painel'] = np.where((df_final['Country'].str.contains('BR')) & (df_final['Tipo Painel'].str.contains('Sem Painel')),'Base Paineis' , df_final['Tipo Painel']) 
        df_final['Tipo Painel'] = np.where((df_final['Country'].str.contains('CL')) & (df_final['Tipo Painel'].str.contains('Sem Painel')),'Base Paineis Chile' , df_final['Tipo Painel'])
        df_final['Tipo Painel'] = np.where((df_final['Country'].str.contains('IN')) & (df_final['Tipo Painel'].str.contains('Sem Painel')),'Base Paineis India' , df_final['Tipo Painel'])
        df_final['Tipo Painel'] = np.where((df_final['Country'] != 'BR')& (df_final['Country'] != 'IN') & 
        (df_final['Country'] != 'CL') & (df_final['Tipo Painel'].str.contains('Sem Painel')),'Base Paineis Mundo (menos Brasil)' , df_final['Tipo Painel'])
        files_it = glob.glob(caminho_it+'*.csv')
        for file in files_it:
            filename2 = file.split('\\')[10].replace('.csv','')
            print('Lendo arquivo it '+filename2)
            df = pd.read_csv(file, encoding='ISO-8859-1', low_memory=False)
            df = df[['Id Payment','Deposit Info','Gateway Reference','Bank Account Name','Bank Account Branch','Bank Account Number','Match Amount','Match Type']]
            df = df.rename(columns={'Id Payment': 'IdPayment'})
            df_final = pd.merge(df_final, df, on='IdPayment', how='left')
            df_final['Bank Account Branch'] = df_final['Bank Account Branch'].apply(lambda x: str(x).replace('nan', ''))
            df_final['Bank Account Number'] = df_final['Bank Account Number'].apply(lambda x: str(x).replace('nan', ''))
            df_final['Client Name'] = df_final['Client Name'].apply(lambda x: str(x).replace('-', ''))
            #print(df_final['IdPayment'].dtypes)
            #df_final = df_final[df_final['IdPayment'] == 446174830]
            #print(df_final)
            df_final['IT'] = df_final['Bank Account Branch'] + ' / ' + df_final['Bank Account Number']
            df_final['IT2'] = df_final['Deposit Info']
            df_final['IT3'] = df_final['Gateway Reference']
            df_final['IT4'] = df_final['Bank Account Name']
            df_final['IT5'] = df_final['Match Amount'] + ' / ' + df_final['Match Type']
            data = pd.DatetimeIndex(df_final['Last Change Date'])
            data2 = pd.DatetimeIndex(df_final['Last Change Date']) + pd.DateOffset(hours=-3)
            df_final['Mes (Original)'] = data.month
            df_final['Data (Original)'] = data.date
            df_final['Data'] = data2.date
            df_final['Mes'] = data2.month
            df_final = df_final[['Mes (Original)','Data (Original)','Mes','Data','PainelAPD','Idboleto','IdPayment','Client Document','Client Name','Payment Method Name','Gateway Name','Retain Agent','Payment Amount','Sent Amount','Creation Date',
            'Last Change Date','Merchant Name','Amount (USD)','Status','Business Unit','External Merchant Id','IT','IT2','IT3','IT4','IT5','Currency','External Id','Codigo','Client Id',
            'Country','User Amount (local)','Test','Id Gateway','Tipo Painel']]
            df_final = df_final.sort_values(by=['Last Change Date','Creation Date'])
            df_auxiliar = pd.DataFrame()
            df_auxiliar = pd.concat([df_final, df_auxiliar])
            df_auxiliar.insert(0, 'Agente/Pais', value = "")
            df_auxiliar['Agente/Pais'] = df_auxiliar['Retain Agent'] + ' ' + '[' + df_auxiliar['Country'] +']'
            df_auxiliar = df_auxiliar.drop_duplicates(subset=['Agente/Pais'], keep='first')
            df_auxiliar = df_auxiliar[['Agente/Pais','Tipo Painel']]
            lista_auxiliar = df_auxiliar.drop_duplicates(subset=['Tipo Painel'], keep='first')
            lista_auxiliar = lista_auxiliar.sort_values(by=['Tipo Painel'])
            lista_auxiliar = list(lista_auxiliar['Tipo Painel'])
            print('Salvando arquivo auxiliar "Paineis Final"')
            df_auxiliar.to_excel(caminho_final+'Paineis Final.xlsx', index = False)
            df_auxiliar = pd.DataFrame()
            print('Iniciando Loop Paineis')
            for painel in lista_auxiliar:
                print('Salvando '+painel)
                df_auxiliar = df_final[df_final['Tipo Painel'] == painel]
                df_auxiliar.to_csv(caminho_final+mes+'.'+ano+' - '+painel+'- CSV'+'.csv', encoding='ISO-8859-1', index = False)
            print('Ajustando arquivo de Check')
            wb = xw.Book(caminho_personal)
            wb2 = wb2 = xw.Book(caminho_final+mes+'.'+ano+' - Base Paineis- CSV'+'.csv')
            tratar = wb.macro('Python.TratarArquivoCheckPainel')
            tratar()
            for painel in lista_auxiliar:
                print('Colando arquivo em '+painel)
                wb = xw.Book(caminho_personal)
                wb2 = xw.Book(caminho_final+mes+'.'+ano+' - '+painel+'- CSV'+'.csv')
                tratar = wb.macro('Python.TratarArquivoPainelPython')
                tratar()
                wb2.close()
            wb.app.quit()
else:
    print('Existem mais ou menos do que um arquivo csv nas pastas Paineis e IT')
time.sleep(5)

