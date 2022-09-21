from ast import pattern
from calendar import month_abbr
from ctypes import sizeof
import pandas as pd
from project_func import file_reader
import numpy as np
from datetime import datetime, date
import time
from os import listdir
import re


def main():
    #Get all Files on Folder
    filenames = []
    path = 'D:\\Python Projects\\bank analysis' #Use path to your folder
    for file in listdir(path):
        if file.endswith('.csv'):
            filenames.append(file)

    #Read files
    df_list = file_reader(filenames)
    df = pd.concat(df_list, ignore_index=True)

    #Transform Date Strings to Date type
    df['Data'] = pd.to_datetime(df['Data'], dayfirst=True, infer_datetime_format=True)
    df['Data'] = df['Data'].apply(lambda x: x.date())
    
    #Drop extra column
    df.drop(axis=1, columns='Unnamed: 6', inplace=True)

    #Create New Columns for Classification
    df['Classificação'] = np.nan
    df['Descrição'] = np.nan

    #Select Start of the Month Vals
    month_start_vals = df.loc[df['Histórico'] == 'Saldo Anterior']
    month_start_vals = month_start_vals[['Data', 'Valor']].values

    #Select End of the Month Indexes and Vals
    month_end_vals = df.loc[df['Histórico'] == 'S A L D O']
    month_end_vals = month_end_vals[['Data', 'Valor']].values

    #DataFrame for Month Summary
    gastos_df = pd.DataFrame(columns = ['Mês', 'Saldo Inicial', 'Saldo Final', '+/-','Salário', 'Maior Gasto'])

    #Append Results
    for index, val in enumerate(month_start_vals):
        month_e = date.strftime(month_end_vals[index,0], '%B')
        gastos_df.loc[len(gastos_df.index)+1, :] = [month_e, val[1], month_end_vals[index][1], month_end_vals[index][1] - val[1], None, None]

    #Free Memory
    del month_end_vals, month_start_vals, month_e

    #Remove Start and End of the Month Vals from Main DF
    df = (df[(df['Histórico'] != 'Saldo Anterior') & (df['Histórico'] != 'S A L D O')]).reset_index(drop=True)

    #Split into Money Spent and Received
    val_in = df[df['Valor']>=0]
    val_out = df[df['Valor']<0]

    #Find Biggest Expense of the Month
    for val in gastos_df['Mês'].items():
        month_num = datetime.strptime(val[1], '%B').month
        pattern = re.compile(f'2022-0{month_num}-\d')
        pattern_list = []
        for i, v in df['Data'].items():
            p = pattern.search(str(v))
            if p != None:
                pattern_list.append(df['Valor'].iloc[i])
        max_val = min(pattern_list)
        gastos_df.loc[val[0],['Maior Gasto']] = max_val

    #DataFrames to Excel
    pix = pd.DataFrame(columns= ['Data', 'Valor', 'Descrição'])
    seguros = pd.DataFrame(columns= ['Data', 'Valor', 'Descrição'])
    faturas = pd.DataFrame(columns= ['Data', 'Valor', 'Descrição'])
    cartao_credito = pd.DataFrame(columns= ['Mês','Data', 'Valor', 'Descrição'])
    
    for i, row in val_out['Histórico'].items():
        #Pix e Subdivisões
        if 'PIX' in row.upper():
            #Clean 'Histórico'
            pattern = re.compile(r'\d{2}(:\d{2})')
            p = pattern.search(row)
            descr = row[p.span()[1]+1:]
            #Append values to DFs
            df.loc[i,['Descrição']] = f'Pix - {descr}'
            pix.loc[len(pix.index)] = [df.loc[i,['Data']].values[0], df.loc[i,['Valor']].values[0], f'Pix - {descr}']
        #Seguros
        if 'PLANO DE SAÚDE' in row.upper():
            df.loc[i,['Descrição']] = 'PLANO DE SAÚDE'
            seguros.loc[len(seguros.index)] = [df.loc[i,['Data']].values[0], df.loc[i,['Valor']].values[0], 'PLANO DE SAÚDE']
        #Contas
        if 'FATURA DE GÁS' in row.upper():
            df.loc[i,['Descrição']] = 'FATURA DE GÁS'
            faturas.loc[len(faturas.index)] = [df.loc[i,['Data']].values[0], df.loc[i,['Valor']].values[0], 'FATURA DE GÁS']
        if 'ENERGIA ELÉTRICA' in row.upper() or 'CONTA LUZ' in row.upper():
            if 'CPFL' in row.upper():
                df.loc[i,['Descrição']] = 'ENERGIA ELÉTRICA S'
                faturas.loc[len(faturas.index)] = [df.loc[i,['Data']].values[0], df.loc[i,['Valor']].values[0], 'ENERGIA ELÉTRICA S']
            elif abs(val_out.loc[i,['Valor']].values[0]) >= 80.00:
                df.loc[i,['Descrição']] = 'ENERGIA ELÉTRICA I'
                faturas.loc[len(faturas.index)] = [df.loc[i,['Data']].values[0], df.loc[i,['Valor']].values[0], 'ENERGIA ELÉTRICA I']
            else:
                df.loc[i,['Descrição']] = 'ENERGIA ELÉTRICA H'
                faturas.loc[len(faturas.index)] = [df.loc[i,['Data']].values[0], df.loc[i,['Valor']].values[0], 'ENERGIA ELÉTRICA H']
        if 'TELEFONE' in row.upper():
            if 'VIVO' in row.upper():
                df.loc[i,['Descrição']] = 'TELEFONE VIVO'
                faturas.loc[len(faturas.index)] = [df.loc[i,['Data']].values[0], df.loc[i,['Valor']].values[0], 'TELEFONE V']
            elif 'CLARO' in row.upper():
                df.loc[i,['Descrição']] = 'TELEFONE CLARO'
                faturas.loc[len(faturas.index)] = [df.loc[i,['Data']].values[0], df.loc[i,['Valor']].values[0], 'TELEFONE C']
            else:
                df.loc[i,['Descrição']] = 'TELEFONE'
                faturas.loc[len(faturas.index)] = [df.loc[i,['Data']].values[0], df.loc[i,['Valor']].values[0], 'TELEFONE']
        if 'CARTÃO CRÉDITO' in row.upper():
            month = df.loc[i,['Data']].values[0].strftime('%B')
            df.loc[i,['Descrição']] = 'CARTÃO DE CRÉDITO'
            cartao_credito.loc[len(cartao_credito.index)] = [month, df.loc[i,['Data']].values[0], df.loc[i,['Valor']].values[0], 'CARTÃO DE CRÉDITO']

    faturas = faturas.sort_values(by=['Descrição','Data'], ignore_index=True)

    #Writing Data to Excel
    with pd.ExcelWriter('Análise.xlsx', engine='xlsxwriter', date_format='YYYY-MM-DD') as writer:
        df.to_excel(writer, 'Extratos')
        gastos_df.to_excel(writer, 'Resumo Mensal')
        pix.to_excel(writer, 'Pix')
        faturas.to_excel(writer, 'Faturas')
        seguros.to_excel(writer, 'Seguros')
        cartao_credito.to_excel(writer, 'Cartão de Crédito')

    
    

if __name__ == '__main__':
    main()
