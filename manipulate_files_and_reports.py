import time
from datetime import datetime, date
import pandas as pd
import os, shutil
import os.path
import glob
import xlrd
import openpyxl
import zipfile

########### First step: Prepare the suppliers data sheet
path_final_fornecedores = r'\Arquivos_Referencia\FORNECEDORES.xlsx'
if os.path.exists(path_final_fornecedores):
    os.remove(path_final_fornecedores)
else:
    print("File does not exist")
time.sleep(2)

## Copy a version of it to another directory so we can manipulate it and others can still use the original file
path_origem_fornecedores = r'\Origem\FORNECEDORES.xlsx'
shutil.copy(path_origem_fornecedores, path_final_fornecedores)
time.sleep(2)

## Converting explicitly into DataFrame
df_fornecedores = pd.read_excel(path_final_fornecedores)

## Format the sheet
new_header = df_fornecedores.iloc[1] 
df_fornecedores = df_fornecedores.rename(columns = new_header)
df_fornecedores = df_fornecedores.drop(df_fornecedores.index[[0,1]])

## Export the new file just to check if it's alright
df_fornecedores.to_excel(r'\fornecedores_Tratado.xlsx', index = None, engine = 'xlsxwriter')
fornecedores_Tratado = r'\Arquivos_Referencia\fornecedores_Tratado.xlsx'

########### Second step: prepare the report file

## Converting explicitly into DataFrame
path_origem_relatorio_documentos = r'\Arquivos_Referencia\Solicitações.xlsx'
df_relatorio_documentos = pd.read_excel(path_origem_relatorio_documentos)

## Format the sheet
new_header = df_relatorio_documentos.iloc[5] 
df_relatorio_documentos = df_relatorio_documentos.rename(columns = new_header)
df_relatorio_documentos = df_relatorio_documentos.drop(df_relatorio_documentos.index[[0,5]])
df_relatorio_documentos = df_relatorio_documentos.astype(str)
df_relatorio_documentos = df_relatorio_documentos[(df_relatorio_documentos['Assunto'] == 'INFORMAÇÕES_B') | (df_relatorio_documentos['Assunto'] == 'INFORMAÇÕES_A') ]

## Format the sheet
df_relatorio_documentos['Período Final'] = df_relatorio_documentos['Período Final'].str[3:5] +'/'+ df_relatorio_documentos['Período Final'].str[0:2] + '/'+ df_relatorio_documentos['Período Final'].str[6:10]
df_relatorio_documentos['Período Inicial'] = df_relatorio_documentos['Período Inicial'].str[3:5] +'/'+ df_relatorio_documentos['Período Inicial'].str[0:2] + '/'+ df_relatorio_documentos['Período Inicial'].str[6:10]
df_relatorio_documentos['Data Criação'] = df_relatorio_documentos['Data Criação'].str[3:5] +'/'+ df_relatorio_documentos['Data Criação'].str[0:2] + '/'+ df_relatorio_documentos['Data Criação'].str[6:10]
df_relatorio_documentos['Data Finalização'] = df_relatorio_documentos['Data Finalização'].str[3:5] +'/'+ df_relatorio_documentos['Data Finalização'].str[0:2] + '/'+ df_relatorio_documentos['Data Finalização'].str[6:10]
df_relatorio_documentos['Data Início'] = df_relatorio_documentos['Data Início'].str[3:5] +'/'+ df_relatorio_documentos['Data Início'].str[0:2] + '/'+ df_relatorio_documentos['Data Início'].str[6:10]
df_relatorio_documentos['Assunto'].replace({"INFORMAÇÕES_B": "RELCAN", "INFORMAÇÕES_A": "RELINC"}, inplace=True)

## Export new report
df_relatorio_documentos.to_excel(r'\Arquivos_Referencia\Relatorio_documentos_Tratado.xlsx', index = None, engine = 'xlsxwriter')
Relatorio_documentos_Tratado = r'\Arquivos_Referencia\Relatorio_documentos_Tratado.xlsx'

########### Third step: copy pdf files
path_origem_documentos_concluidos = r'\Concluido\DOCUMENTOS OK'
## Filter the ones we want
path_arquivos_sem_renomear = r'\Arquivos_Sem_Renomear'
path_ja_enviados_parceria = r'\Concluido\DOCUMENTOS OK\Ja_Enviados'

df_fornecedores = pd.read_excel(fornecedores_Tratado)
def busca_df (df_do_fornecedor, fornecedor, coluna):
    for i in range(0, len(df_do_fornecedor.index)):
        if df_do_fornecedor.iat[i, coluna] == fornecedor:
            return True
    return False

for file in os.listdir(path_origem_documentos_concluidos):
    df_zerado = df_fornecedores
    if file.endswith(".pdf"):
        file_com_pdf = file
        file_sem_pdf = file.replace(".pdf", "")
        file_sem_pdf = file_sem_pdf.split('_')
        ## Each split element gives us a part of data which we're gonna use later
        operacao = file_sem_pdf[0]
        operacao = str(operacao)
        num_fornecedor = str(file_sem_pdf[1].strip())
        cod_cidade = str(file_sem_pdf[2].strip())
        fornecedor = num_fornecedor+' '+cod_cidade
        if operacao in ['RELINC', 'RELCAN']:
            ## Filter DataFrame
            df_do_fornecedor = df_zerado[(df_zerado.Situação == 'Ativo') & (df_zerado.RESPONSAVEL == 'COMPANY_NAME')]

            if busca_df(df_do_fornecedor, fornecedor, coluna = 0):
                ## Filter the DataFrame just to the single row of information we want
                df_do_fornecedor = df_do_fornecedor[ df_do_fornecedor.FORNECEDOR == fornecedor ]

                ## Check if file already exists
                if os.path.exists(os.path.join (path_arquivos_sem_renomear, file) ):
                    print("File already exists here")
                    ## Move the files to another folder, so the next time they don't repeat
                else:
                    shutil.copy(os.path.join (path_origem_documentos_concluidos,file) , path_arquivos_sem_renomear)
                    time.sleep(1)
                    
                    if os.path.exists(os.path.join (path_ja_enviados_parceria, file) ):
                        os.remove(os.path.join (path_ja_enviados_parceria, file) )
                        time.sleep(1)
                        shutil.move(os.path.join (path_origem_documentos_concluidos,file) , path_ja_enviados_parceria)
        else:
            print("Not part of the process")

df_relatorio_documentos = pd.read_excel(Relatorio_documentos_Tratado)

path_arquivos_renomeados = r'\Arquivos_Renomeados'

## Move files
for file in os.listdir(path_arquivos_sem_renomear):
    if not os.path.exists(os.path.join (path_arquivos_renomeados,file)):
        if file.endswith(".pdf"):
            shutil.copy(os.path.join (path_arquivos_sem_renomear,file) , path_arquivos_renomeados)
    else:
        print("File has already been moved: \n")
time.sleep(5)
## Rename the move files
for file in os.listdir(path_arquivos_renomeados):
    df_zerado = df_relatorio_documentos
    if file.endswith(".pdf"):
        file_com_pdf = file
        file_sem_pdf = file.replace(".pdf", "")
        file_sem_extensao_pdf = file.replace(".pdf", "")
        file_sem_pdf = file_sem_pdf.split('_')
        if "key" not in file_sem_extensao_pdf:
            operacao = file_sem_pdf[0]
            operacao = str(operacao)
            num_fornecedor = str(file_sem_pdf[1].strip())
            cod_cidade = str(file_sem_pdf[2].strip())
            fornecedor = num_fornecedor+' '+cod_cidade
            dataInicio = file_sem_pdf[4]
            dataInicioTratada = dataInicio[0:2] + '/' + dataInicio[2:4] + '/' + dataInicio[4:8]   
            dataFim = file_sem_pdf[5]
            dataFimTratada = dataFim[0:2] + '/' + dataFim[2:4] + '/' + dataFim[4:8]   
            df_arquivo = df_zerado
            ## Filter DataFrame for specific data
            df_arquivo.columns = [c.replace(' ', '_') for c in df_arquivo.columns]
            df_arquivo = df_arquivo[(df_arquivo.Cidade == cod_cidade) & (df_arquivo.Assunto == operacao) & (df_arquivo.Período_Inicial == dataInicioTratada) & (df_arquivo.Período_Final == dataFimTratada) & (df_arquivo.Número_FORNECEDOR == num_fornecedor)]
            
            df_arquivo.to_excel(r'\Arquivos_Referencia\DataFrame_Logo_Apos_Filtro.xlsx', index = None, engine = 'xlsxwriter')
            arquivo_concatenado = operacao+'|'+num_fornecedor+'|'+cod_cidade+'|'+dataInicioTratada+'|'+dataFimTratada
            time.sleep(3)
            if df_arquivo.empty == True:
                df_arquivo.to_excel(r'\Arquivos_Referencia\DataFrame_considerado_Vazio.xlsx', index = None, engine = 'xlsxwriter')
                print("\n Empty DataFrame for: ", arquivo_concatenado)
                pass
            else:
                df_arquivo.to_excel(r'\Arquivos_Referencia\DataFrame_Caso_nao_seja_Vazio.xlsx', index = None, engine = 'xlsxwriter')
                print("df_arquivo: \n",df_arquivo['Cidade'], df_arquivo['Número_FORNECEDOR'])
                key_group_relatorio = str(df_arquivo.iat[0, 2])
                print("key_group_relatorio: \n",key_group_relatorio)
                time.sleep(3)
                ## Rename file with found key
                arquivo_renomeado_com_key_group = path_arquivos_renomeados+'/'+file_sem_extensao_pdf+'_'+key_group_relatorio+'_key'+'.pdf'
                print("\n\n Renamed File: arquivo_renomeado_com_key_group \n", arquivo_renomeado_com_key_group)
                os.rename((os.path.join (path_arquivos_renomeados,file)) , arquivo_renomeado_com_key_group)
                
########### Fourth step: Create a report for the found/formatted files and put it into a zip ###############
df_fornecedores = pd.read_excel(fornecedores_Tratado)

lista_files_com_pdf = []
lista_num_fornecedor = []
lista_cod_cidade = []
lista_uf = []
lista_dataInicioTratada = []
lista_dataFimTratada = []
lista_fornecedor = []
lista_keys_group = []
lista_fornecedor_parceria = []
lista_nome_cidade = []
lista_data_criacao_arquivo = []

for file in os.listdir(path_arquivos_renomeados):
    df_zerado = df_fornecedores
    if file.endswith(".pdf") and "key" in file and "sent" not in file:
        file_sem_pdf = file.replace(".pdf", "")
        file_sem_pdf = file_sem_pdf.split('_')

        operacao = str(file_sem_pdf[0].strip())
        num_fornecedor = str(file_sem_pdf[1].strip())
        cod_cidade = str(file_sem_pdf[2].strip())
        fornecedor = num_fornecedor+' '+cod_cidade
        df_do_fornecedor = df_zerado[(df_zerado.Situação == 'Ativo') & (df_zerado.RESPONSAVEL == 'COMPANY_NAME')]

        if busca_df(df_do_fornecedor, fornecedor, coluna = 0):
            df_do_fornecedor = df_do_fornecedor[ df_do_fornecedor.FORNECEDOR == fornecedor ]
            uf = file_sem_pdf[3]
            dataInicio = file_sem_pdf[4]
            dataInicioTratada = str(dataInicio[0:2] + '/' + dataInicio[2:4] + '/' + dataInicio[4:8])
            dataFim = file_sem_pdf[5]
            dataFimTratada = str(dataFim[0:2] + '/' + dataFim[2:4] + '/' + dataFim[4:8])
            key_group = file_sem_pdf[6]
            lista_files_com_pdf.append(file)
            lista_num_fornecedor.append(num_fornecedor)
            lista_cod_cidade.append(cod_cidade)
            lista_fornecedor.append(fornecedor)  
            lista_uf.append(uf)
            lista_dataInicioTratada.append(dataInicioTratada)
            lista_dataFimTratada.append(dataFimTratada)
            lista_keys_group.append(key_group)
                    
            ## After filtering, obtain the data columns we want
            cod_fornecedor_parceria = str(df_do_fornecedor.iat[0,12])
            cod_fornecedor_parceria = cod_fornecedor_parceria.strip()
            lista_fornecedor_parceria.append(cod_fornecedor_parceria)
            nome_cidade = str(df_do_fornecedor.iat[0,1])
            lista_nome_cidade.append(nome_cidade)
            print('File         :', os.path.join(path_arquivos_renomeados, file))
            print('Access time  :', time.ctime(os.path.getatime(os.path.join(path_arquivos_renomeados, file))))
            print('Modified time:', time.ctime(os.path.getmtime(os.path.join(path_arquivos_renomeados, file))))
            print('Change time  :', time.ctime(os.path.getctime(os.path.join(path_arquivos_renomeados, file))))
            print('Size         :', os.path.getsize(os.path.join(path_arquivos_renomeados, file)))

            caminho_arquivo_criacao = os.path.join(path_arquivos_renomeados, file)
            data_criacao_arquivo = (time.ctime(os.path.getctime(caminho_arquivo_criacao)))
            dia_criacao_arquivo = data_criacao_arquivo[8:10]
            mes_criacao_arquivo = data_criacao_arquivo[4:7]
            ano_criacao_arquivo = data_criacao_arquivo[20:24]
            dicionario_datas =  { 'Jan': '01', 'Fev': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06', 
                                    'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'  
                                }
            for mes in dicionario_datas:
                if mes == mes_criacao_arquivo:
                    mes_criacao_arquivo = dicionario_datas[mes]
                    data_tratada_criacao_arquivo = (dia_criacao_arquivo+'/'+mes_criacao_arquivo+'/'+ano_criacao_arquivo)
                    print(data_tratada_criacao_arquivo)
                    lista_data_criacao_arquivo.append(data_tratada_criacao_arquivo)   
        else:
            print("The supplier: ",fornecedor,"is not part of it")

dicionario_arquivos =  {'Nome_Arquivo': lista_files_com_pdf,
                        'Sigla_fornecedor': lista_fornecedor,
                        'Nome_Cidade': lista_nome_cidade,
                        'UF': lista_uf,
                        'Data_Inicio_Periodo': lista_dataInicioTratada,
                        'Data_Fim_Periodo': lista_dataFimTratada,
                        'Data_Criacao_Arquivo': lista_data_criacao_arquivo,
                        'keys_group': lista_keys_group,
                        'Codigo_fornecedor_parceria': lista_fornecedor_parceria
                        }
lista_para_df = pd.DataFrame.from_dict(dicionario_arquivos)
lista_para_df.to_excel(r'\Arquivos_Renomeados\CDP.K9999.lista_de_documentos.xlsx', index = None, engine = 'xlsxwriter')
lista_excel_documentos = r'\Arquivos_Renomeados\CDP.K9999.lista_de_documentos.xlsx'
time.sleep(3)
path_arquivos_enviar = r'\Arquivos_Enviar'

### Using get.date() to rename the zip
today = date.today()
d1 = today.strftime("%Y-%m-%d")
nome_para_o_zip = 'CDP.K9999.documentos'+'_'+d1+'.zip'
final_name_zip = os.path.join(path_arquivos_renomeados, nome_para_o_zip)
fantasy_zip = zipfile.ZipFile(final_name_zip, 'w')

### Erase files from folder
for file in os.listdir(path_arquivos_sem_renomear):
    if file.endswith(".pdf"):
        os.remove(os.path.join(path_arquivos_sem_renomear,file) )

### Create Zip file
for folder, subfolders, files in os.walk(path_arquivos_renomeados):
    for file in files:
        if ( (file.endswith('.pdf') and ('key' in file) and ('sent' not in file) ) or (file.endswith('.xlsx') ) ):
                fantasy_zip.write(os.path.join(folder, file), file, compress_type = zipfile.ZIP_DEFLATED)
fantasy_zip.close()
time.sleep(5)

### If zip exists copy it
if os.path.exists(final_name_zip):
    shutil.move(final_name_zip , path_arquivos_enviar)
else:
    print("zip not found")

### Renaming it just to find the ones sent through zip
time.sleep(3)
for file in os.listdir(path_arquivos_renomeados):
    if file.endswith(".pdf") and "key" in file and "sent" not in file:
        file_sem_extensao_pdf = file.replace(".pdf", "")
        nome_arquivo_pdf_pos_envio_zip = file_sem_extensao_pdf+'_sent.pdf'
        os.rename(os.path.join(path_arquivos_renomeados,file), os.path.join(path_arquivos_renomeados,nome_arquivo_pdf_pos_envio_zip))

for file in os.listdir(path_arquivos_renomeados):
    if file.endswith(".pdf"):
        os.remove(os.path.join(path_arquivos_renomeados,file) )
