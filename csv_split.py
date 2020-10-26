import pandas as pd
import xlsxwriter
import xlrd

path_base_ativa = r'C:\Users\Jobs Tableau\BaseAtiva.csv'

chunk_size = 330000
batch_no = 1

for chunk in pd.read_csv(path_base_ativa, sep=";", encoding='utf-8', chunksize=chunk_size):
    novo_df_base_ativa = pd.DataFrame(data=chunk)
    novo_df_base_ativa['NU_INFORMACAOCNJ_ACO'] = novo_df_base_ativa['NU_INFORMACAOCNJ_ACO'].astype(str).str.rjust(width=20, fillchar='0')
    novo_df_base_ativa['NU_INFORMACAO_GE_ACO'] = novo_df_base_ativa['NU_INFORMACAO_GE_ACO'].astype(str).str.rjust(width=20, fillchar='0')
    chunk = novo_df_base_ativa
    chunk.to_excel('BASE_ATIVA_PT_' + str(batch_no) +'.xlsx', index=False, engine='xlsxwriter')
    batch_no += 1
