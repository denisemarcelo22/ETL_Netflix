import pandas as pd
import os # UTILIZADO PARA MANIPULAR OBJETOS DO SISTEMA OPERACIONAL (EX.: ARQUIVOS)
import glob # LER ARQUIVOS EM MASSA (LER TODOS OS ARQUIVOS QUE ESTÃO DENTRO DE UMA PASTA) | BIBLIOTECA NATIVA

# CAMINHO PARA LER OS ARQUIVOS (RAW)
folder_path = 'src\\data\\raw'

# LISTAR TODOS OS ARQUIVOS DE EXCEL DE FORMA DINÂMICA
excel_files = glob.glob(os.path.join(folder_path,'*.xlsx'))

# CONDIÇÃO CASO NENHUM ARQUIVO SEJA ENCONTRADO
if not excel_files:
    print("Nenhum arquivo compatível com a extensão .xlsx encontrado")
else:
    
    df = []
    
    for excel_file in excel_files:

      try:
           # SALVANDO O CONTEÚDO NO DATAFRAME TEMPORÁRIO
           df_temp = pd.read_excel(excel_file)

           # PEGAR O NOME DO ARQUIVO PARA POSTERIORMENTE USAR COMO INFORMAÇÃO EM UMA NOVA COLUNA
           file_name = os.path.basename(excel_file)

           # PEGAR O ANO E O MÊS DO NOME DO ARQUIVO
           

           # CRIAR UMA NOVA COLUNA PARA PGEAR O NOME DO ARQUIVO DE ONDE VEM A INFORMAÇÃO INICIAL
           df_temp['filename'] = file_name
           
           if 'brasil' in file_name.lower():
              df_temp['location'] = 'br' #CRIANDO A COLUNA LOCATION
           elif 'france' in file_name.lower():
              df_temp['location'] = 'fr'
           elif 'italian' in file_name.lower():
              df_temp['location'] = 'it'

           # CRIAR NOVA COLUNA PARA O NOME DA CAMPANHA
           df_temp['campaign'] = df_temp['utm_link'].str.extract(r'utm_campaign=(.*)')
           
           # JOGANDO TODOS OS DADOS JÁ TRATADOS DENTRO DO DATA FRAME FINAL 
           df.append(df_temp)
        
      except Exception as e:
        print(f"Erro ao ler o arquivo {excel_file}:{e}")

if df:
   
   # CONCATENA TODAS AS TEBELAS SALVAS NO 'DF' EM UMA ÚNICA TABELA
   result = pd.concat(df, ignore_index=True)

   # CAMINHO DE SAÍDA
   output_file = os.path.join('src','data','ready', 'clean.xlsx')

   # CRIANDO VARIÁVEL PARA ESCREVER OS DADOS EM UM ARQUIVO EXCEL (MOTOR DE ESCRITA)
   writer = pd.ExcelWriter(output_file, engine='xlsxwriter') # MUDANDO A ENGINE PARA USAR A BIBLIOTECA QUE FOI IMPORTADA, CASO NÃO COLOQUE ESSA INFORMAÇÃO SERÁ USADO A ENGINE DO PANDAS COMO PADRÃO

   # LEVA OS DADOS DO RESULTADO A SEREM ESCRITOS NO MOTOR DE EXCEL CONFIGURADO
   result.to_excel(writer, index=False)

   # SALVA O ARQUIVO DE EXCEL
   writer._save()
else:
   print("nenhum dado para ser salvo")