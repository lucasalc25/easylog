import pandas as pd
from pathlib import Path
import openpyxl
import re

# Função para ler os contatos de um arquivo TXT
def ler_registros(caminho_planilha):
    try:
        # Lê a planilha usando pandas
        df = pd.read_excel(caminho_planilha)
        
        # Certifique-se de que os nomes das colunas estão corretos
        alunos = []
        for _, row in df.iterrows():
            alunos.append({
                "Aluno": row["Aluno"],         # Nome do aluno
                "Observacao": row["Observação"]
            })

        return alunos
        
    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")
        return []
    
# Função para ler os contatos de um arquivo TXT
def ler_faltosos_mes(caminho_planilha):
    try:
        # Lê a planilha usando pandas
        df = pd.read_excel(caminho_planilha)
        
        # Certifique-se de que os nomes das colunas estão corretos
        alunos = []
        for _, row in df.iterrows():
            alunos.append({
                "Aluno": row["Aluno"],         # Nome do aluno
                "Contato": row["Celular"]
            })

        return alunos
        
    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")
        return []
    
# Função para ler os contatos de um arquivo TXT
def ler_faltosos_dia(caminho_planilha):
    try:
        # Lê a planilha usando pandas
        df = pd.read_excel(caminho_planilha)

        # Certifique-se de que os nomes das colunas estão corretos
        faltosos = []
        for _, row in df.iterrows():
            faltosos.append({
                "Aluno": row["Aluno"],
                "Observacao": row["Observação"],              # Nome do aluno
                "Educador": row["Educador"],        # Nome do Educador
                "Contato": str(row["Celular"])     # Contato do aluno
            })
        
        for aluno in faltosos:
            aluno['Contato'] = formatar_telefone(aluno['Contato'])

        return faltosos
    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")
        return []
    
# Função para formatar o número de telefone
def formatar_telefone(celular):
    # Usando expressão regular para garantir que o número tenha o formato correto
    if celular:
        celular = re.sub(r'\D', '', celular)  # Remove qualquer caractere não numérico
        celular.replace("(", "").replace(")", "").replace(" ", "")
        if len(celular) == 11:  # Verifica se o número tem 11 dígitos
            return f"{celular[:2]}{celular[2]}{celular[3:7]}{celular[7:]}"
    return celular  # Retorna o número original caso não tenha o formato esperado

def filtrar_faltosos_do_dia(planilha_faltosos_do_dia):
    planilha_educadores = Path.home() / "Documents" / "EasyLog" / "Data" / "alunos_e_educadores.xls"

    df_faltosos = pd.read_excel(planilha_faltosos_do_dia, sheet_name='Sheet', header=3)
    df_educadores = pd.read_excel(planilha_educadores, sheet_name='Sheet')

    df_educadores.rename(columns={"Nome Aluno": "Aluno"}, inplace=True)

    # Manter apenas o primeiro nome do educador
    df_educadores['Educador'] = df_educadores['Educador'].str.split().str[0]

    # Selecionar apenas as colunas necessárias
    df_faltosos = df_faltosos[['Contrato', 'Aluno', 'Observação', 'Tel Residencial', 'Celular']]
    df_educadores = df_educadores[['Aluno', 'Educador']]

    # Mesclar as duas planilhas com base na coluna 'aluno'
    df_mesclado = pd.merge(df_faltosos, df_educadores, on='Aluno', how='left')

    # Remover linhas duplicadas
    df_mesclado = df_mesclado.drop_duplicates(subset=['Celular'])

    # Transferir telefones residenciais para a coluna celular quando não houver
    df_mesclado['Celular'] = df_mesclado['Celular'].fillna(df_mesclado['Tel Residencial'])

    # Remover a coluna tel_residencial
    df_mesclado = df_mesclado.drop(columns=['Tel Residencial'])

    df_mesclado = df_mesclado[df_mesclado['Educador'].notna() & (df_mesclado['Educador'] != '')]

    # Deixar a coluna observação vazia
    df_mesclado['Observação'] = ''

    # Selecionar as colunas finais
    df_final = df_mesclado[['Contrato', 'Aluno', 'Observação', 'Educador', 'Celular']]

    # Adicionar "Funcionário" na coluna "Observação" onde a coluna "Educador" contém "Sangela"
    df_final.loc[df_final['Educador'].str.contains('Sangela', na=False), 'Observação'] = 'Funcionário'

    caminho_saida = Path.home() / "Documents" / "EasyLog" / "Data" / "faltosos_do_dia_filtrados.xlsx"

    # Salvar o resultado
    df_final.to_excel(caminho_saida, index=False)

    ajustar_largura_colunas(caminho_saida)
    

def filtrar_faltosos_do_mes(planilha_faltosos_do_mes):
    planilha_celulares = Path.home() / "Documents" / "EasyLog" / "Data" / "celulares_faltosos_do_mes.xls"

    df_celulares_faltosos_do_mes = pd.read_excel(planilha_celulares, sheet_name='Sheet', header=3)
    df_faltosos_do_mes = pd.read_excel(planilha_faltosos_do_mes, sheet_name='Sheet')

    df_faltosos_do_mes.rename(columns={"Nome Aluno": "Aluno"}, inplace=True)

    colunas_removidas = [2, 6, 10, 11]  # Ajuste os índices conforme necessário
    df_celulares_faltosos_do_mes = df_celulares_faltosos_do_mes.drop(df_celulares_faltosos_do_mes.columns[colunas_removidas], axis=1, errors='ignore')

    # Remove as linhas de 1 a 3 (0-indexadas)
    df_celulares_faltosos_do_mes = df_celulares_faltosos_do_mes.drop(index=[0, 1, 2])

    # Remove linhas em branco até encontrar 4 linhas em branco seguidas
    while True:
        # Encontra o índice da primeira linha em branco
        blank_index = df_celulares_faltosos_do_mes[df_celulares_faltosos_do_mes.isnull().all(axis=1)].index
        if blank_index.empty:
            break  # Se não houver linhas em branco, sai do loop

        # Remove a linha em branco e as próximas 3 linhas
        start_index = blank_index[0]
        end_index = start_index + 4
        df_celulares_faltosos_do_mes = df_celulares_faltosos_do_mes.drop(index=range(start_index, min(end_index, len(df_celulares_faltosos_do_mes))))

        # Verifica se há mais de 4 linhas em branco seguidas
        if len(df_celulares_faltosos_do_mes[df_celulares_faltosos_do_mes.isnull().all(axis=1)]) > 4:
            break

    # Manter apenas o primeiro nome do educador
    df_faltosos_do_mes['Educador'] = df_faltosos_do_mes['Educador'].str.split().str[0]

    # Remove a linha 2
    df_faltosos_do_mes = df_faltosos_do_mes.drop(index=0)

    # Selecionar apenas as colunas necessárias
    df_celulares_faltosos_do_mes = df_celulares_faltosos_do_mes[['Contrato', 'Aluno', 'Tel Residencial', 'Celular']]
    df_faltosos_do_mes = df_faltosos_do_mes[['Aluno', 'Educador', 'Data Cadastro Contrato', 'Faltas', 'Reposições']]

    # Remove as linhas onde o valor da coluna "faltas" é igual ao valor da coluna "reposição"
    df_faltosos_do_mes = df_faltosos_do_mes[df_faltosos_do_mes['Faltas'] != df_faltosos_do_mes['Reposições']]

    # Mesclar as duas planilhas com base na coluna 'aluno'
    df_mesclado = pd.merge(df_faltosos_do_mes, df_celulares_faltosos_do_mes, on='Aluno', how='left')

    # Remover linhas duplicadas
    df_mesclado = df_mesclado.drop_duplicates(subset=['Celular'])

     # Transferir telefones residenciais para a coluna celular quando não houver
    df_mesclado['Celular'] = df_mesclado['Celular'].fillna(df_mesclado['Tel Residencial'])

    # Remover a coluna tel_residencial
    df_mesclado = df_mesclado.drop(columns=['Tel Residencial'])

    # Ordena o DataFrame pela coluna "Aluno" em ordem alfabética (A a Z)
    df_mesclado = df_mesclado.sort_values(by='Aluno')

    # Selecionar as colunas finais
    df_final = df_mesclado[['Contrato','Aluno', 'Educador', 'Data Cadastro Contrato', 'Celular', 'Faltas', 'Reposições']]

    caminho_saida = Path.home() / "Documents" / "EasyLog" / "Data" / "faltosos_do_mes_filtrados.xlsx"

    # Salva o DataFrame processado em um novo arquivo Excel
    df_final.to_excel(caminho_saida, index=False)

    ajustar_largura_colunas(caminho_saida)

def ajustar_largura_colunas(arquivo):
    # Carregar o arquivo Excel
    wb = openpyxl.load_workbook(arquivo)
    ws = wb.active

    if arquivo == Path.home() / "Documents" / "EasyLog" / "Data" / "faltosos_do_dia_filtrados.xlsx":
        # Definir larguras específicas para as colunas
        colunas_largura = {
            "A": 10.0,
            "B": 40.0,
            "C": 50.0,
            "D": 12.0,
            "E": 18.0
        }
    elif arquivo == Path.home() / "Documents" / "EasyLog" / "Data" / "faltosos_do_mes_filtrados.xlsx":
        # Definir larguras específicas para as colunas
        colunas_largura = {
            "A": 10.5,
            "B": 40.5,
            "C": 10.0,
            "D": 21.0,
            "E": 14.0,
            "F": 7.0,
            "E": 11.0
        }
        

    # Ajustar a largura das colunas conforme o dicionário
    for coluna, largura in colunas_largura.items():
        ws.column_dimensions[coluna].width = largura

    # Salvar o arquivo depois de ajustar as larguras
    wb.save(arquivo)

def gerar_planilha(tipo, campo_data_inicial, campo_data_final, campo_filtro_educador):
    from bot import gerar_faltosos_do_dia, gerar_faltosos_do_mes

    if tipo == "faltas_do_dia":
        gerar_faltosos_do_dia(campo_data_inicial, campo_data_final, campo_filtro_educador)
    elif tipo == "faltas_do_mes":
        gerar_faltosos_do_mes(campo_data_inicial, campo_data_final)