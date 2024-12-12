import pandas as pd
from pathlib import Path
import openpyxl
import re

# Função para ler os contatos de um arquivo TXT
def ler_alunos(caminho_planilha):
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
def ler_contatos(caminho_planilha):
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

def filtrar_faltosos(planilha_faltosos):
    planilha_educadores = Path.home() / "Documents" / "EasyLog" / "Data" / "alunos_e_educadores.xls"

    df_faltosos = pd.read_excel(planilha_faltosos, sheet_name='Sheet', header=3)
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

    # Deixar a coluna observação vazia
    df_mesclado['Observação'] = ''

    # Selecionar as colunas finais
    df_final = df_mesclado[['Contrato', 'Aluno', 'Observação', 'Educador', 'Celular']]

    # Adicionar "Funcionário" na coluna "Observação" onde a coluna "Educador" contém "Sangela"
    df_final.loc[df_final['Educador'].str.contains('Sangela', na=False), 'Observação'] = 'Funcionário'


    caminho_saida = Path.home() / "Documents" / "EasyLog" / "Data" / "faltosos_filtrados.xlsx"

    # Salvar o resultado
    df_final.to_excel(caminho_saida, index=False)

    ajustar_largura_colunas(caminho_saida)

    print("Processamento concluído. O arquivo de faltosos foi gerado.")

def filtrar_alunos_atencao(planilha_alunos_atencao):
    df_alunos_atencao = pd.read_excel(planilha_alunos_atencao, sheet_name='Sheet')
    
    # Transferir telefones residenciais para a coluna celular quando não houver
    df_alunos_atencao['Telefone Responsável'] = df_alunos_atencao['Telefone Responsável'].fillna(df_alunos_atencao['Telefone Aluno'])
    
    df_alunos_atencao = df_alunos_atencao[['Nome Aluno', 'Educador', 'Data Cadastro Contrato', 'Telefone Responsável', 'Faltas', 'Reposições']]
    
    # Manter apenas o primeiro nome do educador
    df_alunos_atencao['Educador'] = df_alunos_atencao['Educador'].str.split().str[0]
    
    # Filtrar as linhas onde o valor da coluna é diferente do nome_a_remover
    df_alunos_atencao = df_alunos_atencao[df_alunos_atencao['Educador'] != 'Sangela']
    
    caminho_saida = Path.home() / "Documents" / "EasyLog" / "Data" / "alunos_atencao_filtrados.xlsx"

    # Salvar o resultado
    df_alunos_atencao.to_excel(caminho_saida, index=False)
    
def ajustar_largura_colunas(arquivo):
    # Carregar o arquivo Excel
    wb = openpyxl.load_workbook(arquivo)
    ws = wb.active

    # Definir larguras específicas para as colunas
    colunas_largura = {
        "A": 10.0,
        "B": 40.0,
        "C": 50.0,
        "D": 12.0,
        "E": 18.0
    }

    # Ajustar a largura das colunas conforme o dicionário
    for coluna, largura in colunas_largura.items():
        ws.column_dimensions[coluna].width = largura

    # Salvar o arquivo depois de ajustar as larguras
    wb.save(arquivo)

def preparar_data_faltosos(campo_data_inicial, campo_data_final):
    from bot import gerar_faltosos
    data_inicial = campo_data_inicial.get()
    data_final = campo_data_final.get()

    data_inicial = data_inicial.replace("/", "")
    
    gerar_faltosos(data_inicial, data_final) 
    
def preparar_alunos_atencao(campo_data_inicial, campo_data_final):
    from bot import gerar_alunos_atencao
    data_inicial = campo_data_inicial.get()
    data_final = campo_data_final.get()

    data_inicial = data_inicial.replace("/", "")
    
    gerar_alunos_atencao(data_inicial, data_final) 
