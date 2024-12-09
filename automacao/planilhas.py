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
                "nome": row["Nome"],         # Nome do aluno
            })
        
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
                "Aluno": row["Aluno"],              # Nome do aluno
                "Observação": row["Observação"],    # Justificativa da falta
                "Educador": row["Educador"],        # Nome do Educador
                "Contato": str(row["Celular"])     # Contato do aluno
            })
        
        for aluno in faltosos:
            aluno['Contato'] = formatar_telefone(aluno['Contato'])
            aluno['Educador'] = manter_primeiro_nome(aluno['Educador'])

        return faltosos
    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")
        return []
    
# Função para formatar o número de telefone
def formatar_telefone(celular):
    # Usando expressão regular para garantir que o número tenha o formato correto
    if celular:
        celular = re.sub(r'\D', '', celular)  # Remove qualquer caractere não numérico
        if len(celular) == 11:  # Verifica se o número tem 11 dígitos
            return f"({celular[:2]}){celular[2]}{celular[3:7]}-{celular[7:]}"
    return celular  # Retorna o número original caso não tenha o formato esperado

# Função para manter apenas o primeiro nome do educador
def manter_primeiro_nome(nome):
    if isinstance(nome, str):  # Verifica se é uma string
        return nome.split()[0]  # Divide o nome completo e retorna o primeiro nome
    return nome  # Retorna o valor original se não for string

def filtrar_faltosos(caminho_arquivo):
    df_faltosos = pd.read_excel(caminho_arquivo, sheet_name='Sheet', header=3)

    # Renomear colunas para consistência
    df_faltosos = df_faltosos.rename(
        columns={
            "Unnamed: 1": "Aluno",
            "Unnamed: 7": "Observação",
            "Unnamed: 8": "Tel Residencial",
            "Unnamed: 9": "Celular",
            "Unnamed: 0": "Contrato"
        }
    )

    # Selecionar colunas relevantes
    df_faltosos = df_faltosos[["Contrato", "Aluno", "Observação", "Tel Residencial", "Celular"]]

    relacionar_educador(df_faltosos)

def relacionar_educador(df_faltosos):
    # Caminho para os arquivos
    caminho_educadores = Path.home() / "Documents" / "EasyLog" / "Planilhas" / "alunos_e_educadores.xls"

    # Carregar planilhas
    df_educadores = pd.read_excel(caminho_educadores)

    # Renomear colunas para consistência
    df_educadores = df_educadores.rename(
        columns={
            "Contrato": "Contrato",
            "Nome Aluno": "Aluno",
            "Educador": "Educador",
            "Telefone Responsável": "Celular"
        }
    )

    # Selecionar colunas relevantes
    df_educadores = df_educadores[["Contrato", "Aluno", "Educador", "Celular"]]

    # Mesclar dataframes com base no contrato
    df_mesclado = pd.merge(df_faltosos, df_educadores, on="Aluno", how="left")

    # Preencher "Celular" vazio com "Tel Residencial"
    df_mesclado['Celular'] = df_mesclado['Celular_x'].combine_first(df_mesclado['Celular_y'])

    # Remover colunas desnecessárias
    df_mesclado = df_mesclado.drop(columns=["Tel Residencial", "Celular_x", "Celular_y"])

    # Limpar a coluna Observação e remover duplicados
    df_mesclado['Observação'] = None
    df_mesclado = df_mesclado.drop_duplicates()

    print(df_mesclado.columns)

    # Reorganizar colunas
    df_final = df_mesclado[["Contrato", "Aluno", "Observação", "Educador", "Celular"]]

    # Salvar o resultado
    df_final.to_excel("faltosos_filtrados.xlsx", index=False)

    print("Processamento concluído. O arquivo 'faltosos_filtrados.xlsx' foi gerado.")


def ajustar_largura_colunas(arquivo):
    # Carregar o arquivo Excel
    wb = openpyxl.load_workbook(arquivo)
    ws = wb.active

    # Definir larguras específicas para as colunas
    colunas_largura = {
        "B": 12.0,
        "C": 42.0,
        "D": 50.0,
        "E": 16.0
    }

    # Ajustar a largura das colunas conforme o dicionário
    for coluna, largura in colunas_largura.items():
        ws.column_dimensions[coluna].width = largura

    # Salvar o arquivo depois de ajustar as larguras
    wb.save(arquivo)


filtrar_faltosos(Path.home() / "Documents" / "EasyLog" / "Planilhas" / "faltosos.xls")