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
    df = pd.read_excel(caminho_arquivo, sheet_name='Sheet', header=3)
    print(df.columns)

    # 1. Remover linhas onde a coluna 'Curso' tenha "Annual Book - Multimídia"
    df = df[df['Curso'] != 'Annual Book - Multimídia']

    # 4. Preencher células vazias na coluna "Celular" (C) com valores da coluna "Tel Residencial" (B)
    df['Celular'] = df['Celular'].fillna(df['Tel Residencial'])

    # 5. Remove a coluna "B"
    df = df.drop('Tel Residencial', axis=1, errors='ignore')

    # 6. Remover linhas duplicadas
    df = df.drop_duplicates()

    df = relacionar_educador(df)

    colunas_necessarias = {'Aluno', 'Celular'}
    colunas_disponiveis = set(df.columns)

    if not colunas_necessarias.issubset(colunas_disponiveis):
        colunas_faltantes = colunas_necessarias - colunas_disponiveis
        raise ValueError(
            f"A planilha 'faltosos' não contém as colunas necessárias: {', '.join(colunas_faltantes)}"
    )

    # Salvar o arquivo final
    caminho_saida = Path.home() / "Documents" / "EasyLog" / "Planilhas" / "faltosos_filtrados.xlsx"
    df.to_excel(caminho_saida)
    ajustar_largura_colunas(caminho_saida)

    print(f"Arquivo salvo como {caminho_saida}")

def relacionar_educador(df_faltosos):
    caminho_arquivo = Path.home() / "Documents" / "EasyLog" / "Planilhas" / "alunos_e_educadores.xls"
    
    # Carregar as planilhas
    df_alunos_e_educadores = pd.read_excel(caminho_arquivo)
    df_alunos_e_educadores.rename(columns={"Nome Aluno": "Aluno"}, inplace=True)

    colunas_necessarias = {'Aluno', 'Educador'}
    colunas_disponiveis = set(df_alunos_e_educadores.columns)

    if not colunas_necessarias.issubset(colunas_disponiveis):
        colunas_faltantes = colunas_necessarias - colunas_disponiveis
        raise ValueError(
            f"A planilha 'alunos_e_educadores' não contém as colunas necessárias: {', '.join(colunas_faltantes)}"
        )

    # Realizar a mesclagem dos dados
    df_resultado = pd.merge(
        df_faltosos,
        df_alunos_e_educadores,
        on='Aluno',  # Coluna em comum para combinar os dados
        how='inner'  # Apenas alunos presentes nas duas planilhas
    )

    # Filtrar para manter apenas as colunas desejadas
    colunas_desejadas = ["Código", "Aluno", "Educador", "Celular"]
    df_resultado = df_resultado[[coluna for coluna in colunas_desejadas if coluna in df_resultado.columns]]

    
    # Adicionar a coluna "Observação" vazia ou com valor padrão
    df_resultado['Observação'] = None  # ou use `None` para deixar vazio

    # Reorganizar as colunas conforme solicitado: "Aluno", "Observação", "Educador", "Celular"
    df_resultado = df_resultado[['Aluno', 'Observação', 'Educador', 'Celular']]

    # Adicionar "Funcionário" na coluna "Observação" onde o valor da coluna "Educador" for "Sangela"
    df_resultado.loc[df_resultado['Educador'] == 'Sangela Amaro Gomes de Souza', 'Observação'] = 'Funcionário'

    # Salvar o resultado ou exibi-lo
    caminho_saida = Path.home() / "Documents" / "EasyLog" / "Planilhas" / "faltosos_filtrados.xlsx"
    df_resultado.to_excel(caminho_saida, index=False)

    print(f"Dados dos faltosos coletados e salvos em: {caminho_saida}")
    return df_resultado


def ajustar_largura_colunas(arquivo):
    # Carregar o arquivo Excel
    wb = openpyxl.load_workbook(arquivo)
    ws = wb.active

    # Definir larguras específicas para as colunas
    colunas_largura = {
        "B": 35.0,
        "C": 50.0,
        "D": 32.0,
        "E": 16.0,
    }

    # Ajustar a largura das colunas conforme o dicionário
    for coluna, largura in colunas_largura.items():
        ws.column_dimensions[coluna].width = largura

    # Salvar o arquivo depois de ajustar as larguras
    wb.save(arquivo)


