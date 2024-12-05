import pandas as pd
import os
from pathlib import Path
import openpyxl
from automação import atualizar_faltosos_e_educadores

def filtrar_faltosos(df):
    # 1. Remover linhas onde a coluna 'Curso' tenha "Annual Book - Multimídia"
    df = df[df['Curso'] != 'Annual Book - Multimídia']

    # 2. Manter apenas as colunas "aluno", "tel residencial" e "celular"
    df = df[['Aluno', 'Tel Residencial', 'Celular']]

    # 3. Remover colunas "B", "E" e "F"
    cols_to_remove = ['B', 'E', 'F']
    df = df.drop(columns=cols_to_remove, errors='ignore')

    # 4. Preencher células vazias na coluna "Celular" (C) com valores da coluna "Tel Residencial" (B)
    df['Celular'] = df['Celular'].fillna(df['Tel Residencial'])

    # 5. Remove a coluna "B"
    df = df.drop('Tel Residencial', axis=1, errors='ignore')

    # 6. Remover linhas duplicadas
    df = df.drop_duplicates()

    df = relacionar_educador(df)

    # Salvar o arquivo final
    output_file_path = './planilhas/faltosos_filtrados.xlsx'
    df.to_excel(output_file_path)
    ajustar_largura_colunas(output_file_path)

    print(f"Arquivo salvo como {output_file_path}")

def relacionar_educador(df_faltosos):
    #atualizar_faltosos_e_educadores()

    caminho_arquivo = Path.home() / "Documents" / "EasyLog" / "Planilhas" / "faltosos_e_educadores.xls"
    
    # Carregar as planilhas
    df_faltosos_e_educadores = pd.read_excel(caminho_arquivo)
    df_faltosos_e_educadores.rename(columns={"Nome Aluno": "Aluno"}, inplace=True)

    # Garantir que as colunas necessárias estão disponíveis
    colunas_faltosos = ['Aluno', 'Tel Residencial', 'Celular']  # Colunas essenciais da planilha 'faltosos'
    colunas_faltosos_e_educadores = ['Aluno', 'Educador']

    if not all(col in df_faltosos.columns for col in colunas_faltosos):
        raise ValueError("A planilha 'faltosos' não contém as colunas necessárias: 'Aluno', 'Tel Residencial' e 'Celular'.")
    if not all(col in df_faltosos_e_educadores.columns for col in colunas_faltosos_e_educadores):
        raise ValueError("A planilha 'faltosos_e_educadores' não contém todas as colunas necessárias.")

    # Selecionar as colunas necessárias de ambas as planilhas
    df_faltosos = df_faltosos[colunas_faltosos]
    df_faltosos_e_educadores = df_faltosos_e_educadores[colunas_faltosos_e_educadores]

    # Realizar a mesclagem dos dados
    df_resultado = pd.merge(
        df_faltosos,
        df_faltosos_e_educadores,
        on='Aluno',  # Coluna em comum para combinar os dados
        how='inner'  # Apenas alunos presentes nas duas planilhas
    )

    caminho_saida = Path.home() / "Documents" / "EasyLog" / "Planilhas" / "faltosos.xls"

    # Salvar o resultado ou exibi-lo
    if caminho_saida:
        df_resultado.to_excel(caminho_saida, index=False)
        print(f"Dados dos faltosos coletados e salvos em: {caminho_saida}")
    else:
        print("Dados dos faltosos coletados:")
        print(df_resultado)
    
    return df_resultado


def ajustar_largura_colunas(arquivo):
    # Carregar o arquivo Excel
    wb = openpyxl.load_workbook(arquivo)
    ws = wb.active

    # Definir larguras específicas para as colunas
    colunas_largura = {
        "A": 30.0,
        "B": 40.0,
        "C": 32.0,
        "D": 40.0,
    }

    # Ajustar a largura das colunas conforme o dicionário
    for coluna, largura in colunas_largura.items():
        ws.column_dimensions[coluna].width = largura

# Carregar a planilha
caminho_arquivo =  Path.home() / "Documents" / "EasyLog" / "Planilhas" / "faltosos.xls" # Substitua pelo nome correto do arquivo
nome_planilha = 'Sheet'  # Substitua pelo nome correto da aba, caso necessário
df = pd.read_excel(caminho_arquivo, sheet_name=nome_planilha, header=3)
print(df)

filtrar_faltosos(df)

