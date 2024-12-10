import os
import sys

def resource_path(relative_path):
    """Obtém o caminho absoluto do recurso, ajustado para executáveis criados com PyInstaller."""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# Defina os caminhos centralizados aqui
caminhos = {
    "icone": resource_path("imagens/icone.ico"),
    "print_verificacao": resource_path('imagens/print_verificaçao.png'),
    "aba_anexar": resource_path('imagens/aba_anexar.png'),
    "abrir_planilha": resource_path('imagens/abrir_planilha.png'),
    "aluno_encontrado": resource_path('imagens/aluno_encontrado.png'),
    "anexar" : resource_path('imagens/anexar.png'),
    "hub_aberto" : resource_path('imagens/hub_aberto.png'),
    "caixa_mensagem" : resource_path('imagens/caixa_mensagem.png'),
    "campo_nome_planilha" : resource_path('imagens/campo_nome_planilha.png'),
    "campo_pesquisa" : resource_path('imagens/campo_pesquisa.png'),
    "contato_inexistente" : resource_path('imagens/contato_inexistente.png'),
    "contato" : resource_path('imagens/contato.png'),
    "contrato_aberto" : resource_path('imagens/contrato_aberto.png'),
    "descricao" : resource_path('imagens/descricao.png'),
    "exportar" : resource_path('imagens/exportar.png'),
    "faltas_por_periodo" : resource_path('imagens/faltas_por_periodo.png'),
    "fotos_e_videos" : resource_path('imagens/fotos_e_videos.png'),
    "imagem" : resource_path('imagens/imagem.png'),
    "lista_faltosos" : resource_path('imagens/lista_faltosos.png'),
    "logo" : resource_path('imagens/logo.png'),
    "menu_iniciar" : resource_path('imagens/menu_iniciar.png'),
    "nova_conversa" : resource_path('imagens/nova_conversa.png'),
    "ocorrencia" : resource_path('imagens/ocorrencia.png'),
    "opcoes_exportacao" : resource_path('imagens/opcoes_exportacao.png'),
    "pesquisa_aluno" : resource_path('imagens/pesquisa_aluno.png'),
    "pesquisa_contato" : resource_path('imagens/pesquisa_contato.png'),
    "pesquisar" : resource_path('imagens/pesquisar.png'),
    "presencas_e_faltas" : resource_path('imagens/presencas_e_faltas.png'),
    "salvar" : resource_path('imagens/salvar.png'),
    "substituir_arquivo" : resource_path('imagens/substituir_arquivo.png'),
    "visu_presencas_e_faltas" : resource_path('imagens/visu_presencas_e_faltas.png'),
    "visualizar" : resource_path('imagens/visualizar.png'),
    "whatsapp_aberto" : resource_path('imagens/whatsapp_aberto.png'),
    "whatsapp_encontrado":resource_path('imagens/whatsapp_encontrado.png'),
    "bot": resource_path("bot.py"),
    "historicos": resource_path("scripts/historicos.py"),
    "mensagens": resource_path("scripts/mensagens.py"),
    "planilhas": resource_path("scripts/planilhas.py"),
    "ocr": resource_path("scripts/ocr.py"),
    # Adicione outros caminhos necessários
} # type: ignore
