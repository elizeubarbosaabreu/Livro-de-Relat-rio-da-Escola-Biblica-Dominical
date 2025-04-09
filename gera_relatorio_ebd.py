import os
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from math import isclose

BASE_DIR = "Relatorios_EBD"
CLASSES = [
    "Diretoria", "Abraão", "Samuel", "Ana", "Débora", "Venc. em Cristo",
    "Miriã", "Lírio dos Vales", "Gideões", "Cam. p/ o Céu",
    "Rosa de Saron", "Sold. de Cristo", "Querubins"
]
COLUNAS_RELEVANTES = {
    "Matriculados": 2, "Ausentes": 3, "Presentes": 4,
    "Visitantes": 5, "Total": 6, "Bíblias": 7,
    "Revistas": 8, "Ofertas": 9, "% de Presença": 10
}
TRIMESTRES = {
    "1º Trimestre": ["janeiro", "fevereiro", "março"],
    "2º Trimestre": ["abril", "maio", "junho"],
    "3º Trimestre": ["julho", "agosto", "setembro"],
    "4º Trimestre": ["outubro", "novembro", "dezembro"]
}

# Estilos
bold_center = Font(bold=True)
center_align = Alignment(horizontal="center")
rotated_align = Alignment(horizontal="center", vertical="center", textRotation=90)
currency_format = 'R$ #,##0.00'
percent_format = '0.00%'

# Utilitários
def inicializa_linha_totais():
    return {col: 0 for col in COLUNAS_RELEVANTES}

def somar_linha(origem, destino):
    for k in COLUNAS_RELEVANTES:
        destino[k] += origem[k]

def obter_dados_por_classe(arquivo):
    wb = openpyxl.load_workbook(arquivo, data_only=True)
    ws = wb.active
    dados = {classe: inicializa_linha_totais() for classe in CLASSES}
    for row in ws.iter_rows(min_row=3, max_row=3+len(CLASSES)-1):
        nome_classe = row[0].value
        if nome_classe not in dados:
            continue
        for k, idx in COLUNAS_RELEVANTES.items():
            celula = row[idx-1].value
            valor = float(celula) if celula not in (None, '') else 0
            dados[nome_classe][k] += valor
    return dados

def dividir_e_arredondar(valor, divisor):
    if divisor == 0:
        return 0
    if isinstance(valor, float) and isclose(valor % 1, 0.5, abs_tol=1e-6):
        return int(round(valor))  # evitar erro de arredondamento do Python
    return int(round(valor / divisor))
def criar_relatorio(nome, dados, caminho_destino, contagem_domingos):
    wb = Workbook()
    ws = wb.active
    ws.title = "Relatório"

    # Título principal da planilha
    titulo = f"Relatório do {nome.replace('.xlsx', '')} de {os.path.basename(caminho_destino)} da EBD"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLUNAS_RELEVANTES)+1)
    celula_titulo = ws.cell(row=1, column=1, value=titulo)
    celula_titulo.font = Font(bold=True, size=14)
    celula_titulo.alignment = Alignment(horizontal="center")

    # Cabeçalho na linha 2
    cabecalho = ["Classe"] + list(COLUNAS_RELEVANTES.keys())
    for col, texto in enumerate(cabecalho, 1):
        cel = ws.cell(row=2, column=col, value=texto)
        cel.font = bold_center
        if texto == "Classe":
            cel.alignment = center_align
        else:
            cel.alignment = rotated_align

    # Preencher dados por classe a partir da linha 3
    for i, classe in enumerate(CLASSES, start=3):
        ws.cell(row=i, column=1, value=classe)
        for j, k in enumerate(COLUNAS_RELEVANTES.keys(), start=2):
            valor = dados[classe][k]
            cel = ws.cell(row=i, column=j)
            if k == "Ofertas":
                cel.value = valor
                cel.number_format = currency_format
            elif "%" in k:
                media = valor / contagem_domingos if contagem_domingos > 0 else 0
                cel.value = media
                cel.number_format = percent_format
            else:
                cel.value = dividir_e_arredondar(valor, contagem_domingos)
            cel.alignment = center_align

    # Linha de total geral (última linha + 1)
    linha_total = len(CLASSES) + 3
    ws.cell(row=linha_total, column=1, value="Total Geral").font = bold_center
    soma_matriculados = 0
    soma_presentes = 0

    for j, k in enumerate(COLUNAS_RELEVANTES.keys(), start=2):
        total = sum(dados[classe][k] for classe in CLASSES)
        cel = ws.cell(row=linha_total, column=j)
        if k == "Ofertas":
            cel.value = total
            cel.number_format = currency_format
        elif k == "% de Presença":
            soma_matriculados = sum(dados[classe]["Matriculados"] for classe in CLASSES)
            soma_presentes = sum(dados[classe]["Presentes"] for classe in CLASSES)
            porcentagem = soma_presentes / soma_matriculados if soma_matriculados > 0 else 0
            cel.value = porcentagem
            cel.number_format = percent_format
        else:
            cel.value = dividir_e_arredondar(total, contagem_domingos)
        cel.font = bold_center
        cel.alignment = center_align

    # Formatação condicional
    faixa = f"J3:J{linha_total - 1}"
    ws.conditional_formatting.add(
        faixa,
        CellIsRule(operator='greaterThanOrEqual', formula=['0.8'], fill=PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid'))
    )
    ws.conditional_formatting.add(
        faixa,
        CellIsRule(operator='lessThan', formula=['0.6'], fill=PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid'))
    )

    # Ajuste automático da largura das colunas
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ajuste = min(max_length + 2, 25)
        ws.column_dimensions[col_letter].width = ajuste

    wb.save(os.path.join(caminho_destino, nome))


def gerar_relatorios_anuais_e_trimestrais(ano_escolhido):
    dados_ano = {classe: inicializa_linha_totais() for classe in CLASSES}
    planilhas_ano = 0
    ano_dir = os.path.join(BASE_DIR, str(ano_escolhido))

    for trimestre, meses in TRIMESTRES.items():
        dados_trimestre = {classe: inicializa_linha_totais() for classe in CLASSES}
        contagem_trimestre = 0

        for mes in meses:
            mes_dir = os.path.join(ano_dir, mes)
            if not os.path.exists(mes_dir):
                continue
            for arquivo in os.listdir(mes_dir):
                if arquivo.endswith(".xlsx"):
                    caminho = os.path.join(mes_dir, arquivo)
                    dados = obter_dados_por_classe(caminho)
                    for classe in CLASSES:
                        somar_linha(dados[classe], dados_trimestre[classe])
                        somar_linha(dados[classe], dados_ano[classe])
                    contagem_trimestre += 1
                    planilhas_ano += 1

        criar_relatorio(f"{trimestre}.xlsx", dados_trimestre, ano_dir, contagem_trimestre)

    criar_relatorio("Relatório_Anual.xlsx", dados_ano, ano_dir, planilhas_ano)

# Execução
try:
    ano_input = int(input("Digite o ano para o qual deseja gerar os relatórios (ex: 2024): "))
    gerar_relatorios_anuais_e_trimestrais(ano_input)
    print(f"\n✅ Relatórios gerados com sucesso para o ano {ano_input}!")
except ValueError:
    print("❌ Ano inválido. Digite um número inteiro.")
except Exception as e:
    print(f"❌ Erro inesperado: {e}")
