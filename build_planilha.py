#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Desafio DIO — Formação Análise de Dados com Excel e IA
"Ferramenta de Controle de Investimentos com Excel"

Constrói do zero (planilha em branco) a ferramenta que simula investimentos em
FIIs: aporte mensal -> patrimônio (VF) -> dividendos, cenários de 2 a 30 anos,
intervalos nomeados e divisão do aporte por perfil com PROCV + chave composta.

Saída: controle-investimentos.xlsx
"""
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- paleta
NAVY      = "1F4E79"   # azul-marinho (banner, cabeçalhos, títulos)
NAVY_DARK = "17375E"
LIGHT_BLUE= "DDEBF7"   # seções
INPUT_YEL = "FFF2CC"   # células editáveis
CALC_GRAY = "D9D9D9"   # células calculadas (ninguém deve editar)
WHITE     = "FFFFFF"
BORDER_C  = "BFBFBF"

thin   = Side(style="thin", color=BORDER_C)
box    = Border(left=thin, right=thin, top=thin, bottom=thin)
no_bd  = Border()

F_TITLE  = Font(name="Calibri", size=16, bold=True, color=WHITE)
F_SECT   = Font(name="Calibri", size=12, bold=True, color=WHITE)
F_LABEL  = Font(name="Calibri", size=11, color="1F1F1F")
F_VALUE  = Font(name="Calibri", size=11, bold=True, color="1F1F1F")
F_NOTE   = Font(name="Calibri", size=9, italic=True, color="808080")
F_HDR    = Font(name="Calibri", size=11, bold=True, color=WHITE)

FILL_TITLE = PatternFill("solid", fgColor=NAVY)
FILL_SECT  = PatternFill("solid", fgColor=NAVY)
FILL_INPUT = PatternFill("solid", fgColor=INPUT_YEL)
FILL_CALC  = PatternFill("solid", fgColor=CALC_GRAY)
FILL_HDR   = PatternFill("solid", fgColor=NAVY)
FILL_LIGHT = PatternFill("solid", fgColor=LIGHT_BLUE)

CUR = '"R$" #,##0.00'   # moeda
PCT = '0.00%'           # percentual
INT = '0'               # inteiro

PROFILES = ["Conservador", "Moderado", "Agressivo"]
FII_TYPES = ["PAPEL", "TIJOLO", "HÍBRIDOS", "FOFs", "DESENVOLVIMENTO", "HOTELARIAS"]

# percentuais por perfil (exemplo educacional — cada perfil soma exatamente 100%)
ALLOCATION = {
    "Conservador": {"PAPEL": 0.40, "TIJOLO": 0.40, "HÍBRIDOS": 0.10,
                    "FOFs": 0.10, "DESENVOLVIMENTO": 0.00, "HOTELARIAS": 0.00},
    "Moderado":    {"PAPEL": 0.30, "TIJOLO": 0.30, "HÍBRIDOS": 0.10,
                    "FOFs": 0.10, "DESENVOLVIMENTO": 0.10, "HOTELARIAS": 0.10},
    "Agressivo":   {"PAPEL": 0.35, "TIJOLO": 0.10, "HÍBRIDOS": 0.05,
                    "FOFs": 0.05, "DESENVOLVIMENTO": 0.25, "HOTELARIAS": 0.20},
}

# ---------------------------------------------------------------- workbook
wb = Workbook()

# ================================================================ ABA: Simulador
ws = wb.active
ws.title = "Simulador"
ws.sheet_view.showGridLines = False
ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1

widths = {"A": 42, "B": 20, "C": 18, "D": 20, "E": 14, "F": 14}
for col, w in widths.items():
    ws.column_dimensions[col].width = w


def section(row, text, span="A:F"):
    col_start, col_end = span.split(":")
    ws.merge_cells(f"{col_start}{row}:{col_end}{row}")
    c = ws[f"{col_start}{row}"]
    c.value = text
    c.font = F_SECT
    c.fill = FILL_SECT
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 20


def label_val(row, label, formula_or_value, fmt, editable, bold=True, span=None):
    """Monta par rotulo+valor; editable define preenchimento amarelo/cinza."""
    lc = ws[f"A{row}"]
    lc.value = label
    lc.font = F_LABEL
    lc.alignment = Alignment(vertical="center", indent=1)
    vc = ws[f"B{row}"]
    vc.value = formula_or_value
    if fmt is not None:
        vc.number_format = fmt
    vc.font = F_VALUE if bold else F_LABEL
    vc.alignment = Alignment(horizontal="right", vertical="center")
    vc.fill = FILL_INPUT if editable else FILL_CALC
    vc.border = box
    return vc


# -- banner ---------------------------------------------------------------
ws.merge_cells("A1:F1")
banner = ws["A1"]
banner.value = "CONTROLE DE INVESTIMENTOS  ·  FUNDOS IMOBILIÁRIOS (FIIs)"
banner.font = F_TITLE
banner.fill = FILL_TITLE
banner.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 34

ws.merge_cells("A2:F2")
sub = ws["A2"]
sub.value = "Do aporte mensal ao patrimônio e aos dividendos — cenários de 2 a 30 anos, por perfil de investidor"
sub.font = Font(name="Calibri", size=10, italic=True, color="595959")
sub.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[2].height = 16

# -- 1) configurações ------------------------------------------------------
section(3, "CONFIGURAÇÕES")
label_val(4, "Salário (R$)", 2000, CUR, editable=True)                      # salario
label_val(5, "Rendimento da carteira (% a.m.)", 0.006, PCT, editable=True)  # rendimento_carteira
label_val(6, "Sugestão de investimento — 30% do salário (R$)", "=B4*30%", CUR, editable=False)

# -- 2) simulação principal (as 5 perguntas de negócio) --------------------
section(8, "SIMULAÇÃO PRINCIPAL")
label_val(9,  "1) Quanto investir por mês (R$)?", 200, CUR, editable=True)        # aporte
label_val(10, "2) Por quantos anos?", 5, INT, editable=True)                     # qtd_anos
label_val(11, "3) Taxa de rendimento mensal?", 0.01079, PCT, editable=True)      # taxa_mensal
label_val(12, "4) Patrimônio acumulado (R$)", "=FV(taxa_mensal,qtd_anos*12,-aporte)", CUR, editable=False)   # patrimonio
label_val(13, "5) Dividendos por mês (R$)", "=patrimonio*rendimento_carteira", CUR, editable=False)

# -- 3) cenários 2..30 anos (referências absolutas) ------------------------
section(15, "CENÁRIOS — PROJEÇÃO DE 2 A 30 ANOS")
hdr_row = 16
for col, texto in zip("ABCD", ["Anos", "Quanto em N anos?", "Patrimônio (VF)", "Dividendos/mês"]):
    c = ws[f"{col}{hdr_row}"]
    c.value = texto
    c.font = F_HDR
    c.fill = FILL_HDR
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = box

for i, anos in enumerate([2, 5, 10, 20, 30]):
    r = hdr_row + 1 + i
    ws[f"A{r}"] = anos
    ws[f"A{r}"].number_format = INT
    ws[f"A{r}"].alignment = Alignment(horizontal="center")
    ws[f"B{r}"] = f"Quanto em {anos} Anos?"
    ws[f"C{r}"] = f"=FV($B$11,A{r}*12,-$B$9)"            # VF com referências absolutas
    ws[f"C{r}"].number_format = CUR
    ws[f"D{r}"] = f"=C{r}*$B$5"                          # dividendos = patrimônio * rendimento
    ws[f"D{r}"].number_format = CUR
    for col in "ABCD":
        cell = ws[f"{col}{r}"]
        cell.border = box
        cell.fill = FILL_CALC if col in "CD" else PatternFill()
        if col in "BC":
            cell.font = F_LABEL
        if col == "A":
            cell.font = F_LABEL

# -- 4) perfil do investidor + divisão do aporte (PROCV com chave composta) --
section(23, "PERFIL DO INVESTIDOR — DIVISÃO DO APORTE")
label_val(24, "Perfil do investidor", "Moderado", None, editable=True)

dv = DataValidation(type="list", formula1='"Conservador,Moderado,Agressivo"',
                    allow_blank=False, showErrorMessage=True,
                    errorTitle="Perfil inválido",
                    error="Escolha um perfil da lista: Conservador, Moderado ou Agressivo.")
ws.add_data_validation(dv)
dv.add("B24")

label_val(25, "Valor investido por mês (R$)", "=aporte", CUR, editable=False)

tbl_hdr = 27
for col, texto in zip("ACD", ["Tipo de FII", "Percentual sugerido", "Valor por mês (R$)"]):
    c = ws[f"{col}{tbl_hdr}"]
    c.value = texto
    c.font = F_HDR
    c.fill = FILL_HDR
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = box

first_data = tbl_hdr + 1  # 28
for i, tipo in enumerate(FII_TYPES):
    r = first_data + i
    ws[f"A{r}"] = tipo
    ws[f"A{r}"].font = F_LABEL
    ws[f"A{r}"].alignment = Alignment(vertical="center", indent=1)
    ws[f"C{r}"] = f'=VLOOKUP($B$24&"-"&A{r},tabela_perfis,4,FALSE)'   # PROCV chave composta
    ws[f"C{r}"].number_format = PCT
    ws[f"D{r}"] = f"=C{r}*$B$25"
    ws[f"D{r}"].number_format = CUR
    for col in "ACD":
        cell = ws[f"{col}{r}"]
        cell.border = box
        cell.fill = FILL_CALC if col in "CD" else PatternFill()

last_data = first_data + len(FII_TYPES) - 1  # 33
tot_row = last_data + 1                       # 34
ws[f"A{tot_row}"] = "TOTAL"
ws[f"A{tot_row}"].font = Font(name="Calibri", size=11, bold=True)
ws[f"A{tot_row}"].alignment = Alignment(vertical="center", indent=1)
ws[f"C{tot_row}"] = f"=SUM(C{first_data}:C{last_data})"
ws[f"C{tot_row}"].number_format = PCT
ws[f"D{tot_row}"] = f"=SUM(D{first_data}:D{last_data})"
ws[f"D{tot_row}"].number_format = CUR
for col in "ACD":
    cell = ws[f"{col}{tot_row}"]
    cell.border = box
    cell.fill = FILL_LIGHT
    cell.font = Font(name="Calibri", size=11, bold=True)

# -- 5) simulação com aporte sugerido (evolução) ----------------------------
section(36, "SIMULAÇÃO COM O APORTE SUGERIDO (30% DO SALÁRIO)")
label_val(37, "Patrimônio em 10 anos (R$)", "=FV(taxa_mensal,10*12,-sugestao_investimento)", CUR, editable=False)
label_val(38, "Dividendos por mês (R$)", "=B37*rendimento_carteira", CUR, editable=False)

# -- nota ----------------------------------------------------------------
ws.merge_cells("A40:F40")
nota = ws["A40"]
nota.value = "Valores de exemplo · Ferramenta educacional — não é recomendação de investimento."
nota.font = F_NOTE
nota.alignment = Alignment(horizontal="left", indent=1)

ws.print_area = "A1:F40"

# ================================================================ ABA: Perfis
sup = wb.create_sheet("Perfis")
sup.sheet_view.showGridLines = False
sup.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
sup.page_setup.orientation = "landscape"
sup.page_setup.fitToWidth = 1
sup.page_setup.fitToHeight = 1
for col, w in {"A": 26, "B": 18, "C": 20, "D": 12, "E": 4, "F": 22, "G": 12}.items():
    sup.column_dimensions[col].width = w

sup.merge_cells("A1:D1")
t = sup["A1"]
t.value = "TABELA DE APOIO — PERFIL × TIPO DE FII"
t.font = F_TITLE
t.fill = FILL_TITLE
t.alignment = Alignment(horizontal="center", vertical="center")
sup.row_dimensions[1].height = 30

sup.merge_cells("A2:D2")
s2 = sup["A2"]
s2.value = "A chave composta (PERFIL&TIPO) é o que o PROCV procura na aba Simulador"
s2.font = Font(name="Calibri", size=10, italic=True, color="595959")
s2.alignment = Alignment(horizontal="center", vertical="center")

hdr3 = 3
for col, texto in zip("ABCD", ["CHAVE", "PERFIL", "TIPO DE FII", "%"]):
    c = sup[f"{col}{hdr3}"]
    c.value = texto
    c.font = F_HDR
    c.fill = FILL_HDR
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = box

r = hdr3 + 1  # 4
for perfil in PROFILES:
    for tipo in FII_TYPES:
        sup[f"A{r}"] = f'=B{r}&"-"&C{r}'
        sup[f"B{r}"] = perfil
        sup[f"C{r}"] = tipo
        sup[f"D{r}"] = ALLOCATION[perfil][tipo]
        sup[f"D{r}"].number_format = '0%'
        for col in "ABCD":
            cell = sup[f"{col}{r}"]
            cell.border = box
            cell.font = F_LABEL if col != "A" else Font(name="Calibri", size=10, color="808080")
        r += 1

lim = r - 1  # 21 (última linha de dados)

# -- conferência: soma por perfil (evidência de que cada perfil soma 100%) -- 
cr = lim + 2  # 23
sup.merge_cells(f"A{cr}:D{cr}")
chk = sup[f"A{cr}"]
chk.value = "CONFERÊNCIA — soma por perfil (cada perfil deve somar 100%)"
chk.font = F_SECT
chk.fill = FILL_SECT
chk.alignment = Alignment(horizontal="left", vertical="center", indent=1)

hr = cr + 1  # 24
for col, texto in zip("BD", ["PERFIL", "SOMA %"]):
    c = sup[f"{col}{hr}"]
    c.value = texto
    c.font = F_HDR
    c.fill = FILL_HDR
    c.alignment = Alignment(horizontal="center")
    c.border = box

for i, perfil in enumerate(PROFILES):
    rr = hr + 1 + i
    sup[f"B{rr}"] = perfil
    sup[f"D{rr}"] = f'=SUMIF($B${hdr3+1}:$B${lim},"{perfil}",$D${hdr3+1}:$D${lim})'
    sup[f"D{rr}"].number_format = '0.00%'
    for col in "BD":
        cell = sup[f"{col}{rr}"]
        cell.border = box
        cell.font = F_LABEL if col == "B" else F_VALUE

sup.print_area = f"A1:G{hr + 4}"

# ================================================================ intervalos nomeados
named = {
    "salario":              "Simulador!$B$4",
    "rendimento_carteira":  "Simulador!$B$5",
    "sugestao_investimento": "Simulador!$B$6",
    "aporte":               "Simulador!$B$9",
    "qtd_anos":             "Simulador!$B$10",
    "taxa_mensal":          "Simulador!$B$11",
    "patrimonio":           "Simulador!$B$12",
    "tabela_perfis":        "Perfis!$A$4:$D$21",
}
for name, ref in named.items():
    wb.defined_names.add(DefinedName(name, attr_text=ref))

# sanity: cada perfil soma 100%
for p in PROFILES:
    soma = sum(ALLOCATION[p].values())
    assert abs(soma - 1.0) < 1e-9, f"{p} soma {soma}"
print("percentuais OK — Conservador/Moderado/Agressivo somam 100%")

OUT = "/mnt/c/Users/conta/DIO/controle-investimentos-excel/controle-investimentos.xlsx"
wb.save(OUT)
print("salvo:", OUT)