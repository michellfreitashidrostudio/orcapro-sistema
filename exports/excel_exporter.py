import io
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

AZUL = PatternFill(start_color="0066B3", end_color="0066B3", fill_type="solid")
CINZA = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
F_HEADER = Font(color="FFFFFF", bold=True, size=10)
F_BOLD = Font(bold=True, size=10)
BORDA = Border(
    left=Side(style="thin", color="CCCCCC"), right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"), bottom=Side(style="thin", color="CCCCCC")
)

def gerar_excel_orcamento(orcamento, df_itens: pd.DataFrame, df_abc: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Planilha Sintética"

    ws["A1"] = f"ORÇAMENTO: {orcamento.titulo}"
    ws["A1"].font = Font(bold=True, size=14, color="0066B3")
    ws["A2"] = f"Cliente: {orcamento.cliente} | Local: {orcamento.local_obra}"

    headers = [
        "Item", "Fonte", "Código", "Descrição do Serviço", "Unid.",
        "Quant.", "Unit. s/ BDI (R$)", "BDI (%)", "Unit. c/ BDI (R$)", "Total (R$)"
    ]
    ws.append([])
    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=4, column=col_idx)
        cell.fill = AZUL
        cell.font = F_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center")

    linha_atual = 5
    linhas_totais_etapas = []

    for etapa, grupo in df_itens.groupby("etapa", sort=False):
        ws.cell(row=linha_atual, column=1, value=etapa).font = F_BOLD
        ws.merge_cells(start_row=linha_atual, start_column=1, end_row=linha_atual, end_column=9)
        for c in range(1, 11):
            ws.cell(row=linha_atual, column=c).fill = CINZA
        linha_etapa = linha_atual
        linha_atual += 1
        inicio_itens = linha_atual

        for _, row in grupo.iterrows():
            ws.cell(row=linha_atual, column=1, value=row["item_eap"])
            ws.cell(row=linha_atual, column=2, value=row["fonte_base"])
            ws.cell(row=linha_atual, column=3, value=row["codigo"])
            ws.cell(row=linha_atual, column=4, value=row["descricao"])
            ws.cell(row=linha_atual, column=5, value=row["unidade"])
            ws.cell(row=linha_atual, column=6, value=float(row["quantidade"]))
            ws.cell(row=linha_atual, column=7, value=float(row["preco_sem_bdi"]))
            ws.cell(row=linha_atual, column=8, value=float(row["bdi_aplicado"]) / 100.0)

            if orcamento.truncar_valores:
                ws.cell(row=linha_atual, column=9, value=f"=TRUNC(G{linha_atual}*(1+H{linha_atual}),2)")
                ws.cell(row=linha_atual, column=10, value=f"=TRUNC(F{linha_atual}*I{linha_atual},2)")
            else:
                ws.cell(row=linha_atual, column=9, value=f"=ROUND(G{linha_atual}*(1+H{linha_atual}),2)")
                ws.cell(row=linha_atual, column=10, value=f"=ROUND(F{linha_atual}*I{linha_atual},2)")

            ws.cell(row=linha_atual, column=6).number_format = "#,##0.00"
            ws.cell(row=linha_atual, column=7).number_format = "#,##0.00"
            ws.cell(row=linha_atual, column=8).number_format = "0.00%"
            ws.cell(row=linha_atual, column=9).number_format = "#,##0.00"
            ws.cell(row=linha_atual, column=10).number_format = "#,##0.00"

            for c in range(1, 11):
                ws.cell(row=linha_atual, column=c).border = BORDA
            linha_atual += 1

        fim_itens = linha_atual - 1
        cell_sub = ws.cell(row=linha_etapa, column=10, value=f"=SUM(J{inicio_itens}:J{fim_itens})")
        cell_sub.font = F_BOLD
        cell_sub.number_format = "R$ #,##0.00"
        linhas_totais_etapas.append(f"J{linha_etapa}")

    ws.cell(row=linha_atual, column=9, value="TOTAL GERAL").font = F_BOLD
    formula_total = "=" + "+".join(linhas_totais_etapas) if linhas_totais_etapas else 0
    cell_total = ws.cell(row=linha_atual, column=10, value=formula_total)
    cell_total.font = Font(bold=True, size=11, color="0066B3")
    cell_total.number_format = "R$ #,##0.00"

    larguras = {1: 10, 2: 14, 3: 12, 4: 55, 5: 8, 6: 12, 7: 16, 8: 10, 9: 16, 10: 18}
    for col, width in larguras.items():
        ws.column_dimensions[ws.cell(row=4, column=col).column_letter].width = width

    ws_abc = wb.create_sheet(title="Curva ABC")
    ws_abc.append(["Fonte", "Código", "Descrição", "Unid.", "Quant.", "Total (R$)", "Peso (%)", "Acumulado (%)", "Faixa"])
    for c in range(1, 10):
        ws_abc.cell(row=1, column=c).fill = AZUL
        ws_abc.cell(row=1, column=c).font = F_HEADER

    for _, r in df_abc.iterrows():
        ws_abc.append([
            r["fonte_base"], r["codigo"], r["descricao"], r["unidade"],
            float(r["quantidade"]), float(r["preco_total"]),
            round(float(r["peso_pct"]), 2), round(float(r["acumulado_pct"]), 2), r["faixa"]
        ])

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()

def gerar_excel_bdi(perfil_bdi) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Demonstrativo de BDI"
    ws["A1"] = f"DEMONSTRATIVO DE CÁLCULO DO BDI — {perfil_bdi.nome}"
    ws["A1"].font = Font(bold=True, size=13, color="0066B3")
    ws["A2"] = "Fórmula Acórdão nº 2622/2013 - TCU: BDI = [((1+AC+S+R+G)*(1+DF)*(1+L))/(1-T)] - 1"

    ws.append([])
    ws.append(["Componente", "Sigla", "Percentual (%)"])
    for c in range(1, 4):
        ws.cell(row=4, column=c).fill = AZUL
        ws.cell(row=4, column=c).font = F_HEADER

    linhas = [
        ("Administração Central", "AC", perfil_bdi.ac),
        ("Seguros e Garantias", "S+G", perfil_bdi.sg),
        ("Riscos", "R", perfil_bdi.r),
        ("Despesas Financeiras", "DF", perfil_bdi.df),
        ("Lucro", "L", perfil_bdi.l),
        ("Tributos (PIS, COFINS, ISS, CPRB)", "T", perfil_bdi.t),
    ]
    for idx, (nome, sigla, val) in enumerate(linhas, start=5):
        ws.cell(row=idx, column=1, value=nome)
        ws.cell(row=idx, column=2, value=sigla)
        cell_v = ws.cell(row=idx, column=3, value=float(val))
        cell_v.number_format = "0.00%"
        for c in range(1, 4):
            ws.cell(row=idx, column=c).border = BORDA

    ws.cell(row=12, column=1, value="BDI RESULTANTE").font = F_BOLD
    ws.cell(row=12, column=2, value="BDI").font = F_BOLD
    c_res = ws.cell(row=12, column=3, value="=((1+C5+C6+C7)*(1+C8)*(1+C9)/(1-C10))-1")
    c_res.font = Font(bold=True, size=11, color="0066B3")
    c_res.number_format = "0.00%"

    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 18

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def gerar_excel_composicoes_proprias(comps) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Composições Próprias"
    ws.append(["Código", "Descrição do Serviço", "Unidade", "Preço Unitário s/ BDI (R$)"])
    for c in range(1, 5):
        ws.cell(row=1, column=c).fill = AZUL
        ws.cell(row=1, column=c).font = F_HEADER

    for idx, comp in enumerate(comps, start=2):
        ws.cell(row=idx, column=1, value=comp.codigo)
        ws.cell(row=idx, column=2, value=comp.descricao)
        ws.cell(row=idx, column=3, value=comp.unidade)
        c_p = ws.cell(row=idx, column=4, value=float(comp.preco_unitario))
        c_p.number_format = "R$ #,##0.00"
        for c in range(1, 5):
            ws.cell(row=idx, column=c).border = BORDA

    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 65
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 24

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def gerar_excel_cronograma(orcamento, df_crono_fin: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Cronograma Físico-Financeiro"
    ws["A1"] = f"CRONOGRAMA FÍSICO-FINANCEIRO — {orcamento.titulo}"
    ws["A1"].font = Font(bold=True, size=13, color="0066B3")
    ws.append([])

    cols = list(df_crono_fin.columns)
    ws.append(cols)
    for c in range(1, len(cols) + 1):
        ws.cell(row=3, column=c).fill = AZUL
        ws.cell(row=3, column=c).font = F_HEADER

    for r_idx, row in enumerate(df_crono_fin.itertuples(index=False), start=4):
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = BORDA
            if c_idx >= 2 and isinstance(val, (int, float)):
                cell.number_format = "R$ #,##0.00"

    ws.column_dimensions["A"].width = 35
    for c in range(2, len(cols) + 1):
        ws.column_dimensions[ws.cell(row=3, column=c).column_letter].width = 16

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()