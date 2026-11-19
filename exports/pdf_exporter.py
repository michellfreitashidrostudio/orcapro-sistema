import io
from jinja2 import Template
from xhtml2pdf import pisa

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<style>
  @page {
    size: a4 landscape;
    margin: 1.2cm;
  }
  body { font-family: Helvetica, Arial, sans-serif; font-size: 8.5pt; color: #222; }
  .header { border-bottom: 2px solid #1F4E78; padding-bottom: 6px; margin-bottom: 12px; }
  .header h1 { margin: 0; font-size: 13pt; color: #1F4E78; }
  .header p { margin: 3px 0 0 0; color: #555; font-size: 8.5pt; }
  table { width: 100%; border-collapse: collapse; margin-bottom: 14px; }
  th { background-color: #1F4E78; color: #ffffff; padding: 5px 4px; text-align: left; font-size: 8pt; }
  td { border: 0.5pt solid #cccccc; padding: 4px; vertical-align: top; }
  .num { text-align: right; }
  .center { text-align: center; }
  .etapa-row td { background-color: #D9E1F2; font-weight: bold; color: #1F4E78; }
  .total-row td { background-color: #1F4E78; color: #ffffff; font-weight: bold; font-size: 9pt; }
  .badge-A { color: #d9534f; font-weight: bold; }
  .badge-B { color: #f0ad4e; font-weight: bold; }
  .badge-C { color: #5cb85c; font-weight: bold; }
  .page-break { page-break-before: always; }
</style>
</head>
<body>
  <div class="header">
    <h1>PLANILHA ORÇAMENTÁRIA SINTÉTICA — {{ titulo }}</h1>
    <p><strong>Cliente:</strong> {{ cliente }} | <strong>Local:</strong> {{ local_obra }} | <strong>Regra:</strong> {{ regra }}</p>
  </div>

  <table>
    <thead>
      <tr>
        <th style="width:6%">Item</th>
        <th style="width:10%">Fonte</th>
        <th style="width:8%">Código</th>
        <th style="width:36%">Descrição do Serviço</th>
        <th class="center" style="width:5%">Unid.</th>
        <th class="num" style="width:7%">Quant.</th>
        <th class="num" style="width:9%">Unit. s/ BDI</th>
        <th class="num" style="width:5%">BDI</th>
        <th class="num" style="width:7%">Unit. c/ BDI</th>
        <th class="num" style="width:7%">Total (R$)</th>
      </tr>
    </thead>
    <tbody>
      {% for etapa, dados in etapas.items() %}
      <tr class="etapa-row">
        <td colspan="9">{{ etapa }}</td>
        <td class="num">{{ "{:,.2f}".format(dados.subtotal).replace(",", "X").replace(".", ",").replace("X", ".") }}</td>
      </tr>
      {% for item in dados.itens %}
      <tr>
        <td>{{ item.item_eap }}</td>
        <td>{{ item.fonte_base }}</td>
        <td>{{ item.codigo }}</td>
        <td>{{ item.descricao }}</td>
        <td class="center">{{ item.unidade }}</td>
        <td class="num">{{ "{:,.2f}".format(item.quantidade).replace(",", "X").replace(".", ",").replace("X", ".") }}</td>
        <td class="num">{{ "{:,.2f}".format(item.preco_sem_bdi).replace(",", "X").replace(".", ",").replace("X", ".") }}</td>
        <td class="num">{{ "{:.2f}%".format(item.bdi_aplicado).replace(".", ",") }}</td>
        <td class="num">{{ "{:,.2f}".format(item.preco_com_bdi).replace(",", "X").replace(".", ",").replace("X", ".") }}</td>
        <td class="num">{{ "{:,.2f}".format(item.preco_total).replace(",", "X").replace(".", ",").replace("X", ".") }}</td>
      </tr>
      {% endfor %}
      {% endfor %}
      <tr class="total-row">
        <td colspan="9" class="num">TOTAL GERAL DO ORÇAMENTO (R$)</td>
        <td class="num">{{ "{:,.2f}".format(total_geral).replace(",", "X").replace(".", ",").replace("X", ".") }}</td>
      </tr>
    </tbody>
  </table>

  <div class="page-break"></div>
  <div class="header">
    <h1>CURVA ABC DE SERVIÇOS — {{ titulo }}</h1>
  </div>

  <table>
    <thead>
      <tr>
        <th>Fonte</th>
        <th>Código</th>
        <th>Descrição</th>
        <th class="center">Unid.</th>
        <th class="num">Quant.</th>
        <th class="num">Total (R$)</th>
        <th class="num">Peso (%)</th>
        <th class="num">Acumulado (%)</th>
        <th class="center">Faixa</th>
      </tr>
    </thead>
    <tbody>
      {% for r in abc_rows %}
      <tr>
        <td>{{ r.fonte_base }}</td>
        <td>{{ r.codigo }}</td>
        <td>{{ r.descricao }}</td>
        <td class="center">{{ r.unidade }}</td>
        <td class="num">{{ "{:,.2f}".format(r.quantidade).replace(",", "X").replace(".", ",").replace("X", ".") }}</td>
        <td class="num">{{ "{:,.2f}".format(r.preco_total).replace(",", "X").replace(".", ",").replace("X", ".") }}</td>
        <td class="num">{{ "{:.2f}%".format(r.peso_pct).replace(".", ",") }}</td>
        <td class="num">{{ "{:.2f}%".format(r.acumulado_pct).replace(".", ",") }}</td>
        <td class="center"><span class="badge-{{ r.faixa }}">{{ r.faixa }}</span></td>
      </tr>
      {% endfor %}
    </tbody>
  </table>
</body>
</html>
"""

def gerar_pdf_orcamento(orcamento, df_itens, df_abc) -> bytes:
    etapas = {}
    for etapa, grupo in df_itens.groupby("etapa", sort=False):
        etapas[etapa] = {
            "subtotal": float(grupo["preco_total"].sum()),
            "itens": grupo.to_dict(orient="records")
        }

    html_renderizado = Template(HTML_TEMPLATE).render(
        titulo=orcamento.titulo,
        cliente=orcamento.cliente or "Não informado",
        local_obra=orcamento.local_obra or "Não informado",
        regra="Truncamento na 2ª casa decimal (TCU/Caixa)" if orcamento.truncar_valores else "Arredondamento Padrão",
        etapas=etapas,
        total_geral=float(df_itens["preco_total"].sum()),
        abc_rows=df_abc.to_dict(orient="records")
    )

    buffer = io.BytesIO()
    pisa_status = pisa.CreatePDF(html_renderizado, dest=buffer)
    if pisa_status.err:
        raise RuntimeError("Erro ao gerar o PDF.")
    return buffer.getvalue()