import math
import pandas as pd

def calcular_bdi_tcu(ac: float, sg: float, r: float, df: float, l: float, t: float) -> float:
    """
    Fórmula Acórdão 2622/2013 - TCU:
    BDI = [((1 + AC + S + R + G) * (1 + DF) * (1 + L)) / (1 - T)] - 1
    Entradas em decimal (ex: 0.04 para 4%). Retorno em percentual (ex: 25.12).
    """
    if t >= 1.0:
        raise ValueError("A alíquota de tributos não pode ser igual ou superior a 100%.")
    numerador = (1 + ac + sg + r) * (1 + df) * (1 + l)
    bdi_decimal = (numerador / (1 - t)) - 1
    return round(bdi_decimal * 100, 2)

def truncar_2_casas(valor: float) -> float:
    """Trunca o valor na 2ª casa decimal sem arredondar para cima (Padrão Caixa/SINAPI)."""
    return math.floor(round(valor * 10000, 4) / 100) / 100.0

def calcular_valores_item(
    preco_sem_bdi: float,
    bdi_percentual: float,
    quantidade: float,
    truncar: bool = True
) -> tuple[float, float]:
    """Retorna (preco_unit_com_bdi, preco_total)."""
    fator = 1 + (bdi_percentual / 100.0)
    bruto_com_bdi = preco_sem_bdi * fator
    preco_com_bdi = truncar_2_casas(bruto_com_bdi) if truncar else round(bruto_com_bdi, 2)
    preco_total = truncar_2_casas(preco_com_bdi * quantidade) if truncar else round(preco_com_bdi * quantidade, 2)
    return preco_com_bdi, preco_total

def gerar_curva_abc(df_itens: pd.DataFrame) -> pd.DataFrame:
    """Agrupa itens por código/descrição e classifica nas faixas A (<=80%), B (<=95%) e C (>95%)."""
    if df_itens.empty:
        return pd.DataFrame()

    df_abc = (
        df_itens.groupby(["fonte_base", "codigo", "descricao", "unidade"], as_index=False)
        .agg(quantidade=("quantidade", "sum"), preco_total=("preco_total", "sum"))
        .sort_values(by="preco_total", ascending=False)
        .reset_index(drop=True)
    )
    total_geral = df_abc["preco_total"].sum()
    if total_geral == 0:
        df_abc["peso_pct"] = 0.0
        df_abc["acumulado_pct"] = 0.0
        df_abc["faixa"] = "C"
        return df_abc

    df_abc["peso_pct"] = (df_abc["preco_total"] / total_geral) * 100
    df_abc["acumulado_pct"] = df_abc["peso_pct"].cumsum()

    def classificar(acum):
        if acum <= 80.0:
            return "A"
        elif acum <= 95.0:
            return "B"
        return "C"

    df_abc["faixa"] = df_abc["acumulado_pct"].apply(classificar)
    return df_abc