import re
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP
from sqlalchemy.orm import Session
from database.db import BaseReferencia, Composicao

def limpar_numero_br(valor) -> float:
    if pd.isna(valor) or valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip()
    if texto in ("", "-", "N/A", "nan", "None"):
        return 0.0
    texto = re.sub(r"[^\d,.-]", "", texto)
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(Decimal(texto).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP))
    except Exception:
        return 0.0

def limpar_texto(valor) -> str:
    if pd.isna(valor) or valor is None:
        return ""
    return re.sub(r"\s+", " ", str(valor)).strip()

def detectar_linha_cabecalho(df_bruto: pd.DataFrame, palavras_chave: list[str]) -> int:
    for idx in range(min(30, len(df_bruto))):
        linha_str = " ".join([str(v).upper() for v in df_bruto.iloc[idx].values if pd.notna(v)])
        if all(p.upper() in linha_str for p in palavras_chave):
            return idx
    return 0

def padronizar_planilha(
    arquivo,
    tipo_base: str,
    mapa_manual: dict = None,
    linha_cabecalho_manual: int = 0
) -> pd.DataFrame:
    """Lê o arquivo Excel (upload do Streamlit) e retorna DataFrame padronizado."""
    if tipo_base == "GENERICA" and mapa_manual:
        df = pd.read_excel(arquivo, skiprows=linha_cabecalho_manual)
        col_cod = mapa_manual["codigo"]
        col_desc = mapa_manual["descricao"]
        col_unid = mapa_manual["unidade"]
        col_preco = mapa_manual["preco"]
    else:
        df_bruto = pd.read_excel(arquivo, header=None)
        chaves = ["COD", "DESC", "UNID"] if tipo_base != "SINAPI" else ["CODIGO", "DESCRICAO", "UNIDADE"]
        header_row = detectar_linha_cabecalho(df_bruto, chaves)
        arquivo.seek(0)
        df = pd.read_excel(arquivo, skiprows=header_row)
        df.columns = [limpar_texto(c).upper() for c in df.columns]

        col_cod = next((c for c in df.columns if "COD" in c or "CÓD" in c), df.columns[0])
        col_desc = next((c for c in df.columns if "DESC" in c), df.columns[1])
        col_unid = next((c for c in df.columns if "UNID" in c), df.columns[2])
        col_preco = next(
            (c for c in df.columns if any(k in c for k in ["CUSTO TOTAL", "UNITÁRIO", "UNITARIO", "PREÇO", "PRECO", "VALOR"])),
            df.columns[-1]
        )

    df_norm = pd.DataFrame()
    df_norm["codigo"] = df[col_cod].apply(limpar_texto).str.replace(r"\.0$", "", regex=True)
    df_norm["descricao"] = df[col_desc].apply(limpar_texto)
    df_norm["unidade"] = df[col_unid].apply(limpar_texto).str.upper()
    df_norm["preco_unitario"] = df[col_preco].apply(limpar_numero_br)

    df_norm = df_norm[
        (df_norm["codigo"] != "") &
        (df_norm["codigo"].str.lower() != "nan") &
        (df_norm["descricao"] != "") &
        (df_norm["unidade"] != "")
    ].drop_duplicates(subset=["codigo"], keep="first")

    return df_norm

def salvar_base_no_banco(
    db: Session,
    nome: str,
    descricao: str,
    uf: str,
    data_base,
    desonerado: bool,
    df_norm: pd.DataFrame
) -> int:
    base = db.query(BaseReferencia).filter_by(
        nome=nome.upper(), uf=uf.upper(), data_base=data_base, desonerado=desonerado
    ).first()

    if not base:
        base = BaseReferencia(
            nome=nome.upper(), descricao=descricao, uf=uf.upper(),
            data_base=data_base, desonerado=desonerado
        )
        db.add(base)
        db.flush()
    else:
        db.query(Composicao).filter_by(base_id=base.id).delete()

    objetos = [
        Composicao(
            base_id=base.id,
            codigo=row["codigo"],
            descricao=row["descricao"],
            unidade=row["unidade"],
            preco_unitario=row["preco_unitario"]
        )
        for row in df_norm.to_dict(orient="records")
    ]
    db.bulk_save_objects(objetos)
    db.commit()
    return len(objetos)