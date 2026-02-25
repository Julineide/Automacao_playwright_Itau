import pandas as pd
import re
from datetime import datetime

def _coerce_numero_serie_like_excel(x):
    # Mesma ideia de antes, mas agora x SEMPRE vem como string (porque daremos dtype=str ao ler)
    if pd.isna(x):
        return None
    s = str(x).strip()

    # remove .0 do final (padrão comum de export)
    if re.fullmatch(r"\d+(\.0+)?", s):
        s = s.split('.')[0]

    # só dígitos?
    if re.fullmatch(r"\d+", s or ""):
        if len(s) <= 15:
            try:
                return int(s)   # vira número no Excel (multiplicação por 1)
            except Exception:
                return s        # fallback texto
        else:
            return s            # > 15 dígitos: manter como texto pra não perder precisão
    return s if s else None

def _coerce_datetime_like_excel(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    if not s:
        return None
    fmts = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y",
    ]
    for fmt in fmts:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        return None if pd.isna(dt) else dt.to_pydatetime()
    except Exception:
        return None

def processar_itau_base(caminho_base):
    # 👉 LER COMO TEXTO (dtype=str) para as 3 primeiras colunas
    df = pd.read_excel(
        caminho_base,
        header=None,
        dtype={0: str, 1: str, 2: str}  # **ponto-chave**
    )

    # remove 4 primeiras linhas e pega 3 colunas
    df = df.iloc[4:, :3].copy()
    df.columns = ["Placa", "NumeroSerie", "UltimaDataHora"]

    # limpeza básica
    df["Placa"] = df["Placa"].astype(str).str.strip().str.upper()
    df = df[df["Placa"].str.lower() != "placa"]
    df = df[df["Placa"].notna() & (df["Placa"].str.strip() != "")]
    df["Placa"] = df["Placa"].astype(str).str.strip().str.upper()  # normaliza chave

    # tratamentos (emula Excel)
    df["NumeroSerie"] = df["NumeroSerie"].map(_coerce_numero_serie_like_excel)
    df["UltimaDataHora"] = df["UltimaDataHora"].map(_coerce_datetime_like_excel)

    print(f"Tabela tratada com sucesso: {caminho_base} | Linhas: {len(df)}")
    return df.reset_index(drop=True)