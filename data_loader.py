"""Carrega vendas da pasta Vendas/ (mesma raiz do app)."""

from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
VENDAS_DIR = ROOT / "Vendas"

NUMERIC = [
    "Quantidade", "Vl. Unitário", "Vl. Mercadoria", "Vl. Total",
    "Receita Líquida de Vendas", "Resultado Líquido de Vendas",
]


def _list_vendas_files() -> list[Path]:
    return sorted(
        f for f in VENDAS_DIR.glob("*.xlsx")
        if not f.name.startswith("~$") and f.name != "Comissoes.xlsx"
    )


def ultima_atualizacao_vendas() -> tuple[datetime, str]:
    """Data/hora de modificação do arquivo de vendas mais recente."""
    files = _list_vendas_files()
    if not files:
        raise FileNotFoundError(f"Nenhum arquivo em {VENDAS_DIR}")
    latest = max(files, key=lambda f: f.stat().st_mtime)
    return datetime.fromtimestamp(latest.stat().st_mtime), latest.name


def _parse_date(series: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    return pd.to_datetime(series, dayfirst=True, errors="coerce")


def load_vendas() -> pd.DataFrame:
    files = _list_vendas_files()
    if not files:
        raise FileNotFoundError(f"Nenhum arquivo em {VENDAS_DIR}")

    df = pd.concat([pd.read_excel(f) for f in files], ignore_index=True)
    df.columns = df.columns.str.strip()
    df["Emissão"] = _parse_date(df["Emissão"])

    for col in NUMERIC:
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(",", ".").str.strip(),
                errors="coerce",
            )

    df["Item"] = df["Item"].astype(str).str.strip().str.split(".").str[0]
    df["Cod.Cliente"] = df["Cod.Cliente"].astype(str).str.strip().str.split(".").str[0]
    df["AnoMes_str"] = df["Emissão"].dt.strftime("%Y-%m")
    return df.dropna(subset=["Emissão"])
