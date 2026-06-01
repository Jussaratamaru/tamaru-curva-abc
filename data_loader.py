"""Carrega vendas — local: pasta Vendas do BI; nuvem: Vendas/ deste repo."""

from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path

import pandas as pd

APP_ROOT = Path(__file__).resolve().parent
BI_VENDAS_DIR = APP_ROOT.parent / "Base de dados IA" / "Vendas"
LOCAL_VENDAS_DIR = APP_ROOT / "Vendas"

NUMERIC = [
    "Quantidade", "Vl. Unitário", "Vl. Mercadoria", "Vl. Total",
    "Receita Líquida de Vendas", "Resultado Líquido de Vendas",
]

GRUPOS_FORNECEDOR = ["SUMITOMO", "QUAKER", "INSIZE", "TAMARU", "OUTROS"]


def grupo_fornecedor(nome) -> str:
    """Mesma regra do BI: SUMITOMO, QUAKER (RJ+SP), INSIZE, TAMARU (+ Balzers/Topdrill), OUTROS."""
    if pd.isna(nome) or str(nome).strip() == "":
        return "OUTROS"
    n = str(nome).upper()
    if "QUAKER" in n:
        return "QUAKER"
    if "SUMITOMO" in n:
        return "SUMITOMO"
    if "INSIZE" in n:
        return "INSIZE"
    if "TAMARU" in n or "TOPDRILL" in n or "BALZERS" in n:
        return "TAMARU"
    return "OUTROS"


def _has_vendas_xlsx(folder: Path) -> bool:
    return any(
        f.suffix.lower() == ".xlsx"
        and not f.name.startswith("~$")
        and f.name != "Comissoes.xlsx"
        for f in folder.glob("*.xlsx")
    )


def resolve_vendas_dir() -> tuple[Path, str]:
    """
    Local (PC): usa Base de dados IA/Vendas — mesma pasta do BI.
    Nuvem: usa Vendas/ dentro deste repositório.
    """
    env = os.environ.get("VENDAS_DIR", "").strip()
    if env:
        p = Path(env)
        if p.is_dir():
            return p, "variável VENDAS_DIR"

    if BI_VENDAS_DIR.is_dir() and _has_vendas_xlsx(BI_VENDAS_DIR):
        return BI_VENDAS_DIR, "BI (Base de dados IA/Vendas)"

    return LOCAL_VENDAS_DIR, "pasta Vendas deste app (nuvem ou fallback)"


VENDAS_DIR, VENDAS_ORIGEM = resolve_vendas_dir()


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
        raise FileNotFoundError(
            f"Nenhum arquivo em {VENDAS_DIR}. "
            f"Coloque os Excel em: {BI_VENDAS_DIR}"
        )

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
    if "Fornecedor_Fantasia" in df.columns:
        df["Fornecedor_Grupo"] = df["Fornecedor_Fantasia"].apply(grupo_fornecedor)
    else:
        df["Fornecedor_Grupo"] = "OUTROS"
    return df.dropna(subset=["Emissão"])
