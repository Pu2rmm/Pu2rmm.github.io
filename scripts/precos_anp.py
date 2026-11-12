"""Robô do AbasteX: preços médios da ANP por estado e por município.

Lê a página "Levantamento de preços de combustíveis - últimas semanas
pesquisadas", baixa a planilha mais recente do resumo semanal e grava um
JSON pequeno que o app lê:

{
  "fonte": "ANP",
  "de": "2026-11-02", "ate": "2026-11-08",
  "estados":    {"SP": {"etanol": 4.19, "gasolina": 6.29, ...}},
  "municipios": {"SP": {"SAO PAULO": {"etanol": 4.05, ...}}}
}

Os nomes de cidade ficam em maiúsculas e sem acento (do jeito da ANP), pro
app achar fácil.

Uso: python scripts/precos_anp.py abastex/precos.json
"""

import io
import json
import re
import sys
import unicodedata
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

PAGINA = (
    "https://www.gov.br/anp/pt-br/assuntos/precos-e-defesa-da-concorrencia/"
    "precos/levantamento-de-precos-de-combustiveis-ultimas-semanas-pesquisadas"
)
CABECALHO = {"User-Agent": "Mozilla/5.0 (robo-abastex; precos semanais)"}

# Produto da ANP -> chave no app
PRODUTOS = {
    "ETANOL HIDRATADO": "etanol",
    "GASOLINA COMUM": "gasolina",
    "GASOLINA ADITIVADA": "gasolina_aditivada",
    "OLEO DIESEL": "diesel_s500",
    "OLEO DIESEL S10": "diesel_s10",
    "GNV": "gnv",
}

SIGLAS = {
    "ACRE": "AC", "ALAGOAS": "AL", "AMAPA": "AP", "AMAZONAS": "AM",
    "BAHIA": "BA", "CEARA": "CE", "DISTRITO FEDERAL": "DF",
    "ESPIRITO SANTO": "ES", "GOIAS": "GO", "MARANHAO": "MA",
    "MATO GROSSO": "MT", "MATO GROSSO DO SUL": "MS", "MINAS GERAIS": "MG",
    "PARA": "PA", "PARAIBA": "PB", "PARANA": "PR", "PERNAMBUCO": "PE",
    "PIAUI": "PI", "RIO DE JANEIRO": "RJ", "RIO GRANDE DO NORTE": "RN",
    "RIO GRANDE DO SUL": "RS", "RONDONIA": "RO", "RORAIMA": "RR",
    "SANTA CATARINA": "SC", "SAO PAULO": "SP", "SERGIPE": "SE",
    "TOCANTINS": "TO",
}


def limpo(texto) -> str:
    """Maiúsculas, sem acento e sem espaço sobrando."""
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t).strip().upper()


def sigla(estado) -> str | None:
    e = limpo(estado)
    if len(e) == 2 and e in SIGLAS.values():
        return e
    return SIGLAS.get(e)


def achar_planilha() -> str:
    """O link da planilha mais recente do resumo semanal."""
    r = requests.get(PAGINA, headers=CABECALHO, timeout=60)
    r.raise_for_status()
    links = [
        urljoin(PAGINA, a["href"])
        for a in BeautifulSoup(r.text, "html.parser").find_all("a", href=True)
    ]
    planilhas = [l for l in links if re.search(r"\.xlsx?($|[/?])", l, re.I)]
    print(f"Planilhas na página: {len(planilhas)}")
    for l in planilhas[:15]:
        print("  ", l)
    # A do resumo semanal (preços médios); se não achar, a primeira
    resumo = [l for l in planilhas if re.search(r"resumo|semanal", l, re.I)]
    escolhida = (resumo or planilhas or [None])[0]
    if not escolhida:
        raise SystemExit("Nenhuma planilha encontrada na página da ANP.")
    print("Usando:", escolhida)
    return escolhida


def ler_aba(arquivo: bytes, nome: str) -> pd.DataFrame | None:
    """Lê a aba cujo nome contém [nome], achando a linha do cabeçalho."""
    abas = pd.read_excel(io.BytesIO(arquivo), sheet_name=None, header=None)
    print("Abas:", list(abas))
    for aba, bruto in abas.items():
        if nome not in limpo(aba):
            continue
        for i in range(min(30, len(bruto))):
            linha = [limpo(x) for x in bruto.iloc[i].tolist()]
            if "PRODUTO" in linha and any("PRECO MEDIO" in x for x in linha):
                df = bruto.iloc[i + 1 :].copy()
                df.columns = linha
                return df
    return None


def coluna(df: pd.DataFrame, *partes: str) -> str | None:
    for c in df.columns:
        if all(p in c for p in partes):
            return c
    return None


def numero(valor) -> float | None:
    try:
        v = float(str(valor).replace(",", "."))
    except ValueError:
        return None
    return round(v, 3) if v > 0 else None


def para_data(valor) -> str | None:
    """A data como AAAA-MM-DD (a planilha pode trazer data ou texto)."""
    try:
        texto = str(valor).strip()
        iso = re.match(r"\d{4}-\d{2}-\d{2}", texto) is not None
        return pd.to_datetime(valor, dayfirst=not iso).date().isoformat()
    except (ValueError, TypeError):
        return None


def main(saida: str) -> None:
    url = achar_planilha()
    arquivo = requests.get(url, headers=CABECALHO, timeout=120).content

    estados: dict = {}
    municipios: dict = {}
    de = ate = None

    for nome_aba, por_cidade in (("ESTADO", False), ("MUNICIPIO", True)):
        df = ler_aba(arquivo, nome_aba)
        if df is None:
            print(f"Aba de {nome_aba} não encontrada.")
            continue
        c_prod = coluna(df, "PRODUTO")
        c_preco = coluna(df, "PRECO MEDIO", "REVENDA") or coluna(df, "PRECO MEDIO")
        c_est = coluna(df, "ESTADO")
        c_mun = coluna(df, "MUNICIPIO")
        c_de, c_ate = coluna(df, "DATA INICIAL"), coluna(df, "DATA FINAL")
        print(f"{nome_aba}: produto={c_prod} preco={c_preco} estado={c_est} "
              f"municipio={c_mun}")
        for _, l in df.iterrows():
            produto = PRODUTOS.get(limpo(l.get(c_prod, "")))
            preco = numero(l.get(c_preco))
            uf = sigla(l.get(c_est, ""))
            if not produto or preco is None or not uf:
                continue
            if c_de and c_ate and de is None:
                de, ate = para_data(l[c_de]), para_data(l[c_ate])
            if por_cidade:
                cidade = limpo(l.get(c_mun, ""))
                if cidade:
                    municipios.setdefault(uf, {}).setdefault(cidade, {})[
                        produto] = preco
            else:
                estados.setdefault(uf, {})[produto] = preco

    if not estados:
        raise SystemExit("Não consegui ler os preços dos estados. "
                         "O formato da planilha da ANP pode ter mudado.")

    dados = {
        "fonte": "ANP",
        "de": de,
        "ate": ate,
        "estados": estados,
        "municipios": municipios,
    }
    with open(saida, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, separators=(",", ":"),
                  sort_keys=True)
    total_cidades = sum(len(c) for c in municipios.values())
    print(f"OK: {len(estados)} estados e {total_cidades} municípios "
          f"({de} a {ate}).")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "precos.json")
