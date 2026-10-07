"""
juntar_vendas.py
----------------
Junta vários ficheiros CSV de vendas (um por mês, por exemplo) num único
relatório Excel com resumos por mês e por produto.

Uso:
    python juntar_vendas.py --pasta dados_exemplo --saida relatorio_vendas.xlsx

Formato esperado de cada CSV (separador ";" ou ","):
    data;produto;quantidade;preco_unitario
    05/01/2026;T-shirt;3;12,50

Segurança (o que este script faz para proteger quem o usa):
  - Só lê ficheiros .csv dentro da pasta indicada (não segue outras pastas).
  - Recusa ficheiros demasiado grandes (limite configurável).
  - Nunca executa nada que venha dos ficheiros: só lê texto e números.
  - Protege o Excel contra "injeção de fórmulas" (texto que começa por =, +, -, @).
  - Linhas com erros não são apagadas em silêncio: vão para a folha "Erros".
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

# ----------------------------------------------------------------------
# Configurações
# ----------------------------------------------------------------------
COLUNAS_OBRIGATORIAS = ["data", "produto", "quantidade", "preco_unitario"]
TAMANHO_MAXIMO_MB = 20  # ficheiros maiores do que isto são recusados
CARACTERES_PERIGOSOS = ("=", "+", "-", "@", "\t", "\r")


# ----------------------------------------------------------------------
# 1. Encontrar os ficheiros
# ----------------------------------------------------------------------
def listar_csvs(pasta: Path) -> list[Path]:
    """Devolve a lista de ficheiros .csv da pasta, por ordem alfabética."""
    if not pasta.is_dir():
        raise FileNotFoundError(f"A pasta '{pasta}' não existe.")

    ficheiros = []
    for ficheiro in sorted(pasta.glob("*.csv")):
        # Ignora atalhos/links: só ficheiros normais
        if ficheiro.is_symlink() or not ficheiro.is_file():
            continue
        tamanho_mb = ficheiro.stat().st_size / (1024 * 1024)
        if tamanho_mb > TAMANHO_MAXIMO_MB:
            print(f"AVISO: '{ficheiro.name}' tem {tamanho_mb:.1f} MB e foi ignorado.")
            continue
        ficheiros.append(ficheiro)

    if not ficheiros:
        raise FileNotFoundError(f"Não encontrei ficheiros .csv em '{pasta}'.")
    return ficheiros


# ----------------------------------------------------------------------
# 2. Ler um ficheiro
# ----------------------------------------------------------------------
def ler_csv(ficheiro: Path) -> pd.DataFrame:
    """Lê um CSV como texto (tudo string) para depois validarmos nós."""
    df = pd.read_csv(
        ficheiro,
        sep=None,          # deteta sozinho se é ";" ou ","
        engine="python",
        dtype=str,         # lê tudo como texto: nada é interpretado à sorte
        encoding="utf-8-sig",
        keep_default_na=False,
    )
    # Normaliza nomes das colunas: minúsculas e sem espaços à volta
    df.columns = [c.strip().lower() for c in df.columns]

    em_falta = [c for c in COLUNAS_OBRIGATORIAS if c not in df.columns]
    if em_falta:
        raise ValueError(f"'{ficheiro.name}' não tem as colunas: {', '.join(em_falta)}")

    df = df[COLUNAS_OBRIGATORIAS].copy()
    df["ficheiro_origem"] = ficheiro.name
    df["linha_origem"] = range(2, len(df) + 2)  # +2: cabeçalho é a linha 1
    return df


# ----------------------------------------------------------------------
# 3. Limpar e validar
# ----------------------------------------------------------------------
def texto_para_numero(serie: pd.Series) -> pd.Series:
    """Converte '12,50' ou '12.50' em 12.5. O que não for número fica vazio (NaN)."""
    limpo = serie.str.strip().str.replace("€", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(limpo, errors="coerce")


def limpar_dados(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Separa as linhas boas das linhas com erro.
    Devolve (linhas_validas, linhas_com_erro).
    """
    df = df.copy()
    df["produto"] = df["produto"].str.strip()
    df["data"] = pd.to_datetime(df["data"].str.strip(), format="%d/%m/%Y", errors="coerce")
    df["quantidade"] = texto_para_numero(df["quantidade"])
    df["preco_unitario"] = texto_para_numero(df["preco_unitario"])

    # Cada regra marca o motivo do erro (a primeira que falhar)
    df["erro"] = ""
    regras = [
        (df["data"].isna(), "data inválida (usar DD/MM/AAAA)"),
        (df["produto"] == "", "produto em branco"),
        (df["quantidade"].isna(), "quantidade não é número"),
        (df["quantidade"] <= 0, "quantidade tem de ser maior que 0"),
        (df["preco_unitario"].isna(), "preço não é número"),
        (df["preco_unitario"] < 0, "preço negativo"),
    ]
    for condicao, motivo in regras:
        df.loc[condicao & (df["erro"] == ""), "erro"] = motivo

    com_erro = df[df["erro"] != ""].copy()
    validas = df[df["erro"] == ""].drop(columns="erro").copy()
    validas["total"] = (validas["quantidade"] * validas["preco_unitario"]).round(2)
    validas["mes"] = validas["data"].dt.to_period("M").astype(str)  # ex.: "2026-01"
    return validas, com_erro


# ----------------------------------------------------------------------
# 4. Resumos
# ----------------------------------------------------------------------
def criar_resumos(validas: pd.DataFrame) -> dict[str, pd.DataFrame]:
    por_mes = (
        validas.groupby("mes", as_index=False)
        .agg(vendas=("total", "sum"), unidades=("quantidade", "sum"))
    )
    por_produto = (
        validas.groupby("produto", as_index=False)
        .agg(vendas=("total", "sum"), unidades=("quantidade", "sum"))
        .sort_values("vendas", ascending=False)
    )
    mes_x_produto = validas.pivot_table(
        index="produto", columns="mes", values="total", aggfunc="sum", fill_value=0
    ).reset_index()
    return {"Por Mês": por_mes, "Por Produto": por_produto, "Mês x Produto": mes_x_produto}


# ----------------------------------------------------------------------
# 5. Segurança ao escrever o Excel
# ----------------------------------------------------------------------
def proteger_texto(valor):
    """
    Impede 'injeção de fórmulas': se um texto começa por =, +, -, @
    o Excel pode tentar executá-lo como fórmula. Pomos um apóstrofo à frente
    para ser mostrado como texto normal.
    """
    if isinstance(valor, str) and valor.startswith(CARACTERES_PERIGOSOS):
        return "'" + valor
    return valor


def proteger_tabela(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [proteger_texto(str(c)) for c in df.columns]
    for coluna in df.columns:
        if df[coluna].dtype == object or pd.api.types.is_string_dtype(df[coluna]):
            df[coluna] = df[coluna].map(proteger_texto)
    return df


def escrever_excel(caminho: Path, folhas: dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        for nome, tabela in folhas.items():
            proteger_tabela(tabela).to_excel(writer, sheet_name=nome, index=False)
            folha = writer.sheets[nome]
            # Ajusta a largura das colunas ao conteúdo (máx. 40)
            for coluna in folha.columns:
                largura = max(len(str(c.value or "")) for c in coluna) + 2
                folha.column_dimensions[coluna[0].column_letter].width = min(largura, 40)
            folha.freeze_panes = "A2"  # cabeçalho sempre visível


# ----------------------------------------------------------------------
# Programa principal
# ----------------------------------------------------------------------
def juntar_vendas(pasta: Path, saida: Path) -> dict:
    ficheiros = listar_csvs(pasta)
    tabelas = [ler_csv(f) for f in ficheiros]
    todos = pd.concat(tabelas, ignore_index=True)

    validas, com_erro = limpar_dados(todos)
    resumos = criar_resumos(validas)

    resumo_geral = pd.DataFrame({
        "indicador": ["Ficheiros lidos", "Linhas válidas", "Linhas com erro",
                      "Total de vendas (€)", "Unidades vendidas"],
        "valor": [len(ficheiros), len(validas), len(com_erro),
                  round(validas["total"].sum(), 2), int(validas["quantidade"].sum())],
    })

    dados = validas.copy()
    dados["data"] = dados["data"].dt.strftime("%d/%m/%Y")
    erros = com_erro[["ficheiro_origem", "linha_origem", "erro"]]

    folhas = {"Resumo": resumo_geral, **resumos, "Dados": dados, "Erros": erros}
    escrever_excel(saida, folhas)
    return {"ficheiros": len(ficheiros), "validas": len(validas), "erros": len(com_erro)}


def main() -> int:
    parser = argparse.ArgumentParser(description="Junta CSVs de vendas num relatório Excel.")
    parser.add_argument("--pasta", required=True, help="Pasta com os ficheiros .csv")
    parser.add_argument("--saida", default="relatorio_vendas.xlsx", help="Ficheiro Excel a criar")
    args = parser.parse_args()

    saida = Path(args.saida)
    if saida.suffix.lower() != ".xlsx":
        print("ERRO: o ficheiro de saída tem de terminar em .xlsx")
        return 1

    try:
        r = juntar_vendas(Path(args.pasta), saida)
    except (FileNotFoundError, ValueError) as erro:
        print(f"ERRO: {erro}")
        return 1
    except PermissionError:
        print(f"ERRO: não consigo escrever '{saida}'. Está aberto no Excel?")
        return 1

    print(f"Pronto! {r['ficheiros']} ficheiros, {r['validas']} linhas válidas, "
          f"{r['erros']} com erro → {saida}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
