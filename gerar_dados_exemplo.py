"""
Gera 12 ficheiros CSV de vendas FICTÍCIAS (um por mês de 2026) na pasta
dados_exemplo/. Inclui de propósito algumas linhas com erros, para mostrar
que o script principal os deteta.

Uso:
    python gerar_dados_exemplo.py
"""

import csv
import random
from pathlib import Path

PRODUTOS = {"Café": 0.80, "Pastel de nata": 1.30, "Tosta mista": 3.50, "Sumo natural": 2.75, "Bolo do dia": 2.20}
PASTA = Path("dados_exemplo")


def gerar():
    random.seed(42)  # mesma semente = mesmos dados sempre (bom para testes)
    PASTA.mkdir(exist_ok=True)

    for mes in range(1, 13):
        linhas = []
        for _ in range(random.randint(25, 40)):
            produto = random.choice(list(PRODUTOS))
            dia = random.randint(1, 28)
            preco = f"{PRODUTOS[produto]:.2f}".replace(".", ",")  # formato português
            linhas.append([f"{dia:02d}/{mes:02d}/2026", produto, random.randint(1, 8), preco])

        # Erros de propósito, em alguns meses
        if mes == 3:
            linhas.append(["31/02/2026", "Café", 2, "0,80"])        # data impossível
        if mes == 6:
            linhas.append(["10/06/2026", "Tosta mista", "dois", "3,50"])  # quantidade em texto
        if mes == 9:
            linhas.append(["15/09/2026", "=HYPERLINK(\"http://exemplo.invalid\")", 1, "1,00"])  # tentativa de fórmula
            linhas.append(["16/09/2026", "", 1, "2,20"])               # produto em branco

        with open(PASTA / f"vendas_2026_{mes:02d}.csv", "w", newline="", encoding="utf-8") as f:
            escritor = csv.writer(f, delimiter=";")
            escritor.writerow(["data", "produto", "quantidade", "preco_unitario"])
            escritor.writerows(linhas)

    print(f"Criados 12 ficheiros em '{PASTA}/'")


if __name__ == "__main__":
    gerar()
