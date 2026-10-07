"""
Testes básicos. Correr com:
    python -m unittest -v
"""

import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from juntar_vendas import juntar_vendas, limpar_dados, listar_csvs, proteger_texto


def escrever(pasta: Path, nome: str, conteudo: str) -> None:
    (pasta / nome).write_text(conteudo, encoding="utf-8")


class TestJuntarVendas(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.pasta = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()  # apaga os dados de teste no fim

    def test_soma_correta_com_dois_ficheiros(self):
        escrever(self.pasta, "jan.csv", "data;produto;quantidade;preco_unitario\n05/01/2026;Caneca;2;6,50\n")
        escrever(self.pasta, "fev.csv", "data,produto,quantidade,preco_unitario\n03/02/2026,Caneca,1,6.50\n")
        saida = self.pasta / "r.xlsx"
        r = juntar_vendas(self.pasta, saida)
        self.assertEqual(r, {"ficheiros": 2, "validas": 2, "erros": 0})
        resumo = pd.read_excel(saida, sheet_name="Resumo")
        total = resumo.loc[resumo["indicador"] == "Total de vendas (€)", "valor"].iloc[0]
        self.assertAlmostEqual(total, 19.50)

    def test_linhas_invalidas_vao_para_erros(self):
        df = pd.DataFrame({
            "data": ["31/02/2026", "01/01/2026", "01/01/2026", "01/01/2026"],
            "produto": ["Caneca", "", "Boné", "Boné"],
            "quantidade": ["1", "1", "dois", "-3"],
            "preco_unitario": ["1", "1", "1", "1"],
        })
        validas, erros = limpar_dados(df)
        self.assertEqual(len(validas), 0)
        self.assertEqual(len(erros), 4)

    def test_coluna_em_falta_da_erro_claro(self):
        escrever(self.pasta, "mau.csv", "data;produto;quantidade\n05/01/2026;Caneca;2\n")
        with self.assertRaises(ValueError) as ctx:
            juntar_vendas(self.pasta, self.pasta / "r.xlsx")
        self.assertIn("preco_unitario", str(ctx.exception))

    def test_pasta_sem_csv(self):
        with self.assertRaises(FileNotFoundError):
            listar_csvs(self.pasta)

    def test_formula_nao_e_executada_no_excel(self):
        escrever(self.pasta, "x.csv",
                 'data;produto;quantidade;preco_unitario\n01/01/2026;=1+1;1;1\n')
        saida = self.pasta / "r.xlsx"
        juntar_vendas(self.pasta, saida)
        folha = load_workbook(saida)["Dados"]
        valores = [c.value for c in folha["B"]]  # coluna "produto"
        self.assertIn("'=1+1", valores)
        self.assertNotIn("=1+1", valores)

    def test_proteger_texto(self):
        self.assertEqual(proteger_texto("@cmd"), "'@cmd")
        self.assertEqual(proteger_texto("Caneca"), "Caneca")
        self.assertEqual(proteger_texto(-5), -5)  # números não são alterados


if __name__ == "__main__":
    unittest.main()
