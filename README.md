# Juntar Vendas CSV → Relatório Excel

Junta vários ficheiros CSV de vendas (por exemplo, um por mês) num único relatório Excel, com resumos prontos a usar. Linhas com erros não desaparecem em silêncio: aparecem numa folha própria, com o ficheiro e a linha onde estão.

> Todos os dados em `dados_exemplo/` são **fictícios** (uma pastelaria inventada).

## O que o relatório contém

| Folha | Conteúdo |
|---|---|
| Resumo | Ficheiros lidos, linhas válidas, linhas com erro, total de vendas, unidades |
| Por Mês | Vendas e unidades de cada mês |
| Por Produto | Vendas e unidades por produto, do mais vendido para o menos vendido |
| Mês x Produto | Tabela cruzada: produto nas linhas, meses nas colunas |
| Dados | Todas as linhas válidas juntas, com o ficheiro de origem |
| Erros | Linhas rejeitadas: ficheiro, número da linha e motivo |

## Instalação

Precisas de Python 3.11 ou mais recente.

```bash
python -m venv .venv
# Windows:      .venv\Scripts\activate
# Linux/macOS:  source .venv/bin/activate
pip install -r requirements.txt
```

## Como usar

```bash
python gerar_dados_exemplo.py          # cria 12 CSVs fictícios em dados_exemplo/
python juntar_vendas.py --pasta dados_exemplo --saida relatorio_vendas.xlsx
```

Exemplo de saída:

```
Pronto! 12 ficheiros, 425 linhas válidas, 3 com erro → relatorio_vendas.xlsx
```

### Formato dos CSV

O separador pode ser `;` ou `,`. Os preços podem usar `12,50` ou `12.50`, e as datas devem estar em `DD/MM/AAAA`.

```
data;produto;quantidade;preco_unitario
05/01/2026;Café;3;0,80
```

## Segurança

- Só lê ficheiros `.csv` da pasta indicada e ignora atalhos e ficheiros com mais de 20 MB.
- Lê tudo como texto e valida depois. Nunca executa nada que venha dos ficheiros.
- Protege contra **injeção de fórmulas no Excel**: um texto como `=HYPERLINK(...)` fica guardado como texto, e não como fórmula.
- Os dados ficam no teu computador. Não há ligação à internet.
- As dependências estão fixadas em `requirements.txt` e foram verificadas com `pip-audit` (sem vulnerabilidades conhecidas a 07/10/2026).

## Testes

```bash
python -m unittest -v
```

São 6 testes: soma certa com separadores diferentes, linhas inválidas, coluna em falta, pasta vazia, injeção de fórmula e proteção de texto.

## Problemas comuns

| Mensagem | Solução |
|---|---|
| `não tem as colunas: ...` | O cabeçalho do CSV tem de ter `data;produto;quantidade;preco_unitario` |
| `Está aberto no Excel?` | Fecha o relatório antigo no Excel e corre de novo |
| Muitas linhas em "Erros" com data inválida | Confirma que as datas estão em `DD/MM/AAAA` |
