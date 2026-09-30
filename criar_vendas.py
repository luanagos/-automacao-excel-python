import pandas as pd

vendas = {
    "janeiro": {"Produto": ["Notebook", "Mouse", "Teclado"], "Quantidade": [5, 20, 12], "Preco": [3500, 50, 120]},
    "fevereiro": {"Produto": ["Notebook", "Mouse", "Monitor"], "Quantidade": [3, 25, 8], "Preco": [3500, 50, 900]},
    "marco": {"Produto": ["Teclado", "Monitor", "Mouse"], "Quantidade": [15, 6, 30], "Preco": [120, 900, 50]},
}

for mes, dados in vendas.items():
    pd.DataFrame(dados).to_excel(f"vendas_{mes}.xlsx", index=False)

print("Planilhas criadas!")