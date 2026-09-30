import pandas as pd

dados = {
    "Nome": ["Ana ", "Bruno", "Carla", "Bruno", "Diego", None],
    "Departamento": ["Vendas", "TI", "vendas", "TI", "RH", "TI"],
    "Salario": [3000, 4500, None, 4500, 3800, 5000],
}

df = pd.DataFrame(dados)
df.to_excel("funcionarios_sujo.xlsx", index=False)
print("Planilha criada!")