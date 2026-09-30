import pandas as pd

# 1. Ler a planilha
df = pd.read_excel("funcionarios_sujo.xlsx")
print("ANTES DA LIMPEZA:")
print(df)

# 2. Remover linhas onde o Nome está vazio
df = df.dropna(subset=["Nome"])

# 3. Tirar espaços sobrando no começo/fim do nome
df["Nome"] = df["Nome"].str.strip()

# 4. Padronizar o departamento (primeira letra maiúscula)
df["Departamento"] = df["Departamento"].str.capitalize()

# 5. Preencher salário vazio com a média
df["Salario"] = df["Salario"].fillna(df["Salario"].mean())

# 6. Remover duplicatas
df = df.drop_duplicates()

print("\nDEPOIS DA LIMPEZA:")
print(df)

# 7. Salvar em uma nova planilha
df.to_excel("funcionarios_limpo.xlsx", index=False)
print("\nPlanilha limpa salva!")