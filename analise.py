import pandas as pd
import matplotlib.pyplot as plt

# Input dos parâmetros iniciais
parametro1_inicial = int(input("Digite o valor inicial de Parametro1: "))
parametro2_inicial = int(input("Digite o valor inicial de Parametro2: "))

# Carregar os valores reais da planilha
df = pd.read_excel('teste.xlsx')
valores = df['Valor'].tolist()

# Gerar Parametro1 e Parametro2 de forma incremental
parametro1 = [parametro1_inicial + i for i in range(len(valores))]
parametro2 = [parametro2_inicial + i for i in range(len(valores))]

# Verificação de erros
erros = []
for i in range(len(valores)):
    if not parametro1[i] <= valores[i] <= parametro2[i]:
        erros.append(f"Linha {i+1}: Valor {valores[i]} está fora do intervalo ({parametro1[i]} - {parametro2[i]})")

# Exibir resultado da verificação
if erros:
    print("🚫 Foram encontrados erros:")
    for erro in erros:
        print(erro)
else:
    print("✅ Todos os valores estão dentro dos parâmetros!")

# Criar o gráfico
x = range(len(valores))
plt.figure(figsize=(10, 6))
plt.plot(x, parametro2, color='red', label='Parametro2 (linha superior)')
plt.plot(x, parametro1, color='blue', label='Parametro1 (linha inferior)')
plt.scatter(x, valores, color='black', label='Valor (meio)', zorder=5)

plt.title('Gráfico com verificação de parâmetros')
plt.xlabel('Índice')
plt.ylabel('Valores')
plt.grid(True)
plt.legend()
plt.tight_layout()
plt.show()
