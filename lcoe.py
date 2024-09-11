import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import fsolve

# Carregar e ler o arquivo Excel pelo nome do arquivo
nome_arquivo = 'DadosUsinas.xlsx'
dados_usinas = pd.read_excel(nome_arquivo)

# Usina 1 é a Usina termoelétrica a carvão e Usina 2 = Usina termoelétrica a gás

# Extraindo os dados do arquivo Excel
capacidade_mw_usina1 = dados_usinas.iloc[0, 1]
custo_planta_kw_usina1 = dados_usinas.iloc[1, 1]
custo_opex_fixo_kw_ano_usina1 = dados_usinas.iloc[2, 1]
custo_opex_variavel_mwh_usina1 = dados_usinas.iloc[3, 1]
custo_combustivel_mwh_usina1 = dados_usinas.iloc[4, 1]
fator_capacidade_usina1 = dados_usinas.iloc[5, 1]
taxa_desconto_usina1 = dados_usinas.iloc[6, 1] / 100  # Dividido por 100 para converter em decimal
vida_util_usina1 = dados_usinas.iloc[7, 1]

capacidade_mw_usina2 = dados_usinas.iloc[0, 2]
custo_planta_kw_usina2 = dados_usinas.iloc[1, 2]
custo_opex_fixo_kw_ano_usina2 = dados_usinas.iloc[2, 2]
custo_opex_variavel_mwh_usina2 = dados_usinas.iloc[3, 2]
custo_combustivel_mwh_usina2 = dados_usinas.iloc[4, 2]
fator_capacidade_usina2 = dados_usinas.iloc[5, 2]
taxa_desconto_usina2 = dados_usinas.iloc[6, 2] / 100
vida_util_usina2 = dados_usinas.iloc[7, 2]

# Função para calcular o CRF (Capital Recovery Factor)
def calcular_crf(taxa_desconto, vida_util):
    return (taxa_desconto * (1 + taxa_desconto) * vida_util) / ((1 + taxa_desconto) * vida_util - 1)

# Função para calcular LCOE e outros custos
def calcular_lcoe_detalhado(capacidade_mw, custo_planta_kw, custo_opex_fixo_kw_ano, custo_opex_variavel_mwh, custo_combustivel_mwh, fator_capacidade, taxa_desconto, vida_util):
    horas_ano = 8760  # Número de horas em um ano
    capacidade_kw = capacidade_mw * 1000
    capex_total = capacidade_kw * custo_planta_kw
    opex_fixo_anual = capacidade_kw * custo_opex_fixo_kw_ano
    energia_anual_mwh = capacidade_mw * fator_capacidade * horas_ano
    crf = calcular_crf(taxa_desconto, vida_util)
    opex_variavel_total = custo_opex_variavel_mwh + custo_combustivel_mwh
    lcoe = (capex_total * crf + opex_fixo_anual + opex_variavel_total * energia_anual_mwh) / energia_anual_mwh

    # Preparando uma lista com os resultados por ano
    anos = np.arange(1, vida_util + 1)
    tabela_detalhada = []

    custo_total = 0

    for ano in anos:
        custo_anual_combustivel = custo_combustivel_mwh * energia_anual_mwh
        custo_anual_opex_variavel = custo_opex_variavel_mwh * energia_anual_mwh
        total_opex = opex_fixo_anual + custo_anual_opex_variavel + custo_anual_combustivel
        custo_total += capex_total if ano == 1 else 0 + total_opex
        tabela_detalhada.append({
            "Ano": ano,
            "CAPEX Total ($)": capex_total if ano == 1 else 0,
            "OPEX Fixo Anual ($)": opex_fixo_anual,
            "OPEX Variável Anual ($)": custo_anual_opex_variavel,
            "Custo Combustível Anual ($)": custo_anual_combustivel,
            "Energia Anual (MWh)": energia_anual_mwh,
            "Total OPEX Anual ($)": total_opex,
            "Custo Total Anual ($)": custo_total
        })

    return pd.DataFrame(tabela_detalhada), lcoe, custo_total

# Função para encontrar a interseção entre duas curvas
def encontrar_intersecao(fatores_capacidade, lcoe1, lcoe2):
    # Subtrai os LCOEs para encontrar o ponto onde a diferença é zero (interseção)
    func_diferenca = lambda fc: np.interp(fc, fatores_capacidade, lcoe1) - np.interp(fc, fatores_capacidade, lcoe2)
    # Encontrar o fator de capacidade onde a interseção ocorre
    fator_cap_intersecao = fsolve(func_diferenca, 0.5)[0]  # Usar 0.5 como chute inicial
    return fator_cap_intersecao

# Calcular tabela detalhada, LCOE e custo total para a Usina 1
tabela_detalhada_usina1, lcoe_usina1, custo_total_usina1 = calcular_lcoe_detalhado(capacidade_mw_usina1, custo_planta_kw_usina1, custo_opex_fixo_kw_ano_usina1, custo_opex_variavel_mwh_usina1, custo_combustivel_mwh_usina1, fator_capacidade_usina1, taxa_desconto_usina1, vida_util_usina1)

# Calcular tabela detalhada, LCOE e custo total para a Usina 2
tabela_detalhada_usina2, lcoe_usina2, custo_total_usina2 = calcular_lcoe_detalhado(capacidade_mw_usina2, custo_planta_kw_usina2, custo_opex_fixo_kw_ano_usina2, custo_opex_variavel_mwh_usina2, custo_combustivel_mwh_usina2, fator_capacidade_usina2, taxa_desconto_usina2, vida_util_usina2)

# Exibir as tabelas detalhadas
print("Tabela de custos detalhada para a Usina termoelétrica a carvão:")
print(tabela_detalhada_usina1)

print("\nTabela de custos detalhada para a Usina termoelétrica a gás:")
print(tabela_detalhada_usina2)

# Exibir LCOEs
print(f"\nLCOE da Usina termoelétrica a carvão: ${lcoe_usina1:.2f}/MWh")
print(f"LCOE da Usina termoelétrica a gás: ${lcoe_usina2:.2f}/MWh")

# Exibir Custos Totais
print(f"\nCusto Total da Usina termoelétrica a carvão: ${custo_total_usina1:.2f}")
print(f"Custo Total da Usina termoelétrica a gás: ${custo_total_usina2:.2f}")

# Comparar os LCOEs
if lcoe_usina1 < lcoe_usina2:
    print("\nA Usina termoelétrica a carvão tem um custo nivelado de energia mais baixo.")
else:
    print("\nA Usina termoelétrica a gás tem um custo nivelado de energia mais baixo.")

# Comparar os custos totais
if custo_total_usina1 < custo_total_usina2:
    print("\nA Usina termoelétrica a carvão tem um custo total mais baixo.")
else:
    print("\nA Usina termoelétrica a gás tem um custo total mais baixo.")

# Geração de uma lista de fatores de capacidade (de 0.1 até 1 com 100 pontos)
fatores_capacidade = np.linspace(0.1, 1, 100)

# Calculando LCOE e custo total para cada fator de capacidade para as duas usinas
lcoe_usina1_var = [calcular_lcoe_detalhado(capacidade_mw_usina1, custo_planta_kw_usina1, custo_opex_fixo_kw_ano_usina1, custo_opex_variavel_mwh_usina1, custo_combustivel_mwh_usina1, fc, taxa_desconto_usina1, vida_util_usina1)[1] for fc in fatores_capacidade]
lcoe_usina2_var = [calcular_lcoe_detalhado(capacidade_mw_usina2, custo_planta_kw_usina2, custo_opex_fixo_kw_ano_usina2, custo_opex_variavel_mwh_usina2, custo_combustivel_mwh_usina2, fc, taxa_desconto_usina2, vida_util_usina2)[1] for fc in fatores_capacidade]

# Encontrar o ponto de interseção entre as duas curvas
fator_cap_intersecao = encontrar_intersecao(fatores_capacidade, lcoe_usina1_var, lcoe_usina2_var)
lcoe_intersecao = np.interp(fator_cap_intersecao, fatores_capacidade, lcoe_usina1_var)

# Plotar gráficos
plt.figure(figsize=(14, 6))

# Gráfico de LCOE
plt.subplot(1, 2, 1)
plt.plot(fatores_capacidade, lcoe_usina1_var, label='Usina termoelétrica a carvão', color='blue')
plt.plot(fatores_capacidade, lcoe_usina2_var, label='Usina termoelétrica a gás', color='red')
# Adicionar ponto de interseção
plt.scatter(fator_cap_intersecao, lcoe_intersecao, color='black', zorder=5, label=f'Interseção ({fator_cap_intersecao:.2f}, {lcoe_intersecao:.2f})')
plt.xlabel('Fator de Capacidade')
plt.ylabel('LCOE ($/MWh)')
plt.title('LCOE x Fator de Capacidade')
plt.legend()
plt.grid(True)

# Gráfico de Custo Total (sem alteração)
plt.subplot(1, 2, 2)
plt.plot(fatores_capacidade, lcoe_usina1_var, label='Usina termoelétrica a carvão', color='blue')
plt.plot(fatores_capacidade, lcoe_usina2_var, label='Usina termoelétrica a gás', color='red')
plt.xlabel('Fator de Capacidade')
plt.ylabel('Custo Total ($)')
plt.title('Custo Total x Fator de Capacidade')
plt.legend()
plt.grid(True)

# Mostrar os gráficos
plt.tight_layout()
plt.show()