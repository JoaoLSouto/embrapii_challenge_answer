# Relatório de Projeção para 2027

## Objetivo
Estimar os três indicadores principais para 2027:
- número de projetos contratados no ano;
- valor total dos projetos contratados no ano;
- número de projetos concluídos no ano.

## Técnica estatística utilizada
A estimativa foi feita com base em:
- agregação dos dados mensais em totais anuais;
- exclusão de 2026 da modelagem, pois a planilha contém apenas 5 meses de 2026 (incompleto);
- regressão linear simples entre o ano e cada indicador.

Essa técnica é adequada ao objetivo porque o planejamento é anual, os dados históricos apresentam um comportamento de tendência ao longo do tempo e a regressão linear gera uma estimativa clara, replicável e fácil de explicar.

## Conjunto de dados utilizado
- Fonte: `Embrapii_seleção_analista_2026_questao03_Estimativa.xlsx`
- Aba: `dados`
- Dados mensais de 2015 a maio de 2026
- Período usado na modelagem: 2015 a 2025 (12 anos completos)
- 2026 não foi usado no ajuste porque contém apenas 5 meses de dados e pode distorcer a projeção.

## Análise exploratória dos dados
A base tem 137 registros mensais e mostra o comportamento dos três indicadores ao longo do tempo.

Principais observações:
- Há crescimento sustentado de novos projetos contratados entre 2015 e 2025, com destaque para o salto de 2024 para 2025.
- O valor total dos projetos também cresce de forma forte, saindo de cerca de R$ 131 milhões em 2015 para R$ 1,63 bilhão em 2025.
- O número de projetos concluídos é mais volátil: cresceu até 2023, depois caiu em 2024 e 2025, o que sugere defasagem entre contratação e conclusão.
- 2026 aparece com apenas 5 meses registrados e deve ser visto como dado parcial: são 307 contratos e R$ 533 milhões em valor, mas 0 projetos concluídos.

## Interpretação dos indicadores
- `Novos projetos contratados`: a tendência histórica indica continuidade do crescimento, o que apoia a projeção de cerca de 814 contratos em 2027.
- `Valor total dos projetos contratados`: a tendência de aumento de investimento sugere que 2027 pode se aproximar de R$ 1,45 bilhão, ainda que a incerteza seja maior devido à variabilidade dos valores anuais.
- `Projetos concluídos`: o comportamento mais irregular indica que este indicador pode depender de fatores de ciclo e conclusão de projetos iniciados em anos anteriores.

## Gráficos e suporte visual
Os gráficos gerados em `questao03_graficos.png` mostram:
- a série histórica anual dos indicadores completos (2015-2025);
- a linha de tendência linear usada para projeção;
- a projeção para 2027 com intervalo de confiança de 95%;
- a marcação de 2026 como dado parcial, não usado no ajuste.

## Resultados projetados para 2027
Os valores estimados para 2027 são:

- **Novos projetos contratados**: 814 (95% CI: 688 a 939)
- **Valor total dos projetos contratados**: R$ 1.447.110.422,13 (95% CI: R$ 1.062.536.706,52 a R$ 1.831.684.138,74)
- **Projetos concluídos**: 387 (95% CI: 220 a 554)

Esses resultados foram calculados a partir da tendência histórica observada nos totais anuais de 2015 a 2025.

## Limites de confiança e vieses
- As projeções incluem intervalos de confiança de 95% para cada indicador.
- O modelo assume que a tendência histórica continuará de forma linear.
- Não foram usados modelos mais complexos de sazonalidade, pois o objetivo do relatório é fornecer estimativas anuais claras.
- A principal fonte de viés é que o comportamento futuro pode ser impactado por mudanças de política, orçamento ou processos da Embrapii que não estão representados nos dados históricos.
- A exclusão de 2026 evita viés causado por dados parciais e por completude zero no indicador de projetos concluídos.

## Script replicável
O script em Python desenvolvido para gerar esses resultados é:
- `questao03_projecao.py`

Para executar:
```bash
python questao03_projecao.py
```

O script gera também os arquivos:
- `questao03_projecao_resultados.csv`
- `questao03_graficos.png`

## Conclusão
A projeção para 2027 sugere crescimento contínuo dos projetos contratados e do valor total contratado. O indicador de projetos concluídos segue uma tendência mais moderada, mas ainda positiva, com ampla margem de incerteza. A técnica usada é simples, transparente e adequada para apoiar o planejamento inicial.
