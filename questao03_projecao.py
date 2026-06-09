from __future__ import annotations

import argparse
import base64
import pathlib

import matplotlib.pyplot as plt
import pandas as pd
import statsmodels.api as sm
from matplotlib.ticker import FuncFormatter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import (
    Image as PdfImage,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


ARQUIVO_ENTRADA_PADRAO = pathlib.Path(
    "Embrapii_seleção_analista_2026_questao03_Estimativa.xlsx"
)
ARQUIVO_RESULTADOS_PADRAO = pathlib.Path("questao03_projecao_resultados.csv")
INDICADOR_VALOR = "valor_projetos_contratados"

INDICADORES: tuple[dict[str, str], ...] = (
    {
        "coluna": "novos_projetos_contratados",
        "titulo": "Novos projetos contratados",
        "unidade": "unidades",
    },
    {
        "coluna": INDICADOR_VALOR,
        "titulo": "Valor total dos projetos contratados",
        "unidade": "R$",
    },
    {
        "coluna": "projetos_concluidos",
        "titulo": "Projetos concluídos",
        "unidade": "unidades",
    },
)

COLUNAS_INDICADORES = tuple(indicador["coluna"] for indicador in INDICADORES)
TITULOS_INDICADORES = {
    indicador["coluna"]: indicador["titulo"] for indicador in INDICADORES
}
COLUNAS_OBRIGATORIAS = ("ano", "mes", *COLUNAS_INDICADORES)
COLUNAS_PROJECAO = (
    "ano",
    "indicador",
    "estimativa",
    "limite_inferior_95",
    "limite_superior_95",
    "coeficiente_ano",
    "constante",
    "r2",
)
COLUNAS_FORMATAVEIS_PROJECAO = (
    "estimativa",
    "limite_inferior_95",
    "limite_superior_95",
    "coeficiente_ano",
    "constante",
)


def carregar_dados(arquivo: pathlib.Path) -> pd.DataFrame:
    df = pd.read_excel(arquivo, sheet_name="dados", engine="openpyxl")
    df = df.rename(columns=str.strip)
    validar_colunas(df, COLUNAS_OBRIGATORIAS)

    for coluna in COLUNAS_OBRIGATORIAS:
        df[coluna] = pd.to_numeric(df[coluna], errors="raise")

    return df


def validar_colunas(df: pd.DataFrame, colunas: tuple[str, ...]) -> None:
    faltantes = [coluna for coluna in colunas if coluna not in df.columns]
    if faltantes:
        raise ValueError(f"Colunas ausentes na planilha: {', '.join(faltantes)}")


def obter_totais_anuais(df: pd.DataFrame) -> pd.DataFrame:
    totais = (
        df.groupby("ano", as_index=False)
        .agg(
            novos_projetos_contratados=("novos_projetos_contratados", "sum"),
            valor_projetos_contratados=("valor_projetos_contratados", "sum"),
            projetos_concluidos=("projetos_concluidos", "sum"),
            meses_registrados=("mes", "nunique"),
        )
        .sort_values("ano")
    )
    totais["ano_completo"] = totais["meses_registrados"] == 12
    return totais


def selecionar_dados_modelagem(
    totais: pd.DataFrame, ano_previsao: int
) -> pd.DataFrame:
    ano_maximo_modelagem = ano_previsao - 2
    dados_modelagem = totais.loc[
        totais["ano_completo"] & (totais["ano"] <= ano_maximo_modelagem)
    ].copy()

    if len(dados_modelagem) < 3:
        raise ValueError(
            "São necessários pelo menos 3 anos completos para estimar o modelo."
        )

    colunas_modelo = ("ano", *COLUNAS_INDICADORES)
    colunas_com_nulos = [
        coluna for coluna in colunas_modelo if dados_modelagem[coluna].isna().any()
    ]
    if colunas_com_nulos:
        raise ValueError(
            "Há valores ausentes nas colunas usadas no modelo: "
            + ", ".join(colunas_com_nulos)
        )

    return dados_modelagem


def obter_periodo_modelagem(
    totais: pd.DataFrame, ano_previsao: int
) -> tuple[int, int]:
    dados_modelagem = selecionar_dados_modelagem(totais, ano_previsao)
    return int(dados_modelagem["ano"].min()), int(dados_modelagem["ano"].max())


def obter_anos_fora_modelagem(
    totais: pd.DataFrame, ano_previsao: int
) -> list[int]:
    _, ano_final = obter_periodo_modelagem(totais, ano_previsao)
    anos = totais.loc[totais["ano"] > ano_final, "ano"].astype(int).tolist()
    return anos


def montar_observacao_modelagem(totais: pd.DataFrame, ano_previsao: int) -> str:
    anos_fora_modelagem = obter_anos_fora_modelagem(totais, ano_previsao)
    if not anos_fora_modelagem:
        return "Todos os anos elegíveis disponíveis foram usados no ajuste."

    anos_texto = ", ".join(str(ano) for ano in anos_fora_modelagem)
    return (
        f"Anos posteriores ao período usado ({anos_texto}) são exibidos no "
        "histórico, quando existirem, mas não entram no ajuste."
    )


def montar_matriz_anos(anos: pd.Series | list[int] | range) -> pd.DataFrame:
    serie_anos = pd.Series(anos, dtype="float64").reset_index(drop=True)
    return sm.add_constant(pd.DataFrame({"ano": serie_anos}), has_constant="add")


def ajustar_modelo_linear(
    anos: pd.Series, valores: pd.Series
) -> sm.regression.linear_model.RegressionResultsWrapper:
    y = pd.Series(valores, dtype="float64").reset_index(drop=True)
    modelo = sm.OLS(y, montar_matriz_anos(anos)).fit()
    return modelo


def projetar_ano(
    modelo: sm.regression.linear_model.RegressionResultsWrapper, ano: int
) -> dict[str, float]:
    forecast = modelo.get_prediction(montar_matriz_anos([ano])).summary_frame(
        alpha=0.05
    )
    linha = forecast.iloc[0]
    return {
        "ano": int(ano),
        "estimativa": float(linha["mean"]),
        "limite_inferior_95": float(linha["mean_ci_lower"]),
        "limite_superior_95": float(linha["mean_ci_upper"]),
    }


def calcular_projecoes(
    totais: pd.DataFrame, ano_previsao: int = 2027
) -> tuple[
    pd.DataFrame, dict[str, sm.regression.linear_model.RegressionResultsWrapper]
]:
    dados_modelagem = selecionar_dados_modelagem(totais, ano_previsao)
    modelos: dict[str, sm.regression.linear_model.RegressionResultsWrapper] = {}
    resultados: list[dict[str, object]] = []

    anos = dados_modelagem["ano"].astype(float)
    for indicador in COLUNAS_INDICADORES:
        modelo = ajustar_modelo_linear(anos, dados_modelagem[indicador])
        modelos[indicador] = modelo

        previsao = projetar_ano(modelo, ano_previsao)
        resultados.append(
            {
                **previsao,
                "indicador": indicador,
                "coeficiente_ano": float(modelo.params["ano"]),
                "constante": float(modelo.params["const"]),
                "r2": float(modelo.rsquared),
            }
        )

    return pd.DataFrame(resultados, columns=COLUNAS_PROJECAO), modelos


def formatar_brl(valor: float | int) -> str:
    if pd.isna(valor):
        return ""

    sinal = "-" if valor < 0 else ""
    inteiro, decimais = f"{abs(valor):.2f}".split(".")
    partes: list[str] = []
    for i, digito in enumerate(reversed(inteiro)):
        if i and i % 3 == 0:
            partes.append(".")
        partes.append(digito)
    return f"R$ {sinal}{''.join(reversed(partes))},{decimais}"


def formatar_valor_por_indicador(valor: float, indicador: str) -> str:
    if indicador == INDICADOR_VALOR:
        return formatar_brl(valor)
    return f"{valor:.2f}"


def formatar_moeda_eixo(valor: float, _posicao: int) -> str:
    if abs(valor) >= 1_000_000:
        return f"R$ {valor / 1_000_000:.0f} mi"
    return formatar_brl(valor).replace(",00", "")


def formatar_historico_para_exibicao(totais: pd.DataFrame) -> pd.DataFrame:
    historico = totais.copy()
    historico["valor_projetos_contratados"] = historico[
        "valor_projetos_contratados"
    ].map(formatar_brl)
    historico["ano_completo"] = historico["ano_completo"].map(
        {True: "Sim", False: "Não"}
    )
    return historico.rename(
        columns={
            "ano": "Ano",
            "novos_projetos_contratados": "Novos projetos contratados",
            "valor_projetos_contratados": "Valor total contratado",
            "projetos_concluidos": "Projetos concluídos",
            "meses_registrados": "Meses registrados",
            "ano_completo": "Ano completo",
        }
    )


def formatar_projecoes_para_exibicao(projecoes: pd.DataFrame) -> pd.DataFrame:
    tabela = projecoes.copy()
    for campo in COLUNAS_FORMATAVEIS_PROJECAO:
        tabela[campo] = tabela.apply(
            lambda row, c=campo: formatar_valor_por_indicador(row[c], row["indicador"]),
            axis=1,
        )
    tabela["indicador"] = tabela["indicador"].map(TITULOS_INDICADORES)
    tabela["r2"] = tabela["r2"].map(lambda valor: f"{valor:.3f}")
    return tabela.rename(
        columns={
            "ano": "Ano",
            "indicador": "Indicador",
            "estimativa": "Estimativa",
            "limite_inferior_95": "Limite inferior 95%",
            "limite_superior_95": "Limite superior 95%",
            "coeficiente_ano": "Coeficiente anual",
            "constante": "Constante",
            "r2": "R²",
        }
    )


def salvar_resultados(df: pd.DataFrame, arquivo_saida: pathlib.Path) -> None:
    df.to_csv(arquivo_saida, index=False, encoding="utf-8-sig")


def salvar_relatorio_html(
    totais: pd.DataFrame,
    projecoes: pd.DataFrame,
    imagem_path: pathlib.Path,
    arquivo_saida: pathlib.Path,
    ano_previsao: int = 2027,
) -> None:
    imagem_base64 = base64.b64encode(imagem_path.read_bytes()).decode("utf-8")
    historico_html = formatar_historico_para_exibicao(totais).to_html(
        index=False,
        classes="table table-striped table-sm",
        border=0,
    )
    projecoes_html = formatar_projecoes_para_exibicao(projecoes).to_html(
        index=False,
        classes="table table-bordered table-sm",
        border=0,
    )
    ano_inicio, ano_fim = obter_periodo_modelagem(totais, ano_previsao)
    observacao_modelagem = montar_observacao_modelagem(totais, ano_previsao)

    html = f"""
<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <title>Relatório de Projeção Embrapii {ano_previsao}</title>
  <style>
    @page {{ size: A4; margin: 20mm; }}
    body {{ font-family: Arial, sans-serif; margin: 0 auto; max-width: 210mm; color: #222; }}
    h1, h2, h3 {{ color: #2f4f4f; }}
    .section {{ margin-bottom: 30px; padding: 20px; }}
    .table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; }}
    .table th, .table td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
    .table th {{ background: #f8f8f8; }}
    .info-box {{ background: #f1f9ff; padding: 12px; border-left: 4px solid #1f77b4; margin-bottom: 20px; }}
    .button {{ display: inline-block; margin: 10px 0; padding: 10px 18px; background: #1f77b4; color: #fff; text-decoration: none; border-radius: 4px; }}
    .button:hover {{ background: #165a8e; }}
    img {{ max-width: 100%; height: auto; border: 1px solid #ddd; }}
    @media print {{
      .button {{ display: none; }}
      body {{ margin: 0; }}
      .section {{ padding: 0; }}
    }}
  </style>
</head>
<body>
  <h1>Relatório automático de projeção {ano_previsao}</h1>
  <div class="section info-box">
    <p><strong>Fonte dos dados:</strong> {ARQUIVO_ENTRADA_PADRAO.name}, aba <code>dados</code>.</p>
    <p><strong>Período usado na modelagem:</strong> {ano_inicio} a {ano_fim} (anos completos).</p>
    <p><strong>Observação:</strong> {observacao_modelagem}</p>
    <a class="button" href="questao03_relatorio.pdf" download>Baixar relatório em PDF</a>
  </div>

  <div class="section">
    <h2>1. Resumo histórico</h2>
    {historico_html}
  </div>

  <div class="section">
    <h2>2. Projeções para {ano_previsao}</h2>
    {projecoes_html}
  </div>

  <div class="section">
    <h2>3. Gráfico de tendência</h2>
    <p>O gráfico mostra o histórico, a linha de tendência e a projeção para {ano_previsao}.</p>
    <img src="data:image/png;base64,{imagem_base64}" alt="Gráfico de projeção" />
  </div>

  <div class="section">
    <h2>4. Observações</h2>
    <ul>
      <li>O modelo usa regressão linear simples sobre os anos completos selecionados.</li>
      <li>Os intervalos de confiança de 95% indicam a incerteza estatística da estimativa.</li>
      <li>Indicadores com maior variabilidade histórica tendem a gerar intervalos mais amplos.</li>
    </ul>
  </div>
</body>
</html>
"""

    arquivo_saida.write_text(html, encoding="utf-8")


def criar_tabela_pdf(dados: list[list[object]], tamanho_fonte: int = 8) -> Table:
    tabela = Table(dados, repeatRows=1, hAlign="LEFT")
    tabela.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("FONTSIZE", (0, 0), (-1, -1), tamanho_fonte),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return tabela


def montar_dados_historico_pdf(totais: pd.DataFrame) -> list[list[object]]:
    dados = [
        [
            "Ano",
            "Novos projetos",
            "Valor contratado",
            "Projetos concluídos",
            "Meses",
            "Completo",
        ]
    ]
    for _, row in totais.iterrows():
        dados.append(
            [
                int(row["ano"]),
                int(row["novos_projetos_contratados"]),
                formatar_brl(row["valor_projetos_contratados"]),
                int(row["projetos_concluidos"]),
                int(row["meses_registrados"]),
                "Sim" if row["ano_completo"] else "Não",
            ]
        )
    return dados


def montar_dados_projecoes_pdf(
    projecoes: pd.DataFrame, ano_previsao: int
) -> list[list[object]]:
    dados = [
        [
            "Indicador",
            f"Estimativa {ano_previsao}",
            "Inferior 95%",
            "Superior 95%",
            "Coef. anual",
            "Constante",
            "R²",
        ]
    ]
    for _, row in projecoes.iterrows():
        dados.append(
            [
                TITULOS_INDICADORES[row["indicador"]],
                formatar_valor_por_indicador(row["estimativa"], row["indicador"]),
                formatar_valor_por_indicador(
                    row["limite_inferior_95"], row["indicador"]
                ),
                formatar_valor_por_indicador(
                    row["limite_superior_95"], row["indicador"]
                ),
                formatar_valor_por_indicador(row["coeficiente_ano"], row["indicador"]),
                formatar_valor_por_indicador(row["constante"], row["indicador"]),
                f"{row['r2']:.3f}",
            ]
        )
    return dados


def salvar_relatorio_pdf(
    totais: pd.DataFrame,
    projecoes: pd.DataFrame,
    imagem_path: pathlib.Path,
    arquivo_saida: pathlib.Path,
    ano_previsao: int = 2027,
) -> None:
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "Title", parent=styles["Title"], alignment=1, fontSize=18, spaceAfter=14
    )
    heading = ParagraphStyle(
        "Heading", parent=styles["Heading2"], fontSize=14, spaceAfter=10
    )
    normal = ParagraphStyle("Normal", parent=styles["Normal"], fontSize=10, leading=14)

    doc = SimpleDocTemplate(
        str(arquivo_saida),
        pagesize=A4,
        leftMargin=30,
        rightMargin=30,
        topMargin=30,
        bottomMargin=30,
    )
    story: list[object] = []
    ano_inicio, ano_fim = obter_periodo_modelagem(totais, ano_previsao)

    story.append(
        Paragraph(f"Relatório automático de projeção {ano_previsao}", title_style)
    )
    story.append(
        Paragraph(
            f"Fonte dos dados: {ARQUIVO_ENTRADA_PADRAO.name}, aba <b>dados</b>.",
            normal,
        )
    )
    story.append(
        Paragraph(
            f"Período usado na modelagem: {ano_inicio} a {ano_fim} "
            f"(anos completos). {montar_observacao_modelagem(totais, ano_previsao)}",
            normal,
        )
    )
    story.append(Spacer(1, 12))

    story.append(Paragraph("1. Resumo histórico", heading))
    story.append(criar_tabela_pdf(montar_dados_historico_pdf(totais), tamanho_fonte=7))
    story.append(Spacer(1, 12))

    story.append(Paragraph(f"2. Projeções para {ano_previsao}", heading))
    story.append(
        criar_tabela_pdf(
            montar_dados_projecoes_pdf(projecoes, ano_previsao), tamanho_fonte=7
        )
    )
    story.append(PageBreak())

    story.append(Paragraph("3. Observações", heading))
    story.append(
        Paragraph(
            "O modelo usa regressão linear simples sobre os anos completos "
            "selecionados.",
            normal,
        )
    )
    story.append(
        Paragraph(
            "Os intervalos de confiança de 95% indicam a incerteza estatística "
            "da estimativa.",
            normal,
        )
    )
    story.append(Spacer(1, 12))

    try:
        img = PdfImage(str(imagem_path))
        max_width = 520
        max_height = 450
        escala = min(max_width / img.drawWidth, max_height / img.drawHeight, 1.0)
        img.drawWidth *= escala
        img.drawHeight *= escala
        story.append(img)
    except OSError as erro:
        print(f"[AVISO] Gráfico não pôde ser carregado no PDF: {erro}")
        story.append(Paragraph("Gráfico não pôde ser carregado no PDF.", normal))

    doc.build(story)


def salvar_graficos(
    totais: pd.DataFrame,
    projecoes: pd.DataFrame,
    modelos: dict[str, sm.regression.linear_model.RegressionResultsWrapper],
    arquivo_saida: pathlib.Path,
) -> None:
    ano_previsao = int(projecoes["ano"].iloc[0])
    historico_modelagem = selecionar_dados_modelagem(totais, ano_previsao)
    anos_modelagem = set(historico_modelagem["ano"].astype(int))
    dados_fora_modelagem = totais.loc[~totais["ano"].astype(int).isin(anos_modelagem)]

    fig, axes = plt.subplots(
        nrows=len(INDICADORES),
        ncols=1,
        figsize=(12, 15),
        constrained_layout=True,
    )
    fig.suptitle(
        f"Projeção de indicadores Embrapii para {ano_previsao}",
        fontsize=16,
        weight="bold",
    )

    for ax, indicador in zip(axes, INDICADORES):
        coluna = indicador["coluna"]
        modelo = modelos[coluna]

        anos_tendencia = range(
            int(historico_modelagem["ano"].min()), ano_previsao + 1
        )
        tendencia = modelo.predict(montar_matriz_anos(anos_tendencia))

        ax.plot(
            historico_modelagem["ano"],
            historico_modelagem[coluna],
            marker="o",
            linestyle="-",
            label=(
                f"Histórico usado "
                f"{int(historico_modelagem['ano'].min())}-"
                f"{int(historico_modelagem['ano'].max())}"
            ),
            color="#1f77b4",
        )
        ax.plot(
            list(anos_tendencia),
            tendencia,
            linestyle="--",
            color="#2ca02c",
            label="Tendência linear",
        )

        if not dados_fora_modelagem.empty:
            anos_fora = ", ".join(
                str(ano) for ano in dados_fora_modelagem["ano"].astype(int)
            )
            ax.scatter(
                dados_fora_modelagem["ano"],
                dados_fora_modelagem[coluna],
                marker="X",
                color="#7f7f7f",
                s=80,
                zorder=4,
                label=f"Fora do ajuste ({anos_fora})",
            )

        linha_previsao = projecoes.loc[projecoes["indicador"] == coluna].iloc[0]
        estimativa = linha_previsao["estimativa"]
        inferior = linha_previsao["limite_inferior_95"]
        superior = linha_previsao["limite_superior_95"]
        erro_intervalo = [[estimativa - inferior], [superior - estimativa]]

        ax.errorbar(
            [ano_previsao],
            [estimativa],
            yerr=erro_intervalo,
            fmt="o",
            color="#d62728",
            ecolor="#d62728",
            elinewidth=1.5,
            capsize=5,
            markersize=7,
            zorder=6,
            label=f"Projeção {ano_previsao} (IC 95%)",
        )

        ax.set_title(indicador["titulo"], fontsize=14)
        ax.set_xlabel("Ano")
        ax.set_ylabel(indicador["unidade"])
        if coluna == INDICADOR_VALOR:
            ax.yaxis.set_major_formatter(FuncFormatter(formatar_moeda_eixo))
        ax.set_xlim(historico_modelagem["ano"].min() - 0.5, ano_previsao + 0.5)
        ax.grid(True, linestyle="--", alpha=0.4)
        ax.legend()

    fig.savefig(arquivo_saida, dpi=200, bbox_inches="tight")
    plt.close(fig)


def imprimir_relatorio(
    totais: pd.DataFrame, projecoes: pd.DataFrame, ano_previsao: int
) -> None:
    dados_modelagem = selecionar_dados_modelagem(totais, ano_previsao)
    anos_fora_modelagem = obter_anos_fora_modelagem(totais, ano_previsao)

    print("\nProjeção para", ano_previsao)
    print(
        f"Dados usados: anos completos de {int(dados_modelagem['ano'].min())} "
        f"a {int(dados_modelagem['ano'].max())}."
    )
    if anos_fora_modelagem:
        anos_texto = ", ".join(str(ano) for ano in anos_fora_modelagem)
        print(f"Nota: anos fora da modelagem ({anos_texto}).")

    print("\nResumo anual usado na modelagem:")
    print(dados_modelagem.to_string(index=False))
    print("\nProjeções e intervalos de confiança (95%):")
    print(projecoes.to_string(index=False))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Projeção de indicadores Embrapii.")
    parser.add_argument(
        "--input",
        type=pathlib.Path,
        default=ARQUIVO_ENTRADA_PADRAO,
        help="Arquivo Excel de entrada com a aba 'dados'.",
    )
    parser.add_argument(
        "--output",
        type=pathlib.Path,
        default=ARQUIVO_RESULTADOS_PADRAO,
        help="Arquivo CSV de saída com as projeções.",
    )
    parser.add_argument(
        "--ano",
        type=int,
        default=2027,
        help="Ano de previsão.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    df = carregar_dados(args.input)
    totais = obter_totais_anuais(df)
    projecoes, modelos = calcular_projecoes(totais, ano_previsao=args.ano)

    salvar_resultados(projecoes, args.output)
    imagem_saida = args.output.with_name("questao03_graficos.png")
    salvar_graficos(totais, projecoes, modelos, imagem_saida)

    relatorio_html = args.output.with_name("questao03_relatorio.html")
    salvar_relatorio_html(
        totais, projecoes, imagem_saida, relatorio_html, ano_previsao=args.ano
    )

    relatorio_pdf = args.output.with_name("questao03_relatorio.pdf")
    salvar_relatorio_pdf(
        totais, projecoes, imagem_saida, relatorio_pdf, ano_previsao=args.ano
    )

    imprimir_relatorio(totais, projecoes, args.ano)
    print(f"\nResultados salvos em: {args.output}")
    print(f"Gráfico salvo em: {imagem_saida}")
    print(f"Relatório HTML salvo em: {relatorio_html}")
    print(f"Relatório PDF salvo em: {relatorio_pdf}")


if __name__ == "__main__":
    main()
