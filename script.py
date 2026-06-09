import pandas as pd
import numpy as np
import os
import datetime
import unicodedata
from openpyxl.styles import numbers


def remove_accents(text):
    """Remove accents from text"""
    nfd = unicodedata.normalize("NFD", text)
    return "".join(c for c in nfd if unicodedata.category(c) != "Mn")


ARQUIVO_ENTRADA = "Embrapii_seleção_analista_2026_questao01_Eventos.xlsx"
ARQUIVO_SAIDA = "Eventos_Estruturados.xlsx"

xls = pd.ExcelFile(ARQUIVO_ENTRADA)

abas = xls.sheet_names

dfs = []

for aba in abas:
    df_temp = pd.read_excel(ARQUIVO_ENTRADA, sheet_name=aba)

    df_temp["ABA_ORIGEM"] = aba

    dfs.append(df_temp)

df = pd.concat(dfs, ignore_index=True)

print(f"Total de registros: {len(df)}")

df = df.dropna(how="all")
df = df.dropna(axis=1, how="all")

for col in df.select_dtypes(include=["object", "string"]):

    df[col] = df[col].astype(str).str.strip().replace("nan", np.nan)

df.insert(0, "ID_EVENTO", range(1, len(df) + 1))


df.columns = df.columns.map(
    lambda x: remove_accents(str(x)).strip().upper().replace(" ", "_").replace("/", "_")
)

colunas = list(df.columns)

print("\nColunas encontradas:")
print(colunas)

for col in df.select_dtypes(include=["object", "string"]):

    if "UF" in col:
        df[col] = df[col].str.upper()

    elif "CIDADE" in col:
        df[col] = df[col].str.title()

possiveis_datas = [c for c in df.columns if "DATA" in c]

for col in possiveis_datas:

    try:
        df[col] = pd.to_datetime(df[col], errors="coerce")
    except Exception:
        pass

data_inicio = None

for c in df.columns:
    if "INICIO" in c:
        data_inicio = c
        break

data_fim = None

for c in df.columns:
    if "FIM" in c or "FINAL" in c:
        data_fim = c
        break

if data_inicio and data_fim:

    df["DURACAO_DIAS"] = (df[data_fim] - df[data_inicio]).dt.days + 1

# --- Indicadores Adicionais ---
# DURACAO_DIAS - se DATA_FINAL vazia → 1 dia
if data_inicio:
    if "DURACAO_DIAS" not in df.columns:
        df["DURACAO_DIAS"] = 1
    else:
        df["DURACAO_DIAS"] = df["DURACAO_DIAS"].fillna(1)

# QTD_TEMAS - contar temas (separados por vírgula)
temas_col = None
for col in df.columns:
    if "TEMAS" in col:
        temas_col = col
        df["QTD_TEMAS"] = (
            df[col]
            .fillna("")
            .str.split(",")
            .apply(lambda x: len([t for t in x if t.strip()]))
        )
        break
else:
    df["QTD_TEMAS"] = 0

# TEMAS_LISTA - normalizar separadores (/ e ,) para ;
# Permite Power BI fazer split/unpivot de cada tema
if temas_col is not None:
    df["TEMAS_LISTA"] = (
        df[temas_col]
        .fillna("")
        .str.replace("/", ";", regex=False)
        .str.replace(",", ";", regex=False)
        .str.strip()
    )
else:
    df["TEMAS_LISTA"] = ""

# QTD_PROGRAMAS_INICIATIVAS - contar programas/iniciativas
programas_col = None
for col in df.columns:
    if "PROGRAMAS" in col or "INICIATIVAS" in col:
        programas_col = col
        df["QTD_PROGRAMAS_INICIATIVAS"] = (
            df[col]
            .fillna("")
            .str.split(",|;")
            .apply(lambda x: len([p for p in x if p.strip()]))
        )
        break
else:
    df["QTD_PROGRAMAS_INICIATIVAS"] = 0

# PROGRAMAS_LISTA - normalizar separadores (/ e ,) para ;
# Permite Power BI fazer split/unpivot de cada programa
if programas_col is not None:
    df["PROGRAMAS_LISTA"] = (
        df[programas_col]
        .fillna("")
        .str.replace("/", ";", regex=False)
        .str.replace(",", ";", regex=False)
        .str.strip()
    )
else:
    df["PROGRAMAS_LISTA"] = ""

# ORGANIZADOR_LISTA - normalizar separadores (/, , e " e ") para ;
# Permite Power BI fazer split/unpivot de cada organizador
organizador_col = None
for col in df.columns:
    if "ORGANIZADOR" in col:
        organizador_col = col
        df["ORGANIZADOR_LISTA"] = (
            df[col]
            .fillna("")
            .str.replace(" e ", ";", regex=False)
            .str.replace("/", ";", regex=False)
            .str.replace(",", ";", regex=False)
            .str.strip()
        )
        break
else:
    df["ORGANIZADOR_LISTA"] = ""

# EMBRAPII_DAY_NORMALIZADO - converter Sim/Não para 1/0
for col in df.columns:
    if "EMBRAPII" in col and "DAY" in col:
        df["EMBRAPII_DAY_NORMALIZADO"] = (
            df[col]
            .astype(str)
            .str.lower()
            .map({"sim": 1, "s": 1, "yes": 1, "1": 1, "true": 1})
            .fillna(0)
            .astype(int)
        )
        break
else:
    df["EMBRAPII_DAY_NORMALIZADO"] = 0

for c in df.columns:

    if "FORMATO" in c:

        mapa = {
            "online": "Online",
            "virtual": "Online",
            "presencial": "Presencial",
            "hibrido": "Híbrido",
            "híbrido": "Híbrido",
        }

        df[c] = df[c].astype(str).str.lower().map(mapa).fillna(df[c])


duplicados = pd.DataFrame()

possivel_nome = None

for c in df.columns:

    if "EVENTO" in c or "NOME" in c:

        possivel_nome = c
        break

if possivel_nome and data_inicio:

    duplicados = df[df.duplicated(subset=[possivel_nome, data_inicio], keep=False)]

relatorio_qualidade = pd.DataFrame(
    {
        "CAMPO": df.columns,
        "NULOS": df.isnull().sum().values,
        "PERCENTUAL_NULOS": np.round(df.isnull().mean().values * 100, 2),
    }
)


dim_organizador = pd.DataFrame()
org_col = None

for c in df.columns:

    if "ORGANIZADOR" in c:

        org_col = c
        dim_organizador = df[[c]].dropna().drop_duplicates().reset_index(drop=True)

        dim_organizador.insert(0, "ID_ORGANIZADOR", range(1, len(dim_organizador) + 1))

        break

dim_responsavel = pd.DataFrame()
resp_col = None

for c in df.columns:

    if "RESPONSAVEL" in c:

        resp_col = c
        dim_responsavel = df[[c]].dropna().drop_duplicates().reset_index(drop=True)

        dim_responsavel.insert(0, "ID_RESPONSAVEL", range(1, len(dim_responsavel) + 1))

        break

# --- Visão geral / KPIs ---
total_registros = len(df)
total_colunas = len(df.columns)
total_organizadores = len(dim_organizador) if not dim_organizador.empty else 0
registros_duplicados = 0
if not duplicados.empty:
    # conta registros únicos duplicados (por nome+data)
    try:
        registros_duplicados = duplicados.drop_duplicates(
            subset=[possivel_nome, data_inicio]
        ).shape[0]
    except Exception:
        registros_duplicados = duplicados.shape[0]

percentual_medio_nulos = (
    np.round(relatorio_qualidade["PERCENTUAL_NULOS"].mean(), 2)
    if not relatorio_qualidade.empty
    else 0
)

# --- Preparar colunas de suporte para Power BI ---
# Mapear IDs de organizador / responsavel
if org_col is not None and not dim_organizador.empty:
    mapping_org = dict(
        zip(dim_organizador.iloc[:, 1], dim_organizador["ID_ORGANIZADOR"])
    )
    df["ID_ORGANIZADOR"] = df[org_col].map(mapping_org).astype("Int64")
else:
    df["ID_ORGANIZADOR"] = pd.NA

if resp_col is not None and not dim_responsavel.empty:
    mapping_resp = dict(
        zip(dim_responsavel.iloc[:, 1], dim_responsavel["ID_RESPONSAVEL"])
    )
    df["ID_RESPONSAVEL"] = df[resp_col].map(mapping_resp).astype("Int64")
else:
    df["ID_RESPONSAVEL"] = pd.NA

# Coluna de formato normalizado (se existir)
format_cols = [c for c in df.columns if "FORMATO" in c]
if format_cols:
    df["FORMATO_NORMALIZADO"] = df[format_cols[0]]
else:
    df["FORMATO_NORMALIZADO"] = pd.NA

# Se existir o arquivo de saída, tenta removê-lo antes de escrever
output_path = ARQUIVO_SAIDA
if os.path.exists(output_path):
    try:
        os.remove(output_path)
    except PermissionError:
        now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        alt = output_path.replace(".xlsx", f"_{now}.xlsx")
        msg = f"Arquivo {output_path} está em uso. Gravando em {alt}."
        print(msg)
        output_path = alt

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    # Escreve apenas a aba EVENTS para consumo no Power BI
    df.to_excel(writer, sheet_name="EVENTOS", index=False)

    # Formatar colunas de data
    workbook = writer.book
    worksheet = writer.sheets["EVENTOS"]

    # Encontrar colunas de data e aplicar formato de data
    date_fmt = numbers.FORMAT_DATE_XLSX14  # dd/mm/yyyy
    for col_idx, col_name in enumerate(df.columns, 1):
        if col_name in possiveis_datas:
            # Aplicar formato de data a toda coluna (exceto header)
            for row_idx in range(2, len(df) + 2):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.number_format = date_fmt

print("\nProcessamento concluído!")
print(f"Arquivo gerado: {output_path}")
