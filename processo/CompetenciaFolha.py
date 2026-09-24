import pandas as pd
import numpy as np
from pathlib import Path
import re


# ============================================================
# DICIONÁRIO DE HASHES POR MÊS
# ============================================================
HASH_MES = {
    1: "4B3C33DF5930A1E983178C70E7E3E235",
    2: "389CC01684B3EAEFA8CC9835EB21A267",
    3: "1569B5E5E8612C9404C7F0BF3B68DDC5",
    4: "F2888D3B532BDD997FCFF044823E9A23",
    5: "9199844E732087CD5BB4688351757769",
    6: "4DC46816194F1FA38C5689F07862A233",
    7: "53CD2BCC5F2DE3F2E2BEE0E66392473A",
    8: "9D762831BDEDD4D8131BD7B5344BF7C5",
    9: "CFEFB84C9227D1A147F8D5D49CD7064A",
    10: "1673BBE748C6DBCD5338370C00139885",
    11: "1DF561B6D0F5DC85BF3C5E5FF9EB7348",
    12: "626ABB8C93CF362A52116AD0D7CF4C8F"
}

def processar_faltas(rubrica_falta=5040):
    # ============================================================
    # LOCALIZAÇÃO AUTOMÁTICA DO PROJETO
    # ============================================================

    # Competencia.py
    PASTA_PROCESSOS = Path(__file__).resolve().parent

    # 11Ghost
    PASTA_PRINCIPAL = PASTA_PROCESSOS.parent

    # 11Ghost/base
    PASTA_BASE = PASTA_PRINCIPAL / "base"

    # 11Ghost/base/TodasasFolhas
    PASTA_FOLHAS = PASTA_BASE / "TodasasFolhas"

    # 11Ghost/base/Vigencia
    PASTA_VIGENCIA = PASTA_BASE / "Vigencia"

    # Resultado será salvo em processo
    ARQUIVO_SAIDA = PASTA_BASE / "Resultado_Faltas.xlsx"

    # ============================================================
    # MOSTRAR CAMINHOS
    # ============================================================

    print("=" * 70)
    print("CAMINHOS IDENTIFICADOS")
    print("=" * 70)

    print(f"Projeto:    {PASTA_PRINCIPAL}")
    print(f"Base:       {PASTA_BASE}")
    print(f"Folhas:     {PASTA_FOLHAS}")
    print(f"Vigência:   {PASTA_VIGENCIA}")
    print(f"Saída:      {ARQUIVO_SAIDA}")
    print(f"Rubrica:    {rubrica_falta}")

    print("=" * 70)

    # ============================================================
    # VERIFICAR PASTAS
    # ============================================================

    if not PASTA_BASE.exists():
        raise FileNotFoundError(
            f"A pasta base não existe:\n{PASTA_BASE}"
        )

    if not PASTA_FOLHAS.exists():
        raise FileNotFoundError(
            f"A pasta TodasasFolhas não existe:\n{PASTA_FOLHAS}"
        )

    if not PASTA_VIGENCIA.exists():
        raise FileNotFoundError(
            f"A pasta Vigencia não existe:\n{PASTA_VIGENCIA}"
        )

    # ============================================================
    # LOCALIZAR PLANILHA PRINCIPAL
    # ============================================================

    extensoes = {
        ".xlsx",
        ".xls",
        ".xlsm",
        ".xlsb",
        ".csv"
    }

    arquivos_base = [
        arquivo
        for arquivo in PASTA_BASE.iterdir()
        if arquivo.is_file()
        and arquivo.suffix.lower() in extensoes
    ]

    print("\nPlanilhas encontradas diretamente em BASE:")

    for arquivo in arquivos_base:
        print(f"  - {arquivo.name}")

    if not arquivos_base:
        raise FileNotFoundError(
            f"Nenhuma planilha encontrada diretamente em:\n{PASTA_BASE}"
        )

    if len(arquivos_base) > 1:
        raise RuntimeError(
            "Foram encontradas várias planilhas diretamente na pasta BASE.\n"
            "Deixe somente uma planilha principal dentro de BASE."
        )

    ARQUIVO_BASE = arquivos_base[0]

    print(f"\nPlanilha principal:")
    print(f"  {ARQUIVO_BASE}")

    # ============================================================
    # LOCALIZAR TODAS AS FOLHAS
    # ============================================================

    arquivos_folhas = []

    for extensao in [
        "*.csv",
        "*.xlsx",
        "*.xls",
        "*.xlsm",
        "*.xlsb"
    ]:
        arquivos_folhas.extend(
            PASTA_FOLHAS.glob(extensao)
        )

    print(
        f"\nFolhas encontradas: {len(arquivos_folhas)}"
    )

    for arquivo in arquivos_folhas:
        print(f"  - {arquivo.name}")

    # ============================================================
    # LOCALIZAR TODAS AS VIGÊNCIAS
    # ============================================================

    arquivos_vigencia = []

    for extensao in [
        "*.csv",
        "*.xlsx",
        "*.xls",
        "*.xlsm",
        "*.xlsb"
    ]:
        arquivos_vigencia.extend(
            PASTA_VIGENCIA.glob(extensao)
        )

    print(
        f"\nVigências encontradas: {len(arquivos_vigencia)}"
    )

    for arquivo in arquivos_vigencia:
        print(f"  - {arquivo.name}")

    # ============================================================
    # VALIDAR FOLHAS
    # ============================================================

    if not arquivos_folhas:
        raise ValueError(
            f"Nenhuma folha encontrada em:\n{PASTA_FOLHAS}"
        )

    # ============================================================
    # 1. LER A BASE DE FALTAS
    # ============================================================

    print("\nLendo a base de faltas...")

    if ARQUIVO_BASE.suffix.lower() == ".csv":

        try:
            df_base = pd.read_csv(
                ARQUIVO_BASE,
                sep=";",
                encoding="utf-8-sig"
            )
        except UnicodeDecodeError:
            df_base = pd.read_csv(
                ARQUIVO_BASE,
                sep=";",
                encoding="cp1252"
            )

    elif ARQUIVO_BASE.suffix.lower() == ".xlsb":

        df_base = pd.read_excel(
            ARQUIVO_BASE,
            engine="pyxlsb"
        )

    else:

        df_base = pd.read_excel(
            ARQUIVO_BASE
        )

    print(
        f"Base carregada: {ARQUIVO_BASE.name}"
    )

    print(
        f"Registros encontrados: {len(df_base):,}"
    )


    # ============================================================
    # 2. CONVERTER AS DATAS
    # ============================================================

    df_base["DataInicial"] = pd.to_datetime(
        df_base["DataInicial"],
        errors="coerce",
        dayfirst=True
    )

    df_base["DataFinal"] = pd.to_datetime(
        df_base["DataFinal"],
        errors="coerce",
        dayfirst=True
    )


    # ============================================================
    # 2.1 CRIAR COLUNAS DE HASH, ANO E COMBINADO (BASE)
    # ============================================================

    # Verifica se as datas inicial e final são diferentes
    mask_diferentes = df_base["DataInicial"] != df_base["DataFinal"]

    # Hash do Mês
    df_base["HASH_DATA"] = np.where(
        mask_diferentes,
        "Aviso: Data Inicial e Data Final são diferentes",
        df_base["DataInicial"].dt.month.map(HASH_MES)
    )

    # Ano
    df_base["ANO_DATA"] = df_base["DataInicial"].dt.year

    # Ano + Hash
    df_base["ANO_E_HASH"] = np.where(
        mask_diferentes,
        "Aviso: Datas diferentes",
        df_base["ANO_DATA"].astype(str).str.replace(".0", "", regex=False) + "_" + df_base["HASH_DATA"]
    )


    # ============================================================
    # 3. CRIAR O MÊS DA FALTA
    # ============================================================

    df_base["MES_FALTA"] = (
        df_base["DataInicial"]
        .dt.to_period("M")
    )


    # ============================================================
    # 4. PADRONIZAR MATRÍCULA E VÍNCULO
    # ============================================================

    df_base["Func"] = (
        pd.to_numeric(
            df_base["Func"],
            errors="coerce"
        ).astype("Int64")
    )

    df_base["Vinc"] = (
        pd.to_numeric(
            df_base["Vinc"],
            errors="coerce"
        ).astype("Int64")
    )


    # ============================================================
    # 5. LISTAR TODAS AS FOLHAS
    # ============================================================

    arquivos = []

    for extensao in [
        "*.csv",
        "*.xlsx",
        "*.xls",
        "*.xlsb"
    ]:
        arquivos.extend(
            PASTA_FOLHAS.glob(extensao)
        )

    print(
        f"Foram encontrados {len(arquivos)} arquivos."
    )

    if not arquivos:
        raise ValueError(
            "Nenhuma planilha foi encontrada na pasta informada."
        )


    # ============================================================
    # 6. LER AS FOLHAS DE PAGAMENTO
    # ============================================================

    lista_folhas = []

    for i, arquivo in enumerate(
        arquivos,
        start=1
    ):

        print(
            f"Lendo arquivo {i}/{len(arquivos)}: "
            f"{arquivo.name}"
        )

        try:

            if arquivo.suffix.lower() == ".csv":
                try:
                    df = pd.read_csv(
                        arquivo,
                        sep=";",
                        encoding="utf-8-sig",
                        low_memory=False
                    )
                except UnicodeDecodeError:
                    df = pd.read_csv(
                        arquivo,
                        sep=";",
                        encoding="cp1252",
                        low_memory=False
                    )
            elif arquivo.suffix.lower() == ".xlsb":
                df = pd.read_excel(
                    arquivo,
                    engine="pyxlsb"
                )
            else:
                df = pd.read_excel(
                    arquivo
                )


            df.columns = (
                df.columns
                .astype(str)
                .str.strip()
            )


            colunas_necessarias = [
                "NUMFUNC",
                "NUMVINC",
                "MES_ANO_FOLHA",
                "NUMRUBR",
                "COMPETENCIA"
            ]

            faltando = [
                coluna
                for coluna in colunas_necessarias
                if coluna not in df.columns
            ]

            if faltando:
                print(
                    f"  -> Ignorado. "
                    f"Colunas ausentes: {faltando}"
                )
                continue


            # FILTRAR SOMENTE RUBRICA 5040
            df["NUMRUBR"] = pd.to_numeric(
                df["NUMRUBR"],
                errors="coerce"
            )

            df = df[
                df["NUMRUBR"] == rubrica_falta
            ].copy()

            if df.empty:
                print(
                    "  -> Nenhum registro com rubrica 5040."
                )
                continue


            # PADRONIZAR FUNC E VINC
            df["NUMFUNC"] = (
                pd.to_numeric(
                    df["NUMFUNC"],
                    errors="coerce"
                ).astype("Int64")
            )

            df["NUMVINC"] = (
                pd.to_numeric(
                    df["NUMVINC"],
                    errors="coerce"
                ).astype("Int64")
            )


            # CRIAR MÊS DA COMPETÊNCIA (YYYY-MM)
            df["MES_COMPETENCIA"] = (
                pd.to_datetime(
                    df["COMPETENCIA"],
                    errors="coerce",
                    dayfirst=True
                )
                .dt.to_period("M")
            )

            mask_invalida = df["MES_COMPETENCIA"].isna()

            if mask_invalida.any():
                competencia_texto = (
                    df.loc[mask_invalida, "COMPETENCIA"]
                    .astype(str)
                    .str.strip()
                )

                datas_yyyymm = (
                    pd.to_datetime(
                        competencia_texto,
                        format="%Y%m",
                        errors="coerce"
                    )
                    .dt.to_period("M")
                )

                df.loc[mask_invalida, "MES_COMPETENCIA"] = datas_yyyymm


            # ====================================================
            # GERAR HASH, ANO E ANO_HASH DA COMPETÊNCIA
            # ====================================================
            df["HASH_COMPETENCIA"] = df["MES_COMPETENCIA"].dt.month.map(HASH_MES)
            df["ANO_COMPETENCIA"] = df["MES_COMPETENCIA"].dt.year
            df["ANO_E_HASH_COMPETENCIA"] = (
                df["ANO_COMPETENCIA"].astype(str).str.replace(".0", "", regex=False)
                + "_"
                + df["HASH_COMPETENCIA"].astype(str)
            )


            # MANTER SOMENTE AS COLUNAS NECESSÁRIAS
            df = df[
                [
                    "NUMFUNC",
                    "NUMVINC",
                    "MES_ANO_FOLHA",
                    "COMPETENCIA",
                    "HASH_COMPETENCIA",
                    "ANO_COMPETENCIA",
                    "ANO_E_HASH_COMPETENCIA",
                    "MES_COMPETENCIA"
                ]
            ]

            df["ARQUIVO_ORIGEM"] = arquivo.name
            lista_folhas.append(df)

        except Exception as erro:
            print(
                f"  -> ERRO ao ler "
                f"{arquivo.name}: {erro}"
            )


    # ============================================================
    # 7. JUNTAR TODAS AS FOLHAS
    # ============================================================

    if not lista_folhas:
        raise ValueError(
            "Nenhuma folha válida foi encontrada."
        )

    df_folhas = pd.concat(
        lista_folhas,
        ignore_index=True
    )

    print(
        f"\nTotal de registros com "
        f"rubrica 5040: {len(df_folhas):,}"
    )


    # ============================================================
    # 8. REMOVER DUPLICIDADES
    # ============================================================

    df_folhas = df_folhas.drop_duplicates(
        subset=[
            "NUMFUNC",
            "NUMVINC",
            "MES_COMPETENCIA"
        ]
    )


    # ============================================================
    # 9. CRIAR CHAVE PARA CRUZAMENTO (MATRÍCULA + VÍNCULO)
    # ============================================================

    df_base["CHAVE"] = (
        df_base["Func"].astype("string")
        + "_"
        + df_base["Vinc"].astype("string")
    )

    df_folhas["CHAVE"] = (
        df_folhas["NUMFUNC"].astype("string")
        + "_"
        + df_folhas["NUMVINC"].astype("string")
    )


    # ============================================================
    # 10. CRUZAR AS FALTAS COM AS FOLHAS
    # ============================================================

    print(
        "\nProcurando os descontos..."
    )

    descontos_por_servidor = (
        df_folhas
        .dropna(subset=["MES_COMPETENCIA"])
        .groupby("CHAVE")
        .apply(
            lambda grupo: list(
                zip(
                    grupo["MES_COMPETENCIA"],
                    grupo["ARQUIVO_ORIGEM"],
                    grupo["MES_ANO_FOLHA"]
                )
            )
        )
        .to_dict()
    )


    # ============================================================
    # 11. PROCURAR A COMPETÊNCIA EXATA DA FALTA (MÊS E ANO)
    # ============================================================

    def encontrar_descontos(row):

        chave = row["CHAVE"]
        mes_falta = row["MES_FALTA"]

        if pd.isna(mes_falta):
            return pd.Series(["", "", ""])

        registros = descontos_por_servidor.get(chave, [])

        # REGRA: A COMPETENCIA tem que ser EXATAMENTE IGUAL (==) ao mês/ano da falta
        registros_validos = [
            registro
            for registro in registros
            if registro[0] == mes_falta  # Exatidão de mês e ano
        ]

        if not registros_validos:
            return pd.Series(["", "", ""])

        meses_texto = ", ".join(
            reg[0].strftime("%m/%Y") for reg in registros_validos
        )

        arquivos_texto = "; ".join(
            set(reg[1] for reg in registros_validos)
        )

        folhas_execucao = "; ".join(
            set(str(reg[2]) for reg in registros_validos)
        )

        return pd.Series(
            [
                meses_texto,
                arquivos_texto,
                folhas_execucao
            ]
        )


    # ============================================================
    # 12. APLICAR A PROCURA NA BASE
    # ============================================================

    df_base[
        [
            "COMPETENCIA_DESCONTO_ENCONTRADA",
            "ARQUIVOS_DESCONTO",
            "FOLHAS_EXECUCAO"
        ]
    ] = df_base.apply(
        encontrar_descontos,
        axis=1
    )


    # ============================================================
    # 13. INFORMAR SE FOI ENCONTRADO
    # ============================================================

    df_base["DESCONTO_ENCONTRADO"] = (
        df_base["COMPETENCIA_DESCONTO_ENCONTRADA"]
        .ne("")
        .map(
            {
                True: "SIM",
                False: "NÃO"
            }
        )
    )


    # ============================================================
    # 14. FORMATAR MES_FALTA PARA EXIBIÇÃO
    # ============================================================

    df_base["MES_FALTA"] = (
        df_base["MES_FALTA"]
        .apply(
            lambda x:
            x.strftime("%m/%Y")
            if pd.notna(x)
            else ""
        )
    )


    # ============================================================
    # 15. REMOVER COLUNA AUXILIAR
    # ============================================================

    df_base = df_base.drop(
        columns=["CHAVE"]
    )


    # ============================================================
    # 16. SALVAR RESULTADO
    # ============================================================

    print(
        "\nSalvando resultado..."
    )

    df_base.to_excel(
        ARQUIVO_SAIDA,
        index=False
    )


    # ============================================================
    # 17. RESUMO
    # ============================================================

    total = len(df_base)

    encontrados = (
        df_base["DESCONTO_ENCONTRADO"] == "SIM"
    ).sum()

    nao_encontrados = (
        df_base["DESCONTO_ENCONTRADO"] == "NÃO"
    ).sum()


    print("\n==========================================")
    print("PROCESSO FINALIZADO")
    print("==========================================")
    print(f"Total de faltas analisadas: {total:,}")
    print(f"Desconto encontrado:        {encontrados:,}")
    print(f"Não encontrado:              {nao_encontrados:,}")
    print("\nArquivo salvo em:")
    print(ARQUIVO_SAIDA)
    print("==========================================")

if __name__ == "__main__":
    processar_faltas()