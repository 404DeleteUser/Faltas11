import pandas as pd
from pathlib import Path

# ============================================================
# 1. DICIONÁRIO DE COLUNAS DO SEAP
# ============================================================
# Defina aqui quais colunas você quer trazer do arquivo do SEAP para a Base 11.
# Formato -> "NOME_DA_COLUNA_NO_SEAP": "NOME_QUE_FICARA_NA_BASE_FINAL"
COLUNAS_PARA_PUXAR = {
    "FREQ_DETALHES": "FaltasSeap",
    # Adicione ou altere as colunas conforme a necessidade da sua planilha SEAP
}


def realizar_cruzamento_seap():
    print("Iniciando cruzamento com as faltas do SEAP...\n")

    # ============================================================
    # 2. DESCOBERTA DINÂMICA DE CAMINHOS (SEM HARDCODE)
    # ============================================================
    # Pega o diretório absoluto onde este script está rodando (esperado: pasta "processo")
    diretorio_atual = Path(__file__).resolve().parent
    
    # Sobe um nível e entra na pasta "base"
    pasta_inicial = diretorio_atual.parent
    pasta_base = diretorio_atual.parent / "base"
    pasta_seap = pasta_base / "FaltasSeap"
    pasta_processo = diretorio_atual  # Pasta onde o script está localizado

    if not pasta_base.exists():
        raise FileNotFoundError(f"A pasta base não foi encontrada no caminho esperado: {pasta_base}")
    if not pasta_seap.exists():
        raise FileNotFoundError(f"A pasta FaltasSeap não foi encontrada dentro da base: {pasta_seap}")

    # ============================================================
    # 3. IDENTIFICAR O ARQUIVO DA BASE (11)
    # ============================================================
    # Procura arquivos Excel na raiz da pasta base (ignorando subpastas)
    arquivos_base = [arq for arq in pasta_base.glob("*.xlsx") if arq.is_file()]
    
    if not arquivos_base:
        raise FileNotFoundError("Nenhum arquivo Excel encontrado na raiz da pasta 'base'.")
    
    arquivo_base_11 = arquivos_base[0]  # Pega o primeiro arquivo Excel encontrado na base
    print(f"Planilha Base (11) identificada: {arquivo_base_11.name}")

    df_base = pd.read_excel(arquivo_base_11)
    df_base.columns = df_base.columns.str.strip()

    # Validação das colunas da Base
    colunas_base_req = ["Func", "Vinc", "DataInicial", "DataFinal"]
    for col in colunas_base_req:
        if col not in df_base.columns:
            raise ValueError(f"Coluna obrigatória '{col}' ausente na planilha base.")

    # ============================================================
    # 4. LER E CONSOLIDAR OS ARQUIVOS DO SEAP
    # ============================================================
    arquivos_seap = []
    for extensao in ["*.xlsx", "*.xls", "*.csv"]:
        arquivos_seap.extend(pasta_seap.glob(extensao))

    if not arquivos_seap:
        raise ValueError("Nenhum arquivo de falta encontrado na pasta 'FaltasSeap'.")

    lista_df_seap = []
    for arquivo in arquivos_seap:
        print(f"Lendo arquivo SEAP: {arquivo.name}")
        if arquivo.suffix.lower() == ".csv":
            df_temp = pd.read_csv(arquivo, sep=";", encoding="latin1", low_memory=False)
        else:
            df_temp = pd.read_excel(arquivo)
        
        df_temp.columns = df_temp.columns.str.strip()
        lista_df_seap.append(df_temp)

    df_seap = pd.concat(lista_df_seap, ignore_index=True)

    # Validação das colunas do SEAP
    colunas_seap_req = ["NUMFUNC", "NUMVINC", "DTINI", "DTFIM"]
    for col in colunas_seap_req:
        if col not in df_seap.columns:
            raise ValueError(f"Coluna obrigatória '{col}' ausente nas planilhas do SEAP.")

    # Verifica se as colunas extras solicitadas no dicionário existem no SEAP
    colunas_validas_extrair = []
    for col_origem, col_destino in COLUNAS_PARA_PUXAR.items():
        if col_origem in df_seap.columns:
            colunas_validas_extrair.append(col_origem)
        else:
            print(f"Aviso: Coluna '{col_origem}' definida no dicionário não existe no arquivo SEAP e será ignorada.")

    # ============================================================
    # 5. TRATAMENTO DOS DADOS PARA O CRUZAMENTO
    # ============================================================
    # Converter chaves numéricas para texto a fim de evitar erros de ".0" no cruzamento
    df_base["Func_Str"] = pd.to_numeric(df_base["Func"], errors="coerce").astype("Int64").astype(str)
    df_base["Vinc_Str"] = pd.to_numeric(df_base["Vinc"], errors="coerce").astype("Int64").astype(str)
    
    df_seap["NUMFUNC_Str"] = pd.to_numeric(df_seap["NUMFUNC"], errors="coerce").astype("Int64").astype(str)
    df_seap["NUMVINC_Str"] = pd.to_numeric(df_seap["NUMVINC"], errors="coerce").astype("Int64").astype(str)

    # Padronizar as Datas, removendo horas se existirem
    df_base["DataInicial_Fmt"] = pd.to_datetime(df_base["DataInicial"], errors="coerce", dayfirst=True).dt.normalize()
    df_base["DataFinal_Fmt"] = pd.to_datetime(df_base["DataFinal"], errors="coerce", dayfirst=True).dt.normalize()

    df_seap["DTINI_Fmt"] = pd.to_datetime(df_seap["DTINI"], errors="coerce", dayfirst=True).dt.normalize()
    df_seap["DTFIM_Fmt"] = pd.to_datetime(df_seap["DTFIM"], errors="coerce", dayfirst=True).dt.normalize()

    # Reduzir o dataframe do SEAP apenas para as colunas-chave e as colunas do dicionário
    colunas_necessarias_seap = ["NUMFUNC_Str", "NUMVINC_Str", "DTINI_Fmt", "DTFIM_Fmt"] + colunas_validas_extrair
    df_seap_reduzido = df_seap[colunas_necessarias_seap].drop_duplicates()

    # ============================================================
    # 6. EXECUTAR O CRUZAMENTO (LEFT JOIN)
    # ============================================================
    print("\nExecutando o cruzamento de dados...")
    
    # Fazemos um merge (procv) baseado nas 4 condições
    df_resultado = pd.merge(
        left=df_base,
        right=df_seap_reduzido,
        left_on=["Func_Str", "Vinc_Str", "DataInicial_Fmt", "DataFinal_Fmt"],
        right_on=["NUMFUNC_Str", "NUMVINC_Str", "DTINI_Fmt", "DTFIM_Fmt"],
        how="left" # Mantém todos os registros da base da 11
    )

    # Identificar se a falta foi achada no SEAP (baseado se DTINI_Fmt foi preenchida durante o cruzamento)
    df_resultado["ENCONTRADO_NO_SEAP"] = df_resultado["DTINI_Fmt"].notna().map({True: "SIM", False: "NÃO"})

    # Renomear as colunas puxadas do SEAP conforme o dicionário
    renomeacoes = {col: COLUNAS_PARA_PUXAR[col] for col in colunas_validas_extrair}
    df_resultado.rename(columns=renomeacoes, inplace=True)

    # Limpar colunas auxiliares usadas no cruzamento
    colunas_limpeza = ["Func_Str", "Vinc_Str", "NUMFUNC_Str", "NUMVINC_Str", "DataInicial_Fmt", "DataFinal_Fmt", "DTINI_Fmt", "DTFIM_Fmt"]
    df_resultado.drop(columns=colunas_limpeza, errors="ignore", inplace=True)

    # ============================================================
    # 7. SALVAR ARQUIVO FINAL
    # ============================================================
    arquivo_saida = pasta_inicial / "FaltasSEAP.xlsx"
    df_resultado.to_excel(arquivo_saida, index=False)
    
    print("\n==========================================")
    print("PROCESSO SEAP FINALIZADO")
    print("==========================================")
    print(f"Total de registros na Base 11: {len(df_resultado)}")
    print(f"Encontrados no SEAP:           {(df_resultado['ENCONTRADO_NO_SEAP'] == 'SIM').sum()}")
    print(f"Não encontrados no SEAP:        {(df_resultado['ENCONTRADO_NO_SEAP'] == 'NÃO').sum()}")
    print(f"Arquivo salvo em: {arquivo_saida}")
    print("==========================================\n")

if __name__ == "__main__":
    realizar_cruzamento_seap()