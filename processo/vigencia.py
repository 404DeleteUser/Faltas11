import pandas as pd
from pathlib import Path

def verificar_vigencias():
    print("Iniciando a verificação de Vigência (Matrícula e Vínculo)...\n")

    # ============================================================
    # 1. DESCOBERTA DINÂMICA DE CAMINHOS
    # ============================================================
    # Pega o diretório atual onde o script está rodando (pasta "processo")
    diretorio_processo = Path(__file__).resolve().parent
    
    # Sobe um nível para chegar na raiz
    diretorio_raiz = diretorio_processo.parent
    
    # Caminho da pasta de vigência
    pasta_vigencia = diretorio_raiz / "base" / "Vigencia"

    if not pasta_vigencia.exists():
        raise FileNotFoundError(f"A pasta de vigência não foi encontrada em: {pasta_vigencia}")

    # ============================================================
    # 2. IDENTIFICAR OS DOIS ARQUIVOS NA RAIZ
    # ============================================================
    # Procura por arquivos Excel apenas soltos na raiz (ignora pastas)
    arquivos_raiz = [arq for arq in diretorio_raiz.iterdir() if arq.is_file() and arq.suffix.lower() in [".xlsx", ".xls"]]
    
    if not arquivos_raiz:
        raise FileNotFoundError("Nenhum arquivo Excel foi encontrado na raiz do diretório.")
    
    print(f"Foram encontrados {len(arquivos_raiz)} arquivo(s) na raiz para verificação.")

    # ============================================================
    # 3. LER AS PLANILHAS DE VIGÊNCIA (BASE DE CONSULTA)
    # ============================================================
    arquivos_ref = []
    for ext in ["*.xlsx", "*.xls", "*.csv"]:
        arquivos_ref.extend(pasta_vigencia.glob(ext))
        
    if not arquivos_ref:
        raise ValueError("Nenhum arquivo de referência encontrado na pasta 'base/Vigencia'.")

    lista_df_ref = []
    for arq in arquivos_ref:
        print(f"Lendo base de vigência: {arq.name}")
        df_temp = pd.read_excel(arq) if arq.suffix.lower() != ".csv" else pd.read_csv(arq, sep=";")
        df_temp.columns = df_temp.columns.str.strip()
        lista_df_ref.append(df_temp)

    df_ref = pd.concat(lista_df_ref, ignore_index=True)

    # Validar colunas da referência
    for col in ["Func", "Vinc", "Vigente"]:
        if col not in df_ref.columns:
            raise ValueError(f"A coluna obrigatória '{col}' não foi encontrada nos arquivos de Vigência.")

    # ============================================================
    # 4. CRIAR DICIONÁRIO DE MATRÍCULAS E VÍNCULOS ATIVOS
    # ============================================================
    # Padronizar para evitar erros de tipagem (.0 no final, etc)
    df_ref["Func_Str"] = pd.to_numeric(df_ref["Func"], errors="coerce").astype("Int64").astype(str)
    df_ref["Vinc_Str"] = pd.to_numeric(df_ref["Vinc"], errors="coerce").astype("Int64").astype(str)
    
    # Filtra apenas quem está vigente. 
    # Obs: Ajuste as palavras da lista abaixo conforme estiver escrito na sua planilha (ex: "SIM", "ATIVO", "VIGENTE")
    mask_vigente = df_ref["Vigente"].astype(str).str.strip().str.upper().isin(["SIM", "S", "ATIVO", "TRUE", "VIGENTE", "1", "VIG"])
    df_ativos = df_ref[mask_vigente]

    # Cria um dicionário onde a chave é a Matrícula (Func) e o valor é uma lista de Vínculos (Vinc) ativos
    dict_ativos = (
        df_ativos.groupby("Func_Str")["Vinc_Str"]
        .apply(list)
        .to_dict()
    )

    # ============================================================
    # 5. PROCESSAR OS ARQUIVOS DA RAIZ
    # ============================================================
    def checar_vigencia_row(row):
        # Transforma o Func e Vinc do arquivo alvo em string limpa
        func = str(row["Func"]).replace(".0", "").strip() if pd.notna(row.get("Func")) else ""
        vinc = str(row["Vinc"]).replace(".0", "").strip() if pd.notna(row.get("Vinc")) else ""
        
        vinc_final = "não encontrado"
        
        if func in dict_ativos:
            vincs_ativos_da_matricula = dict_ativos[func]
            
            # Prioridade 1: O vínculo que já está na planilha alvo é válido?
            if vinc in vincs_ativos_da_matricula:
                vinc_final = vinc
            else:
                # Prioridade 2: O vínculo atual não é válido. Pega outro vínculo ativo dessa matrícula.
                vinc_final = vincs_ativos_da_matricula[0]
                
        return pd.Series([func, vinc_final])


    for arquivo_alvo in arquivos_raiz:
        print(f"\nProcessando arquivo da raiz: {arquivo_alvo.name}")
        
        df_alvo = pd.read_excel(arquivo_alvo)
        df_alvo.columns = df_alvo.columns.str.strip()
        
        if "Func" not in df_alvo.columns or "Vinc" not in df_alvo.columns:
            print(f"  -> Ignorado: Colunas 'Func' ou 'Vinc' não encontradas em {arquivo_alvo.name}")
            continue

        # Aplica a regra linha a linha e cria as duas novas colunas
        df_alvo[["VigenteFunc", "VigenteVinc"]] = df_alvo.apply(checar_vigencia_row, axis=1)
        
        # Salva o arquivo sobrescrevendo no mesmo local e com o mesmo nome
        df_alvo.to_excel(arquivo_alvo, index=False)
        print(f"  -> Salvo com sucesso. Atualizado com as colunas VigenteFunc e VigenteVinc.")

    print("\n==========================================")
    print("VERIFICAÇÃO DE VIGÊNCIA CONCLUÍDA")
    print("==========================================")

if __name__ == "__main__":
    verificar_vigencias()