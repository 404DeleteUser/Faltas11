import sys
import time
from pathlib import Path

# Importação dos scripts na pasta 'processo'
try:
    from CompetenciaFolha import processar_faltas
    from FaltasSeap import realizar_cruzamento_seap
    from vigencia import verificar_vigencias
    from formatação import formatar_faltas_folha, formatar_faltas_seap  # Arquivo renomeado para formatacao.py
except ImportError as e:
    print(f"Erro ao importar os módulos: {e}")
    print("Verifique se todos os arquivos (.py) estão dentro da mesma pasta 'processo'.")
    sys.exit(1)


def executar_pipeline():
    inicio_total = time.time()
    
    print("==================================================")
    print("      INICIANDO PROCESSAMENTO COMPLETO            ")
    print("==================================================\n")

    # ----------------------------------------------------
    # ETAPA 1: COMPETÊNCIA
    # ----------------------------------------------------
    print(">>> [ETAPA 1/4] Processando Competência Folha...")
    try:
        processar_faltas()
    except Exception as e:
        print(f"❌ Erro na Etapa 1 (Competência): {e}")
        return

    # ----------------------------------------------------
    # ETAPA 2: FALTAS SEAP
    # ----------------------------------------------------
    print("\n>>> [ETAPA 2/4] Cruzando dados com Faltas SEAP...")
    try:
        realizar_cruzamento_seap()
    except Exception as e:
        print(f"❌ Erro na Etapa 2 (Faltas SEAP): {e}")
        return

    # ----------------------------------------------------
    # ETAPA 3: VIGÊNCIA
    # ----------------------------------------------------
    print("\n>>> [ETAPA 3/4] Verificando Vigências...")
    try:
        verificar_vigencias()
    except Exception as e:
        print(f"❌ Erro na Etapa 3 (Vigência): {e}")
        return

    # ----------------------------------------------------
    # ETAPA 4: FORMATAÇÃO VISUAL
    # ----------------------------------------------------
    print("\n>>> [ETAPA 4/4] Aplicando Formatação nas Planilhas...")
    try:
        formatar_faltas_folha()
        formatar_faltas_seap()
    except Exception as e:
        print(f"❌ Erro na Etapa 4 (Formatação): {e}")
        return

    tempo_total = time.time() - inicio_total
    print("\n==================================================")
    print(f"  ALL PROCESSES FINISHED IN {tempo_total:.2f}s")
    print("==================================================")


if __name__ == "__main__":
    executar_pipeline()