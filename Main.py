
import os
import warnings
from playwright_Itau import baixar_itau_base
from processar_itau_base import processar_itau_base
from atualizar_base_principal import atualizar_base_principal

warnings.simplefilter("ignore")
warnings.filterwarnings("ignore")

def main():

    pasta = r"C:\Users\uih34193\OneDrive - Aumovio SE\Análise Rastreadores Itaú"

    arquivo_vdo = os.path.join(pasta, "Itau_ultimos_registros.xlsx")
    arquivo_goal = os.path.join(pasta, "Itau_instalacoes_atendidas.xlsx")
    arquivo_principal = os.path.join(pasta, "Relatório_Placas Sem Sinal_Status Itau.xlsx")

    print("1) Iniciando downloads (VDO e GOAL)...")
    baixar_itau_base(arquivo_vdo, arquivo_goal)

    print("2) Processando arquivo baixado...")
    df_limpo = processar_itau_base(arquivo_vdo)

    print("3) Atualizando base principal...")
    atualizar_base_principal(df_limpo, arquivo_principal, arquivo_goal)
    
    if os.path.exists(arquivo_vdo):
            os.remove(arquivo_vdo)
            print(f"Arquivo temporário VDO removido.")
   
    print("\nProcesso finalizado com sucesso!")

if __name__ == "__main__":
    main()

