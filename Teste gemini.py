import os
import pandas as pd
import threading

# Função para processar arquivos Excel e salvar como CSV
def processar_arquivos_excel(caminho):
    total_arquivos = len(os.listdir(caminho))
    arquivos_processados = 0
    arquivos_com_erro = 0

    for arquivo in os.listdir(caminho):
        if arquivo.endswith('.xlsx'):
            try:
                caminho_completo = os.path.join(caminho, arquivo)
                df = pd.read_excel(caminho_completo)
                df = df.iloc[1:]

                # Salva backup do arquivo XLSX
                caminho_backup = os.path.join(caminho, 'BACKUP', arquivo)
                if os.path.exists(caminho_completo) and not os.path.exists(caminho_backup):
                    os.makedirs(os.path.dirname(caminho_backup), exist_ok=True)
                    os.rename(caminho_completo, caminho_backup)

                # Converte para CSV e salva na pasta escolhida
                pasta_csv = filedialog.askdirectory(title='Selecione a pasta para salvar os arquivos CSV')
                nome_csv = arquivo.replace('.xlsx', '.csv')
                caminho_csv = os.path.join(pasta_csv, nome_csv)
                df.to_csv(caminho_csv, index=False)

                arquivos_processados += 1

            except Exception as e:
                arquivos_com_erro += 1
                print(f'Erro ao processar arquivo {arquivo}: {e}')

    # Consolidação dos arquivos CSV
    consolidar_csv(pasta_csv)

# Função para consolidar arquivos CSV
def consolidar_csv(caminho):
    arquivos_csv = [arq for arq in os.listdir(caminho) if arq.endswith('.csv')]
    df_total = pd.concat([pd.read_csv(os.path.join(caminho, arq)) for arq in arquivos_csv])
    df_total.to_csv(os.path.join(caminho, 'consolidado.csv'), index=False)

# Função principal
def main():
    # Seleciona a pasta com os arquivos Excel
    caminho = filedialog.askdirectory(title='Selecione a pasta com os arquivos Excel')

    # Inicia o processamento dos arquivos
    threading.Thread(target=processar_arquivos_excel, args=(caminho,)).start()

    # Mostra a mensagem de conclusão
    messagebox.showinfo('Conclusão', 'Processamento concluído com sucesso!')

if __name__ == '__main__':
    main()
