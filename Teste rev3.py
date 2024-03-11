import os
import pandas as pd
import threading
import queue
import time
from tkinter import Tk, Label, Button, messagebox, filedialog, ttk

# Variáveis globais para controle do progresso
total_arquivos = 0
arquivos_processados = 0
arquivos_com_erro = 0
inicio_processamento = 0

# Função para processar arquivos Excel e salvar como CSV
def processar_arquivos_excel(caminho, fila, progresso_var):
    global total_arquivos, arquivos_processados, arquivos_com_erro, inicio_processamento
    inicio_processamento = time.time()
    
    for pasta, subpastas, arquivos in os.walk(caminho):
        for arquivo in arquivos:
            if arquivo.endswith('.xlsx'):
                total_arquivos += 1
    
    for pasta, subpastas, arquivos in os.walk(caminho):
        for arquivo in arquivos:
            if arquivo.endswith('.xlsx'):
                try:
                    caminho_completo = os.path.join(pasta, arquivo)
                    df = pd.read_excel(caminho_completo)
                    df = df.iloc[1:]  # Remove a primeira linha
                    nome_csv = arquivo.replace('.xlsx', '.csv')
                    caminho_csv = os.path.join(pasta, nome_csv)
                    
                    # Verifica se o arquivo CSV já existe
                    if os.path.exists(caminho_csv):
                        resposta = messagebox.askyesnocancel('Arquivo Existente', f'O arquivo {nome_csv} já existe. Deseja substituir?')
                        if resposta:  # Substituir
                            df.to_csv(caminho_csv, index=False)
                        elif resposta is None:  # Cancelar
                            continue
                    else:
                        df.to_csv(caminho_csv, index=False)
                    
                    # Salva backup do arquivo XLSX
                    caminho_backup = os.path.join(pasta, 'BACKUP', arquivo)
                    if os.path.exists(caminho_completo) and not os.path.exists(caminho_backup):
                        os.makedirs(os.path.dirname(caminho_backup), exist_ok=True)
                        os.rename(caminho_completo, caminho_backup)
                    
                    arquivos_processados += 1
                    progresso_var.set(f'Criação e Edição: {arquivos_processados}/{total_arquivos} arquivos processados')
                    
                except Exception as e:
                    arquivos_com_erro += 1
                    progresso_var.set(f'Erro: {e}')
    
    # Consolidação dos arquivos CSV
    consolidar_csv(caminho, fila, progresso_var)

# Função para consolidar arquivos CSV
def consolidar_csv(caminho, fila, progresso_var):
    global arquivos_processados
    for pasta, _, arquivos in os.walk(caminho):
        arquivos_csv = [arq for arq in arquivos if arq.endswith('.csv')]
        df_total = pd.concat([pd.read_csv(os.path.join(pasta, arq)) for arq in arquivos_csv])
        df_total.to_csv(os.path.join(pasta, 'consolidado.csv'), index=False)
        arquivos_processados += len(arquivos_csv)
        progresso_var.set(f'Consolidação: {arquivos_processados}/{total_arquivos} arquivos processados')
        fila.put('Consolidação concluída')
        time.sleep(10)  # Espera 10 segundos antes de concluir

# Função para atualizar a interface gráfica
def atualizar_interface(fila, janela, progresso_var, progresso_barra):
    global inicio_processamento
    while True:
        try:
            mensagem = fila.get_nowait()
            if mensagem == 'Consolidação concluída':
                tempo_total = time.time() - inicio_processamento
                messagebox.showinfo('Conclusão', f'Processamento concluído com sucesso! Tempo total: {tempo_total:.2f} segundos')
                janela.destroy()
                break
            else:
                progresso_barra['value'] += 1
                janela.update_idletasks()
        except queue.Empty:
            continue

# Função principal
def main():
    fila = queue.Queue()
    progresso_var = queue.Queue()
    
    # Configuração da interface gráfica
    janela = Tk()
    janela.title('Processamento de Arquivos Excel')
    
    # Label para mostrar informações do progresso
    label_progresso = Label(janela, textvariable=progresso_var)
    label_progresso.pack(pady=10)
    
    # Barra de progresso
    progresso_barra = ttk.Progressbar(janela, orient='horizontal', length=300, mode='determinate')
    progresso_barra.pack(pady=20)
    
    # Botão para iniciar o processamento
    botao_iniciar = Button(janela, text='Iniciar Processamento', command=lambda: threading.Thread(target=processar_arquivos_excel, args=(os.getcwd(), fila, progresso_var)).start())
    botao_iniciar.pack(pady=10)
    
    # Botão para cancelar o processo
    botao_cancelar = Button(janela, text='Anular Processo', command=lambda: messagebox.askyesno('Cancelar', 'Deseja cancelar o processamento?'))
    botao_cancelar.pack(pady=10)
    
    # Inicia a thread de atualização da interface
    threading.Thread(target=atualizar_interface, args=(fila, janela, progresso_var, progresso_barra), daemon=True).start()
    
    janela.mainloop()

if __name__ == '__main__':
    main()
