import os
import shutil
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import zipfile
import time
import win32com.client as win32
import tkinter as tk
from tkinter import scrolledtext
from threading import Thread
import sys
import pandas as pd
from datetime import datetime

diretorio_download = os.path.join(os.path.expanduser('~'), 'Downloads')

# Função para atualizar o texto na janela de saída
def update_output(message):
    output_text.config(state=tk.NORMAL)
    output_text.insert(tk.END, message + '\n')
    output_text.yview(tk.END)
    output_text.config(state=tk.DISABLED)

# Função para iniciar o processo de download
def iniciar_download():
    data_inicio_baixar = data_inicio_entry.get()
    data_fim_baixar = data_fim_entry.get()

    def download_process():
        try:
            options = Options()
            options.add_argument("--start-maximized")
            options.add_experimental_option('prefs', {
                "download.default_directory": pasta_download,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True
            })

            servico = Service(caminho_chromedriver)
            driver = webdriver.Chrome(service=servico, options=options)
            update_output("Acessando Alltech!")
            time.sleep(2)
            driver.get('https://alltech.allcare.com.br/AllTechLoginSistema')
            time.sleep(2)
            update_output("Realizando o login")
            driver.find_element(By.ID, 'UserNameLogin').send_keys('comissoes.allcare')
            driver.find_element(By.ID, 'UserPwdLogin').send_keys('*Comiss0es@2024')
            driver.find_element(By.ID, 'btnConectar').click()
            time.sleep(2)
            update_output("Acessando a página de notas recebidas...")
            driver.get('https://alltech.allcare.com.br/AllTechTool?CodTool=11093')
            time.sleep(2)
            update_output("Aplicando filtro de datas...")
            driver.find_element(By.CSS_SELECTOR, '.StripButtonFiltro').click()
            time.sleep(1)
            campos_data = driver.find_elements(By.CSS_SELECTOR, '.DateBox')
            campos_data[0].send_keys(data_inicio_baixar)
            campos_data[1].send_keys(data_fim_baixar)
            driver.find_element(By.XPATH, '//*[contains(@id, "FiltroBtnConsultar_")]').click()
            time.sleep(2)
            update_output("Exportando dados...")
            driver.find_element(By.CSS_SELECTOR, '.StripButtonGrid').click()
            time.sleep(1)
            driver.find_elements(By.CSS_SELECTOR, '.StripMenu li a')[1].click()
            time.sleep(1)
            # Captura a lista de arquivos no diretório de downloads antes do download
            arquivos_antes = set(os.listdir(diretorio_download))

            # Espera a geração do arquivo e tenta clicar no botão de download até que esteja disponível
            update_output("Aguardando para baixar o arquivo...")
            while True:
                try:
                    campo_baixar = WebDriverWait(driver, 2).until(
                        EC.visibility_of_element_located((By.CSS_SELECTOR, '.MensagemSistemaBotoes > input#MensagemSistema_Down'))
                    )
                    campo_baixar.click()
                    time.sleep(20)
                    update_output("Arquivo baixado com sucesso!")
                    break  # Sai do loop quando o botão é clicado
                except Exception as e:
                    update_output("Botão não disponível ainda, tentando novamente...")

            # Verifica se o arquivo foi baixado, comparando os arquivos antes e depois
            update_output("Verificando se o arquivo foi baixado...")
            arquivo_baixado = None
            for _ in range(10):  # Tenta verificar até 10 vezes
                arquivos_depois = set(os.listdir(diretorio_download))
                novos_arquivos = arquivos_depois - arquivos_antes
                
                if novos_arquivos:
                    arquivo_baixado = novos_arquivos.pop()  # Pega o primeiro (e provavelmente único) novo arquivo
                    update_output(f"Arquivo baixado: {arquivo_baixado}")
                    break
                time.sleep(2)  # Espera um pouco antes de verificar novamente

            if not arquivo_baixado:
                update_output("Erro: Nenhum arquivo novo foi encontrado no diretório de downloads.")
        finally:
            driver.quit()
            criar_arquivo_alterar()
            apagar_arquivos_pasta(pasta_destino_2)
            coletar_notas_fiscais()
            renomear_arquivos_excel()
            gerar_relatorio_excel()
            # Iniciar outro script Python /// FUNCIONAAA!!
            os.system(r'python "F:\01-AllCare\FIN01-Financeiro\Comissionamento\LÇTO NF\\Apps\\AllSave\\HUD-Organizar.pyw"')


    # Executa o processo de download em uma nova thread
    Thread(target=download_process).start()

# Função para criar o arquivo Alterar.csv
def criar_arquivo_alterar():
    try:
        lista_arquivos = os.listdir(pasta_download)
        caminho_arquivo_recente = max([os.path.join(pasta_download, f) for f in lista_arquivos], key=os.path.getctime)
        shutil.copy2(caminho_arquivo_recente, pasta_destino)
        caminho_arquivo_copiado = os.path.join(pasta_destino, os.path.basename(caminho_arquivo_recente))

        if zipfile.is_zipfile(caminho_arquivo_copiado):
            with zipfile.ZipFile(caminho_arquivo_copiado, 'r') as zip_ref:
                zip_ref.extractall(pasta_destino)
            os.remove(caminho_arquivo_copiado)

        arquivo_antigo = os.path.join(pasta_destino, 'Alterar.csv')
        if os.path.exists(arquivo_antigo):
            os.remove(arquivo_antigo)

        arquivos_extraidos = os.listdir(pasta_destino)
        novo_arquivo = max([os.path.join(pasta_destino, f) for f in arquivos_extraidos if f != 'Alterar.csv'], key=os.path.getctime)
        os.rename(novo_arquivo, os.path.join(pasta_destino, 'Alterar.csv'))
        update_output("Novo arquivo 'Alterar.csv' adicionado!")
    except Exception as e:
        update_output(f"Ocorreu um erro ao criar o arquivo 'Alterar.csv': {str(e)}")

# Função para apagar arquivos da pasta Notas
def apagar_arquivos_pasta(pasta_destino_2):
    # Verifica se o diretório existe
    if os.path.exists(pasta_destino_2):
        update_output(f'A pasta "{pasta_destino_2}" foi encontrada.')
        
        # Percorre todos os arquivos e subpastas na pasta especificada
        for arquivo in os.listdir(pasta_destino_2):
            caminho_arquivo = os.path.join(pasta_destino_2, arquivo)  # Cria o caminho completo do arquivo ou pasta
            try:
                # Verifica se é um arquivo
                if os.path.isfile(caminho_arquivo):
                    update_output(f'Removendo arquivo: {caminho_arquivo}')
                    os.remove(caminho_arquivo)  # Remove o arquivo
                
            except Exception as e:
                print(f'Erro ao remover {caminho_arquivo}: {e}')
    else:
        print(f'A pasta "{pasta_destino_2}" não existe.')

# Função para coletar notas fiscais da pasta
def coletar_notas_fiscais():
    try:
        pasta_origem = r'\\oci.grupo.allcare.com.br\DFS\Share\DIRECTORY_ORACLE\TOPADMINISTRADORA\COMISSOES\NF'
        pasta_destino_notas = os.path.join(pasta_destino, 'Notas')

        def filtrar_arquivos_por_data(origem, inicio, fim):
            arquivos_filtrados = []
            try:
                with os.scandir(origem) as entries:
                    for entry in entries:
                        if entry.is_file():
                            data_modificacao = datetime.fromtimestamp(entry.stat().st_mtime)
                            if inicio <= data_modificacao <= fim:
                                arquivos_filtrados.append(entry.path)
            except Exception as e:
                update_output(f"Erro ao acessar a pasta de origem: {e}")
            return arquivos_filtrados

        def copiar_arquivos(arquivos, destino):
            if not os.path.exists(destino):
                os.makedirs(destino)
            for arquivo in arquivos:
                try:
                    shutil.copy(arquivo, destino)
                    update_output(f'Arquivo {arquivo} copiado para {destino}')
                except Exception as e:
                    update_output(f"Erro ao copiar arquivo {arquivo}: {e}")
        
        data_inicio_baixar = datetime.strptime(data_inicio_entry.get(), '%d/%m/%Y')
        data_fim_baixar = datetime.strptime(data_fim_entry.get(), '%d/%m/%Y') + timedelta(days=1)

        arquivos_filtrados = filtrar_arquivos_por_data(pasta_origem, data_inicio_baixar, data_fim_baixar)
        copiar_arquivos(arquivos_filtrados, pasta_destino_notas)

        update_output(f'{len(arquivos_filtrados)} arquivos foram copiados para {pasta_destino_notas}.')
    except Exception as e:
        update_output(f"Ocorreu um erro ao coletar notas fiscais: {str(e)}")
    finally:
        update_output("Processo de coleta de NFs concluído!")

# Função para executar a macro no Excel
def execute_macro(excel, workbook, macro_name):
    try:
        update_output(f"Executando a macro: {macro_name}")
        excel.Application.Run(f"'{workbook.Name}'!{macro_name}")
    except Exception as e:
        update_output(f"Erro ao executar a macro '{macro_name}': {e}")

# Função para verificar se a macro foi concluída
def check_macro_completion(ws, timeout=None):
    start_time = time.time()
    while True:
        try:
            status = ws.Cells(7, 5).Value  # Linha 7, Coluna E é a 5ª coluna
            if status == "Processo Concluído":
                update_output("Macro concluída com sucesso!")
                break
        except Exception as e:
            update_output(f"Erro ao verificar o status da macro: {str(e)}")
        time.sleep(1)
        
        if timeout and (time.time() - start_time) > timeout:
            update_output("Tempo limite excedido esperando a conclusão da macro.")
            break
        
# Função para renomear arquivos no Excel
def renomear_arquivos_excel():
    try:
        caminho_excel = r'F:\01-AllCare\FIN01-Financeiro\Comissionamento\LÇTO NF\Notas Recebidas\Controle para Alteração v7.xlsm'
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = True  # Definindo como visível para debug
        wb = excel.Workbooks.Open(caminho_excel)

        try:
            wb.RefreshAll()
            update_output('Atualizando Excel')
            time.sleep(5)

            for ws in wb.Sheets:
                if ws.AutoFilterMode:
                    ws.AutoFilterMode = False
            update_output('Removendo todos os filtros existentes')

            ws_macro_alteracao = wb.Sheets['Macro Alteração']
            update_output('Selecionando aba Macro Alteração')
            time.sleep(3)

            ws_macro_alteracao.Range('A2:B1048576').ClearContents()  # Limpa todas as células a partir de A2 e B2 até o final da planilha
            update_output('Apagando o valor nas células A2 e B2 adiante')
            time.sleep(3)

            ws_macro_alteracao.Cells(7, 5).ClearContents()
            update_output('Apagando o valor na célula E7')
            time.sleep(3)
            
            ws_csv_alterar = wb.Sheets['CSV.Alterar']
            update_output('Selecionando aba CSV.Alterar')
            time.sleep(3)

            ws_csv_alterar.AutoFilterMode = False
            ws_csv_alterar.Range('T1').AutoFilter(Field=20, Criteria1='VERDADEIRO')
            update_output('Filtrando conteúdo VERDADEIRO')
            time.sleep(3)
            
            if ws_csv_alterar.AutoFilter is not None:
                linha_destino = 2

                while True:
                    valor_q = ws_csv_alterar.Cells(linha_destino, 17).Value
                    valor_r = ws_csv_alterar.Cells(linha_destino, 18).Value

                    if valor_q is None and valor_r is None:
                        break

                    ws_macro_alteracao.Cells(linha_destino, 1).Value = valor_q
                    ws_macro_alteracao.Cells(linha_destino, 2).Value = valor_r

                    update_output(f'Colando dados na linha {linha_destino}')

                    linha_destino += 1
            else:
                update_output("Erro: Filtro não foi aplicado corretamente.")

            update_output('Executando a macro Renomear')
            execute_macro(excel, wb, 'Renomear')
            update_output('Macro Renomear executada com sucesso')

            update_output('Verificando conclusão da macro...')
            check_macro_completion(ws_macro_alteracao)
            update_output('Verificação de conclusão da macro concluída')
        finally:
            update_output('Salvando e fechando excel!')
            wb.Save()
            wb.Close()
            excel.Application.Quit()
    except Exception as e:
        update_output(f"Ocorreu um erro durante o processo de renomeação no Excel: {str(e)}")

# Função para reiniciar o aplicativo
def reiniciar_aplicativo():
    python = sys.executable
    caminho_script = os.path.abspath(sys.argv[0])
    os.execl(python, python, f'"{caminho_script}"', *sys.argv[1:])

def gerar_relatorio_excel():
    # Obtenha o conteúdo do output_text
    conteudo = output_text.get("1.0", tk.END).strip()

    if not conteudo:
        output_text.insert(tk.END, "Nenhum conteúdo para gerar o relatório.\n")
        return

    # Crie uma lista de strings (linhas)
    linhas = conteudo.split("\n")

    # Converta para um DataFrame
    df = pd.DataFrame(linhas, columns=["Relatório"])

    # Nome do arquivo Excel com data e hora
    nome_arquivo = f"Relatorio_Salvar_{datetime.now().strftime('%d-%m-%Y_%H-%M')}.xlsx"

    # Caminho para salvar o arquivo (pode mudar para o caminho que desejar)
    caminho_pasta = r'F:\01-AllCare\FIN01-Financeiro\Comissionamento\LÇTO NF\Notas Recebidas\Relatórios'
    if not os.path.exists(caminho_pasta):
        os.makedirs(caminho_pasta)

    caminho_arquivo = os.path.join(caminho_pasta, nome_arquivo)

    # Salvar o DataFrame em um arquivo Excel
    df.to_excel(caminho_arquivo, index=False, engine='openpyxl')

    output_text.insert(tk.END, f"\n\nRelatório salvo em {caminho_arquivo}\n")

def sair():
    root.destroy()

# Caminho para o ChromeDriver
caminho_chromedriver = r'F:\01-AllCare\FIN01-Financeiro\Comissionamento\LÇTO NF\Apps\ChromeDriver\chromedriver.exe'

# Caminho para a pasta de downloads
pasta_download = os.path.join(os.path.expanduser('~'), 'Downloads')

# Caminho para a pasta de destino
pasta_destino = r'F:\01-AllCare\FIN01-Financeiro\Comissionamento\LÇTO NF\Notas Recebidas'

# Caminho para a pasta de notas
pasta_destino_2 = r'F:\01-AllCare\FIN01-Financeiro\Comissionamento\LÇTO NF\Notas Recebidas\Notas'

# Configuração da interface gráfica
root = tk.Tk()
root.title("Salvar Notas")

# Ajustes no tamanho e posicionamento da janela
root.geometry("580x600")
root.resizable(False, False)

# Título da aplicação
titulo_label = tk.Label(root, text="Salvar Notas", font=("Helvetica", 18, "bold"))
titulo_label.pack(pady=20)

# Frame principal para centralizar os widgets
main_frame = tk.Frame(root)
main_frame.pack(expand=True, pady=20)

# Sub-frame para os campos de entrada e botões principais
entry_frame = tk.Frame(main_frame)
entry_frame.pack(pady=10)

# Configuração da janela de entrada de dados
tk.Label(entry_frame, text="Data de Início (dd/mm/aaaa):", font=("Arial", 10)).grid(row=0, column=0, sticky="e", pady=5)
data_inicio_entry = tk.Entry(entry_frame, font=("Arial", 12))
data_inicio_entry.grid(row=0, column=1, pady=5, ipadx=5, ipady=5)

tk.Label(entry_frame, text="Data de Fim (dd/mm/aaaa):", font=("Arial", 10)).grid(row=1, column=0, sticky="e", pady=5)
data_fim_entry = tk.Entry(entry_frame, font=("Arial", 12))
data_fim_entry.grid(row=1, column=1, pady=5, ipadx=5, ipady=5)

# Frame para botões principais (Iniciar Download e Gerar Relatório)
top_button_frame = tk.Frame(main_frame)
top_button_frame.pack(pady=10)

# Botão para iniciar o download
start_button = tk.Button(top_button_frame, text="Iniciar Download", command=iniciar_download, bg="#4CAF50", fg="white", font=("Arial", 10))
start_button.grid(row=0, column=0, padx=10, ipadx=10, ipady=5)

# Botão "Relatório"
btn_relatorio = tk.Button(top_button_frame, text="Gerar Relatório", command=gerar_relatorio_excel, bg="#2196F3", fg="white", font=("Arial", 10))
btn_relatorio.grid(row=0, column=1, padx=10, ipadx=10, ipady=5)

# Sub-frame para os botões "Sair" e "Reiniciar"
bottom_button_frame = tk.Frame(main_frame)
bottom_button_frame.pack(pady=10)

# Botão "Sair"
btn_sair = tk.Button(bottom_button_frame, text="Sair", command=sair, bg="#f44336", fg="white", font=("Arial", 10))
btn_sair.grid(row=0, column=0, padx=10, ipadx=10, ipady=5)

# Botão "Reiniciar"
btn_reiniciar = tk.Button(bottom_button_frame, text="Reiniciar", command=reiniciar_aplicativo, bg="#FFC107", fg="black", font=("Arial", 10))
btn_reiniciar.grid(row=0, column=1, padx=10, ipadx=10, ipady=5)

# Configuração da janela de saída
output_frame = tk.Frame(root, padx=20, pady=10)
output_frame.pack(fill=tk.BOTH, expand=True)

output_text = scrolledtext.ScrolledText(output_frame, wrap=tk.WORD, height=10, state=tk.DISABLED, font=("Helvetica", 10))
output_text.pack(fill=tk.BOTH, expand=True)

root.mainloop()

