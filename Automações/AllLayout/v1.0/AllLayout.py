import os
from openpyxl import load_workbook
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import zipfile
import time
import win32com.client as win32

# Caminho do arquivo
caminho_base_formula = r'F:\01-AllCare\FIN01-Financeiro\Comissionamento\LÇTO NF\\Base Formula CSV - Giovanni V1 - Python.xlsx'

# Verificar se o arquivo existe
arquivo_existe = os.path.exists(caminho_base_formula)

# Caminho para o ChromeDriver
caminho_chromedriver = r'F:\01-AllCare\FIN01-Financeiro\Comissionamento\LÇTO NF\Apps\ChromeDriver\\chromedriver.exe'

# Caminho para a pasta de downloads
pasta_download = os.path.join(os.path.expanduser('~'), 'Downloads')

# Caminho para a pasta de destino
pasta_destino = r'F:\\01-AllCare\\FIN01-Financeiro\\Comissionamento\\LÇTO NF'

# Configuração do WebDriver
options = Options()
options.add_argument("--start-maximized")  # Inicia o navegador maximizado
options.add_experimental_option('prefs', {
    "download.default_directory": pasta_download,  # Define a pasta de download padrão
    "download.prompt_for_download": False,  # Impede que o navegador solicite confirmação para download
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# Inicializar o WebDriver com as opções configuradas
servico = Service(caminho_chromedriver)
driver = webdriver.Chrome(service=servico, options=options)

# Solicita ao usuário as datas de início e fim para baixar
data_inicio_baixar_str = input("Digite a data de início para baixar (dd/mm/aaaa): ")
data_fim_baixar_str = input("Digite a data de fim para baixar (dd/mm/aaaa): ")

# Converte as strings de data para objetos datetime
data_inicio_baixar = datetime.strptime(data_inicio_baixar_str, '%d/%m/%Y')
data_fim_baixar = datetime.strptime(data_fim_baixar_str, '%d/%m/%Y')

print(f'Período de tempo para LAYOUT: {data_inicio_baixar} a {data_fim_baixar}')



# Função para iniciar o processo de download
def iniciar_download():
    global data_inicio_baixar, data_fim_baixar

    def processo_downLoadLayout():

        try:
            print("Abrindo a página de login...")
            driver.get('https://alltech.allcare.com.br/AllTechLoginSistema')

            time.sleep(2)

            print("Fazendo login...")
            campo_usuario_2 = driver.find_element(By.ID, 'UserNameLogin')
            campo_usuario_2.send_keys('giovanni.moreira')

            campo_senha_2 = driver.find_element(By.ID, 'UserPwdLogin')
            campo_senha_2.send_keys('25663851.D1a')

            botao_login_2 = driver.find_element(By.ID, 'btnConectar')
            botao_login_2.click()

            time.sleep(2)

            print("Acessando a página de notas recebidas...")
            driver.get('https://alltech.allcare.com.br/AllTechTool?CodTool=5692')

            time.sleep(2)

            print("Aplicando filtro de datas...")
            time.sleep(1)

            data_inicio_str = data_inicio_baixar.strftime('%d/%m/%Y')
            data_fim_str = data_fim_baixar.strftime('%d/%m/%Y')

            campos_data = driver.find_elements(By.CSS_SELECTOR, '.DateBox')
            campos_data[0].send_keys(data_inicio_str)
            campos_data[1].send_keys(data_fim_str)

            botao_aplicar = driver.find_element(By.XPATH, '//*[contains(@id, "FiltroBtnConsultar_")]').click()

            time.sleep(2)

            print("Exportando dados...")
            driver.find_element(By.CSS_SELECTOR, '.StripButtonGrid').click()
            time.sleep(1)

            campo_exportar = driver.find_elements(By.CSS_SELECTOR, '.StripMenu li a')[1]
            campo_exportar.click()
            time.sleep(5)

            arquivos_antes = set(os.listdir(pasta_download))

            print("Aguardando para baixar o arquivo...")
            while True:
                try:
                    campo_baixar = WebDriverWait(driver, 2).until(
                        EC.visibility_of_element_located((By.CSS_SELECTOR, '.MensagemSistemaBotoes > input#MensagemSistema_Down'))
                    )
                    campo_baixar.click()
                    time.sleep(12)
                    print("Arquivo baixado com sucesso!")
                    break
                except Exception as e:
                    print("Botão não disponível ainda, tentando novamente...")

            print("Verificando se o arquivo foi baixado...")
            arquivo_baixado = None
            for _ in range(10):
                arquivos_depois = set(os.listdir(pasta_download))
                novos_arquivos = arquivos_depois - arquivos_antes
                
                if novos_arquivos:
                    arquivo_baixado = novos_arquivos.pop()
                    print(f"Arquivo baixado: {arquivo_baixado}")
                    break
                time.sleep(2)

            if not arquivo_baixado:
                print("Erro: Nenhum arquivo novo foi encontrado no diretório de downloads.")
            
        finally:
            driver.quit()
            processo_renomear_layout()
            Excel_Base(caminho_base_formula)
            
            
    processo_downLoadLayout()
    #Excel_Base(caminho_base_formula)

# done -- 
def processo_renomear_layout():
    lista_arquivos = os.listdir(pasta_download)
    caminho_arquivo_recente = max([os.path.join(pasta_download, f) for f in lista_arquivos], key=os.path.getctime)

    print("Descompactando o arquivo baixado...")
    if zipfile.is_zipfile(caminho_arquivo_recente):
        with zipfile.ZipFile(caminho_arquivo_recente, 'r') as zip_ref:
            zip_ref.extractall(pasta_destino)

    print("Apagando o arquivo 'Notas.csv' existente...")
    arquivo_antigo = os.path.join(pasta_destino, 'Notas.csv')
    if os.path.exists(arquivo_antigo):
        os.remove(arquivo_antigo)

    print("Renomeando o novo arquivo para 'Notas.csv'...")
    arquivos_extraidos = os.listdir(pasta_destino)
    novo_arquivo = max([os.path.join(pasta_destino, f) for f in arquivos_extraidos if f != 'Notas.csv'], key=os.path.getctime)
    os.rename(novo_arquivo, os.path.join(pasta_destino, 'Notas.csv'))

    print("Processo concluído! Esperando por 5 segundos antes de encerrar...")
    print("Encerrando script.")
    time.sleep(5)


def Excel_Base(caminho_base_formula):
    try:
        # Inicia uma nova instância do Excel sem tentar definir a visibilidade imediatamente
        excel = win32.DispatchEx('Excel.Application')  # Usar DispatchEx para garantir uma nova instância
        excel.Visible = True  # Se ainda der erro, pode remover esta linha e tentar definir a visibilidade mais tarde
        
        print('Abrindo arquivo Base formula')
        wb_base = excel.Workbooks.Open(caminho_base_formula)
        sheet_base = wb_base.Sheets(1)
        
        # Remove os filtros existentes, se houver
        if sheet_base.AutoFilterMode:
            sheet_base.AutoFilterMode = False
        
        print('Atualizando Base formula')
        contagem1 = 0
        while contagem1 <= 5: 
            wb_base.RefreshAll()
            contagem1 += 1
            print(f'Atualizando base {contagem1}')
            time.sleep(2)
        
        time.sleep(5)
        
        # Aplica filtro na célula AE1 para "Allcare Administradora - SP", "Allcare Administradora - DF", e "Allcare Administradora - RJ"
        print('Aplicando filtro na célula AE1')
        sheet_base.Range("AE1").AutoFilter(Field=31, Criteria1=["Allcare Administradora - SP", "Allcare Administradora - DF", "Allcare Administradora - RJ"], Operator=7)
        
        time.sleep(5)
        
        # Chama a função para abrir o segundo arquivo e copiar os dados filtrados
        caminho_filefill = r'F:\01-AllCare\FIN01-Financeiro\Comissionamento\LÇTO NF\\Arquivo de Preenchimento - Giovanni V1 - Copia.xlsx'
        File_Fill(excel, sheet_base, caminho_filefill)

        print("Arquivo Base formula atualizado e filtrado!")
        
    except Exception as e:
        print(f"Erro ao abrir o arquivo Base formula: {e}")

def File_Fill(excel, sheet_base, caminho_filefill):
    try:
        
        time.sleep(5)
        
        # 2. Continuar com o File Fill
        print('Abrindo arquivo File Fill')
        
        # Desativa alertas para evitar o prompt manual
        excel.DisplayAlerts = False
        
        wb_fill = excel.Workbooks.Open(caminho_filefill, UpdateLinks=1)  # UpdateLinks=1 aceita automaticamente a atualização dos vínculos
        sheet_fill = wb_fill.Sheets(1)

        contagem2 = 0
        while contagem2 <=5: 
            wb_fill.RefreshAll()
            contagem2 = contagem2 +1
            print(f'Atualizando base {contagem2}')
            time.sleep(2)
        
        time.sleep(5)
        # Aguarda alguns segundos para garantir que o arquivo seja carregado corretamente
        time.sleep(5)

        # Apaga as células do range A3:AE3 para baixo
        print('Limpando células de A3 até AE3')
        last_row = sheet_fill.Cells(sheet_fill.Rows.Count, "A").End(-4162).Row  # -4162 é xlUp
        sheet_fill.Range(f"A3:AE{last_row}").ClearContents()

        # Copia linha por linha dos dados filtrados do Excel base e cola no File Fill
        print('Copiando dados filtrados do Excel base linha por linha')
        visible_rows = sheet_base.Range("A2:AE" + str(sheet_base.Cells(sheet_base.Rows.Count, "A").End(-4162).Row)).SpecialCells(12)  # 12 = xlCellTypeVisible
        
        target_row = 3  # Começar a colar na célula A3 do File Fill
        for row in visible_rows.Rows:
            row.Copy(sheet_fill.Range(f"A{target_row}"))
            target_row += 1
        
        time.sleep(3)

        # Chama a função DownForm para descer as fórmulas nas colunas especificadas
        DownForm(sheet_fill, target_row - 1, excel)

        print("Arquivo File Fill atualizado com os dados filtrados e fórmulas!")
        
    except Exception as e:
        print(f"Erro ao abrir o arquivo File Fill: {e}")
    finally:
        # wb_fill.Save()  # Salva o File Fill
        excel.DisplayAlerts = True  # Reativa os alertas

def DownForm(sheet_fill, last_row, excel):
    try:
        print('Copiando fórmulas da primeira linha para o restante da coluna com base no range de dados colados')
        time.sleep(8)
        
        # Define as colunas que devem ter suas fórmulas copiadas
        colunas_com_formulas = ['A', 'C', 'F', 'G', 'H', 'M', 'P', 'U', 'V', 'X', 'Y']

        # Itera sobre cada coluna
        for col in colunas_com_formulas:
            formula_source = sheet_fill.Range(f"{col}1")

            if formula_source.HasFormula:  # Verifica se a célula de origem tem uma fórmula
                # Define o intervalo destino, começando na terceira linha até a última linha de dados colados
                formula_destination = sheet_fill.Range(f"{col}3:{col}{last_row}")
                formula_destination.Formula = formula_source.Formula  # Aplica a fórmula em todo o intervalo de uma vez

                print(f"Fórmula de {col}1 copiada para o intervalo {col}3:{col}{last_row}")

        print('Fórmulas aplicadas com sucesso!')

        # Configurações de segurança para ignorar alertas
        excel.DisplayAlerts = False  # Desativa alertas temporários
        excel.Application.AutomationSecurity = 2  # Forçar desativação de macros e alertas de segurança

        # Abrir e atualizar o Geral de Comissões
        caminho_geral_comissoes = r'F:\\01-AllCare\\FIN01-Financeiro\\Comissionamento\\1. CONTROLE COMISSÃO_TAXAS_FINANCEIRO\\2022\\CONTROLE GERAL DE COMISSÃO.xlsx'
        print('Abrindo e atualizando o arquivo Geral de Comissões')

        wb_geral_comissoes = excel.Workbooks.Open(caminho_geral_comissoes, UpdateLinks=1)  # UpdateLinks=1 aceita automaticamente a atualização dos vínculos
        sheet_geral_comissoes = wb_geral_comissoes.Sheets(1)

        # Verificar se o arquivo foi aberto corretamente
        if wb_geral_comissoes:
            print("Arquivo Geral de Comissões aberto com sucesso.")
        else:
            raise Exception("Erro ao abrir o arquivo Geral de Comissões.")

        # Atualizar todas as conexões de dados no arquivo
        wb_geral_comissoes.RefreshAll()
        print('Aguardando a atualização do arquivo Geral de Comissões...')
        time.sleep(70)  # Tempo para garantir que as atualizações sejam concluídas

        print('Atualização do Geral de Comissões concluída.')

    except Exception as e:
        print(f"Erro ao descer as fórmulas ou atualizar o Geral de Comissões: {e}")

    finally:
        # Restaurar configurações
        excel.Application.AutomationSecurity = 1  # Restaurar para o padrão de segurança
        excel.DisplayAlerts = True  # Reativa os alertas   
        
# Iniciar o processo de download
iniciar_download()
