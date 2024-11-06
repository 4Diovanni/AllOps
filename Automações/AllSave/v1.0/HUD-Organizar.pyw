import os
import shutil
import tkinter as tk
from tkinter import scrolledtext
import threading
import sys
import pandas as pd
from datetime import datetime


# Variáveis globais
pasta_origem = r'F:\01-AllCare\FIN01-Financeiro\Comissionamento\LÇTO NF\Notas Recebidas\Notas'
destino_base_adm = r'F:\01-AllCare\FIN01-Financeiro\Comissionamento\ALLCARE ADM\2024'
destino_base_fapes = r'F:\01-AllCare\FIN01-Financeiro\Comissionamento\FAPES\2024'
prefixos_adm = ['ADM']
prefixos_fapes = ['FAPES']
estados = ['SP', 'RN', 'MG', 'MA', 'BA', 'DF', 'ES', 'RJ', 'UB']
# Controle para saber quando ambas as cópias (ADM e FAPES) terminarem
processos_concluidos = {"ADM": False, "FAPES": False}

def copiar_arquivos_por_prefixo(pasta_origem, prefixos, destino_base, nome_pasta, data_atual, pasta_destino):
    try:
        # Verifica se a pasta de origem existe
        if not os.path.exists(pasta_origem):
            output_text.insert(tk.END, f"A pasta de origem {pasta_origem} não existe.\n")
            return

        # Filtra os arquivos que começam com os prefixos definidos
        arquivos = [f for f in os.listdir(pasta_origem) if any(f.startswith(prefixo) for prefixo in prefixos)]
        
        # Se não houver arquivos correspondentes, não cria a pasta e informa o usuário
        if not arquivos:
            output_text.insert(tk.END, f"Nenhum arquivo encontrado para {nome_pasta}. Pasta não será criada.\n")
            processos_concluidos[nome_pasta] = True  # Marca o processo como concluído
            verificar_processos_concluidos()
            return

        output_text.insert(tk.END, f"Encontrados {len(arquivos)} arquivos para {nome_pasta}.\n")

        # Define o caminho final de destino
        destino_final = os.path.join(destino_base, pasta_destino, data_atual)
        
        # Cria a pasta de destino se não existir
        if not os.path.exists(destino_final):
            os.makedirs(destino_final)
            output_text.insert(tk.END, f"Criada a pasta {destino_final}\n")

        # Copia os arquivos para o destino final
        for arquivo in arquivos:
            origem = os.path.join(pasta_origem, arquivo)
            destino = os.path.join(destino_final, arquivo)
            shutil.copy(origem, destino)
            output_text.insert(tk.END, f"Copiando {arquivo} para {destino_final}\n")

        # Marca o processo como concluído
        processos_concluidos[nome_pasta] = True
        verificar_processos_concluidos()

    except Exception as e:
        # Tratamento de erros
        output_text.insert(tk.END, f"Ocorreu um erro: {e}\n")
        processos_concluidos[nome_pasta] = True  # Mesmo em caso de erro, marca como concluído para evitar travamentos
        verificar_processos_concluidos()

def verificar_processos_concluidos():
    if all(processos_concluidos.values()):
        output_text.insert(tk.END, "\n\nProcesso de organização de ADM e FAPES concluído com sucesso!\n")
        gerar_relatorio_excel()

def iniciar_organizar_adm_fapes():
    data_atual = input_box.get().strip()
    if not data_atual:
        output_text.insert(tk.END, "Data não informada. Por favor, insira a data no formato dd.mm.yy.\n")
        return
    # Resetar controle de processos
    processos_concluidos["ADM"] = False
    processos_concluidos["FAPES"] = False
    
    # Listar as subpastas disponíveis para ADM
    subpastas_adm = listar_pastas(destino_base_adm)
    if not subpastas_adm:
        return
    mostrar_pastas_disponiveis(subpastas_adm, "ADM")
    root.bind('<Return>', lambda event: capturar_input_adm(subpastas_adm, data_atual))

def capturar_input_adm(subpastas_adm, data_atual):
    numero = input_box.get().strip()
    if numero.isdigit():
        numero = int(numero) - 1
        if 0 <= numero < len(subpastas_adm):
            pasta_destino_adm = subpastas_adm[numero]
            output_text.insert(tk.END, f"Pasta ADM selecionada: {pasta_destino_adm}\n")

            # Listar as subpastas disponíveis para FAPES
            subpastas_fapes = listar_pastas(destino_base_fapes)
            if not subpastas_fapes:
                return
            mostrar_pastas_disponiveis(subpastas_fapes, "FAPES")
            root.bind('<Return>', lambda event: capturar_input_fapes(subpastas_fapes, data_atual, pasta_destino_adm))
            input_box.delete(0, tk.END)
        else:
            output_text.insert(tk.END, "Escolha inválida. Tente novamente.\n")
    else:
        output_text.insert(tk.END, "Entrada inválida. Por favor, insira um número.\n")

def capturar_input_fapes(subpastas_fapes, data_atual, pasta_destino_adm):
    numero = input_box.get().strip()
    if numero.isdigit():
        numero = int(numero) - 1
        if 0 <= numero < len(subpastas_fapes):
            pasta_destino_fapes = subpastas_fapes[numero]
            output_text.insert(tk.END, f"Pasta FAPES selecionada: {pasta_destino_fapes}\n")

            # Iniciar a cópia dos arquivos para ADM e FAPES
            threading.Thread(target=copiar_arquivos_por_prefixo, args=(pasta_origem, prefixos_adm, destino_base_adm, "ADM", data_atual, pasta_destino_adm)).start()
            threading.Thread(target=copiar_arquivos_por_prefixo, args=(pasta_origem, prefixos_fapes, destino_base_fapes, "FAPES", data_atual, pasta_destino_fapes)).start()

            input_box.delete(0, tk.END)
        else:
            output_text.insert(tk.END, "Escolha inválida. Tente novamente.\n")
    else:
        output_text.insert(tk.END, "Entrada inválida. Por favor, insira um número.\n")

def iniciar_organizar_corretora():
    global etapa_atual, caminho_base, destino_final
    etapa_atual = 1
    caminho_base = r'F:\01-AllCare\FIN01-Financeiro\Comissionamento\ALLCARE CORRETORA\2024'
    pastas = listar_pastas(caminho_base)
    
    if not pastas:
        return
    
    mostrar_pastas_disponiveis(pastas, "CORRETORA")
    root.bind('<Return>', lambda event: capturar_input(pastas))

def listar_pastas(caminho):
    try:
        pastas = [f.name for f in os.scandir(caminho) if f.is_dir()]
        return pastas
    except FileNotFoundError:
        output_text.insert(tk.END, f"Caminho não encontrado: {caminho}\n")
        return []

def mostrar_pastas_disponiveis(pastas, indicador):
    output_text.insert(tk.END, f"Pastas disponíveis em {indicador}:\n")
    for i, pasta in enumerate(pastas):
        output_text.insert(tk.END, f"{i+1}: {pasta}\n")

def capturar_input(pastas):
    global etapa_atual, pasta_selecionada, subpastas, caminho_selecionado, destino_final, numero_parte

    if etapa_atual == 1:
        numero = input_box.get().strip()
        if numero.isdigit():
            numero = int(numero) - 1
            if 0 <= numero < len(pastas):
                pasta_selecionada = pastas[numero]
                output_text.insert(tk.END, f"Pasta selecionada: {pasta_selecionada}\n")
                caminho_selecionado = os.path.join(caminho_base, pasta_selecionada)
                subpastas, caminho_selecionado = listar_subpastas(caminho_selecionado)
                etapa_atual = 2
                root.bind('<Return>', lambda event: capturar_input(subpastas))
                input_box.delete(0, tk.END)
                return
        output_text.insert(tk.END, "Escolha inválida. Tente novamente.\n")

    elif etapa_atual == 2:
        numero = input_box.get().strip()
        if numero.isdigit():
            numero = int(numero)
            if numero == len(subpastas) + 1:
                output_text.insert(tk.END, "Digite a data desejada no formato dd.mm.yyyy:\n")
                
                def capturar_data(event):
                    global numero_parte
                    data_especificada = input_box.get().strip()
                    nova_pasta = os.path.join(caminho_selecionado, data_especificada)
                    if not os.path.exists(nova_pasta):
                        os.makedirs(nova_pasta)
                        output_text.insert(tk.END, f"Criada a pasta {nova_pasta}\n")
                    root.unbind('<Return>')
                    numero_parte = input_parte.get().strip()
                    if numero_parte:
                        threading.Thread(target=organizar_arquivos_por_estado, args=(pasta_origem, estados, numero_parte, nova_pasta)).start()
                    else:
                        output_text.insert(tk.END, "Entrada inválida. Digite a parte desejada.\n")

                root.bind('<Return>', capturar_data)
                input_box.delete(0, tk.END)
                return

            elif 0 <= numero <= len(subpastas):
                destino_final = os.path.join(caminho_selecionado, subpastas[numero - 1])
                output_text.insert(tk.END, f"Subpasta selecionada: {destino_final}\n")
                etapa_atual = 3
                root.bind('<Return>', lambda event: capturar_input([]))
                input_box.delete(0, tk.END)
                return
        output_text.insert(tk.END, "Escolha inválida. Tente novamente.\n")
    
    elif etapa_atual == 3:
        numero_parte = input_parte.get().strip()
        if numero_parte:
            threading.Thread(target=organizar_arquivos_por_estado, args=(pasta_origem, estados, numero_parte, destino_final)).start()
        else:
            output_text.insert(tk.END, "Entrada inválida. Digite a parte desejada.\n")
        root.bind('<Return>', None)

def listar_subpastas(caminho):
    subpastas = listar_pastas(caminho)
    if subpastas:
        output_text.insert(tk.END, "Subpastas disponíveis:\n")
        for i, subpasta in enumerate(subpastas):
            output_text.insert(tk.END, f"{i+1}: {subpasta}\n")
        output_text.insert(tk.END, f"{len(subpastas)+1}: Criar nova subpasta com uma data especificada\n")
    else:
        output_text.insert(tk.END, "Nenhuma subpasta encontrada.\n")
    return subpastas, caminho

def organizar_arquivos_por_estado(pasta_origem, estados, numero_parte, destino_final):
    try:
        if not os.path.exists(pasta_origem):
            output_text.insert(tk.END, f"A pasta de origem {pasta_origem} não existe.\n")
            return

        arquivos = [f for f in os.listdir(pasta_origem) if f.lower().endswith('.pdf')]
        output_text.insert(tk.END, f"Encontrados {len(arquivos)} arquivos PDF na pasta de origem.\n")

        pasta_parte = os.path.join(pasta_origem, f"PARTE {numero_parte}")
        if not os.path.exists(pasta_parte):
            os.makedirs(pasta_parte)
            output_text.insert(tk.END, f"Criada a pasta {pasta_parte}\n")

        for estado in estados:
            arquivos_estado = [f for f in arquivos if f.startswith(estado)]
            if arquivos_estado:
                pasta_estado = os.path.join(pasta_parte, estado)
                if not os.path.exists(pasta_estado):
                    os.makedirs(pasta_estado)
                    output_text.insert(tk.END, f"Criada a pasta {pasta_estado}\n")

                for arquivo in arquivos_estado:
                    origem = os.path.join(pasta_origem, arquivo)
                    destino = os.path.join(pasta_estado, arquivo)
                    shutil.copy(origem, destino)
                    output_text.insert(tk.END, f"Copiando {arquivo} para {pasta_estado}\n")
            else:
                output_text.insert(tk.END, f"Nenhum arquivo encontrado para o estado {estado}. Pasta não será criada.\n")

        shutil.move(pasta_parte, destino_final)
        output_text.insert(tk.END, f"Pasta {pasta_parte} movida para {destino_final}\n")

        # Exibir mensagem de conclusão
        output_text.insert(tk.END, "\n\nProcesso de organização de arquivos da corretora concluído com sucesso!\n")
        gerar_relatorio_excel()


    except Exception as e:
        output_text.insert(tk.END, f"Ocorreu um erro: {e}\n")
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

# Janela principal
root = tk.Tk()
root.title("Organizador de Arquivos")

# Ajustes no tamanho e posicionamento da janela
root.geometry("580x600")
root.resizable(False, False)

# Frame principal com padding
frame = tk.Frame(root, padx=20, pady=20)
frame.pack(expand=True, fill=tk.BOTH)

# Título do aplicativo
title_label = tk.Label(frame, text="Organizador de Arquivos", font=("Arial", 18, "bold"))
title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

# Campo para capturar o número da parte ou data
label = tk.Label(frame, text="Digite o número da parte ou a data (dd.mm.yy):", font=("Arial", 10))
label.grid(row=1, column=0, sticky="e", padx=(0, 10))

input_box = tk.Entry(frame, font=("Arial", 10))
input_box.grid(row=1, column=1, sticky="w", ipadx=5, ipady=5)

# Campo para capturar número da parte
label_parte = tk.Label(frame, text="Número da parte:", font=("Arial", 10))
label_parte.grid(row=2, column=0, sticky="e", padx=(0, 10), pady=(10, 0))

input_parte = tk.Entry(frame, font=("Arial", 10))
input_parte.grid(row=2, column=1, sticky="w", ipadx=5, ipady=5, pady=(10, 0))

# Botões de ação
btn_adm_fapes = tk.Button(frame, text="Organizar ADM/FAPES", command=iniciar_organizar_adm_fapes, bg="#4CAF50", fg="white", font=("Arial", 10))
btn_adm_fapes.grid(row=3, column=0, pady=(20, 0), sticky="e")

btn_corretora = tk.Button(frame, text="Organizar CORRETORA", command=iniciar_organizar_corretora, bg="#2196F3", fg="white", font=("Arial", 10))
btn_corretora.grid(row=3, column=1, pady=(20, 0), sticky="w")

# Área de texto para a saída
output_text = scrolledtext.ScrolledText(frame, width=70, height=15, font=("Arial", 10))
output_text.grid(row=4, column=0, columnspan=2, pady=20)

# Botão "Sair"
btn_sair = tk.Button(frame, text="Sair", command=sair, bg="#f44336", fg="white", font=("Arial", 10))
btn_sair.grid(row=5, column=0, pady=(10, 0), ipadx=10, ipady=5, sticky="w")

# Botão "Reiniciar"
btn_reiniciar = tk.Button(frame, text="Reiniciar", command=reiniciar_aplicativo, bg="#FFC107", fg="black", font=("Arial", 10))
btn_reiniciar.grid(row=5, column=1, pady=(10, 0), ipadx=10, ipady=5, sticky="e")

# Botão "Relatório"
btn_relatorio = tk.Button(frame, text="Gerar Relatório", command=gerar_relatorio_excel, bg="#2196F3", fg="white", font=("Arial", 10))
btn_relatorio.grid(row=3, columns=5, pady=(10, 0), ipadx=10, ipady=5, sticky="w")


root.mainloop()
