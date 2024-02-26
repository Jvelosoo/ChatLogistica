
import os
import schedule
import pandas as pd
import time
import pyautogui
import pyperclip

df_logistica = pd.read_excel()#Caminho planilha
colunas_desejadas = ['Unidade', 'Coordenação', 'Celular']
df_geral = pd.read_excel("", header=2, usecols=colunas_desejadas, sheet_name=1) #Caminho planilha
linkApp = '' # Link aplicativo geração de memorandos

def abrir_whatsapp_web():
    os.system("start chrome https://web.whatsapp.com")
    time.sleep(15) 

def enviar_mensagem_coordenador(local):
    coordenador = df_geral.loc[df_geral['Unidade'] == local, 'Coordenação'].iloc[0]
    return coordenador

def enviar_mensagens_coordenadores():

    abrir_whatsapp_web()

    for _, linha in df_logistica.iterrows():
        local = linha["Local Do Chamado"]
        data = linha["Data Da Solicitação"]
        retirada = linha["Retirada"]
        descricao = linha["Descrição"]
        nChamado = linha["Chamado"]

        nome = enviar_mensagem_coordenador(local)
        if nome in ["Joãozinho"]: # Regra de negocio 
            nome = "Benzinho"
        else:
             nome = enviar_mensagem_coordenador(local)   

        if retirada.lower() == "não":
                mensagem = f'*Atenção {nome}* \n\n' \
                           f'A partir de  *{data}*, será realizada a  *{descricao}* na {local}. ' \
                           f' Chamado: {nChamado} \n\n' \
                           f'Esta é uma mensagem automática do setor CSMB/STI, Favor não responder está mensagem, obrigado pela atenção.'                            
        else:
                mensagem = f'*Atenção {nome}* \n\n' \
                           f'A partir de *{data}*, será realizada a  *{descricao}* na {local}. ' \
                           f' Chamado: {nChamado} \n\n' \
                           f'Esta é uma mensagem automática do setor CSMB/STI, Favor não responder está mensagem, obrigado pela atenção. \n\n'   \
                           f'Por favor, faça o memorando utilizando nosso aplicativo. \n\n'   \
                           f'{linkApp}'  
                         
        pyautogui.click(320, 180) 
        pyperclip.copy(nome)
        pyautogui.hotkey('ctrl', 'v')       
        pyautogui.press('enter')
        time.sleep(4)
        pyperclip.copy(mensagem)
        pyautogui.hotkey('ctrl', 'v')    
        pyautogui.press('enter')
        time.sleep(4)

def enviar_mensagens_grupo():

    abrir_whatsapp_web()

    locais = df_logistica['Local Do Chamado'].unique()

    for local in locais:
        for _, linha in df_logistica[df_logistica['Local Do Chamado'] == local].iterrows():
            data = linha["Data Da Solicitação"]
            entrega = linha["Entrega"]
            retirada = linha["Retirada"]

            if retirada.lower() == "não":
                mensagem_grupo = f'*ENTREGA* solicitada na data *{data}* entrega de *{entrega}* no local {local}. ' \
                                  f'Esta é uma mensagem automática do setor de TI. \n\n' \
                                      
            else:
                mensagem_grupo = f'*RETIRADA* solicitada na data *{data}* retirada de *{retirada}* no local {local}. ' \
                                  f'Esta é uma mensagem automática do setor de TI.\n\n'

            pyautogui.click(320, 180) 
            pyautogui.write("teste") 
            pyautogui.press('enter')
            time.sleep(4)
            pyperclip.copy(mensagem_grupo)
            pyautogui.hotkey('ctrl', 'v')    
            pyautogui.press('enter')
            time.sleep(4)

    os.system("taskkill /f /im chrome.exe")

schedule.every().day.at("13:15").do(enviar_mensagens_coordenadores)
schedule.every().day.at("13:15").do(enviar_mensagens_grupo)

while True:
    schedule.run_pending()
    time.sleep(1)
