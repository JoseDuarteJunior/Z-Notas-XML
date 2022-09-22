#Programa que sincroniza e-mails

import PySimpleGUIQt as sg
import imaplib
import email
from datetime import datetime, timedelta
import os
import time
import shutil
import re
import webbrowser
from zipfile import ZipFile
import py7zr
import patoolib
import threading
#---------------------------------------------------Parte Arquivo Config.------------------------------------
configuracoes = open('config2.txt','r')
conteudo_configuracoes = configuracoes.read().splitlines()
configuracoes.close()
cfg_email = conteudo_configuracoes[0]
cfg_senha = conteudo_configuracoes[1]
cfg_dias =  conteudo_configuracoes[2]
cfg_dias = eval(cfg_dias)
cfg_tempo = conteudo_configuracoes[3]
cfg_tempo = eval(cfg_tempo)
cfg_tempo2 = 000000000000000
cfg_server = conteudo_configuracoes[4]
#cfg_server = eval(cfg_server)
cfg_porta = conteudo_configuracoes[5]
cfg_porta = eval(cfg_porta)
cfg_pasta_temp = conteudo_configuracoes[6]
#cfg_pasta_temp = eval(cfg_pasta_temp)
cfg_pasta = conteudo_configuracoes[7]
#cfg_pasta = eval(cfg_pasta)
cfg_tamanho_xml = conteudo_configuracoes[8]
cfg_tamanho_xml = eval(cfg_tamanho_xml)
cfg_limpeza_log = conteudo_configuracoes[9]
cfg_limpeza_log = eval(cfg_limpeza_log)
cfg_pasta_erro = conteudo_configuracoes[10]
cfg_filtros = conteudo_configuracoes[11]
numero = []
#----------------------------------------------------Parte Grafica-------------------------------------------
logo = b'iVBORw0KGgoAAAANSUhEUgAAAgAAAAIACAYAAAD0eNT6AAAACXBIWXMAAAzeAAAM3gETUBB6AAAAIGNIUk0AAHolAACAgwAA+f8AAIDpAAB1MAAA6mAAADqYAAAXb5JfxUYAAByPSURBVHja7N17uJZ1ne/xzw+ELAd1QhjFU5FZk4qp2466NR2dMK/UzG0HFQRRBNGNg6BBYoqpUOxCVJCTk51M0Syx0bwuZXYeNqamZIltHE1Nk5wwJQ9LuPcf62Gu5trZSC581rrv1+u67n+L3/f3eP/ePIt1372qqoqr664k/ZL8jySXJ7kxyc+SPJNkXZLK5XL1+Ov3Se5LckWS/dz3XD316hXesFJK71LKsaWUJUlWJbkqyagkQ5PsnmRAkmJSUAtbJnl/kmFJbiulLC+l7G8s9DQC4I0f/p9M8kCSbyQ5JMlbTAUaZdckt5RSxhgFAqAZB/87Sin/O8n1Sd5nItBovZNcUkr5ilEgAOp9+O+X5O4k+5gG8Cf+qZQy0RgQAPU8/E9K8uMkW5kG8GdcWEo52hgQAPU6/D+XZE6SPqYBvNatIsk/l1L2NQoEQD0O/72SzDcJ4HV4S5LrSynvNQoEQM8+/AcmuS7JW00DeJ3+NsmPSil/ZxQIgJ5repLtjQHYQO9IckMpZTOjQAD0vL/9757kWJMA/kr/Lcl3Sym9jQIB0PP+9m9OwBtxaJLZxoAA6Dl/+/9AkoNNAugCo0spk4wBAdAzfMoIgC50QSnlM8aAAOj+DjMCoAuVJFeUUv67USAAuut/paXsnMTv8AJd7S1Jvl9K+XujQAB0T3sbAbCR/G2SG0spWxsFAqD7GWQEwEb0jnhGAAJAAACNtFeSqzwjAAHQvWxjBMCb4BNJLjEGBED3sUVX/o9NnDhx6ZNPPvlbl8vV/a5ly5Y93Ob7zUmllDPddhEA3UPpyv+xgQMH9h40aNDfuVyu7ndtu+22W3SDe86XSymfdetFAAA07y8dV5RS9jMKBABAs/SNZwQgAAAaacskP/KMAAQAQPPsmGSJZwQgAACaZ88k3/OMAAQAQPMckuRSY0AAADTPiaWUs4wBAQDQPOeXUj5nDAgAgGYpSRaVUvY3CgQAQLP0TXJdKeV9RoEAAGiWLZPcWErxsjIEAEDD7JjkhlLK3xgFAgCgWdY/I2ATo0AAADTL0HhGAAIAoJFGlVK+YAwIAIDmmVZK+bwxIAAAmqUkWVhK+ZhRIAAAmqVvkmtLKbsYBQIAoFm2jGcEIAAAGmmHJEs8IwABANA8eyS52jMCEAAAzfPxJJcZAwIAoHlOKKVMNgYEAEDzTCulHGMMCACA5llQSjnAGBAAAM3iGQEIAICG2iLJj0opg4wCAQDQLNun8xkB/YwCAQDQLO+PZwQgAAAa6R+TzDEGBABA84wspUwxBgEAQPOcV0o51hgEAADNs6CUcqAxCAAAmqVPksWllF2NQgAA0CxbJLnRMwIEAADNs30rAjwjQAAA0DC7J7nGMwIEAADNc3CSucYgAABonhGllLONQQAA0DxfKqUcZwwCAIDmme8ZAQIAgObpk+TaUspuRiEAAGiWzdP564HbGoUAAKBZtkuyxDMCBAAAzbN7Oh8Z7BkBAgCAhjkoyeXGIAAAaJ7jSylTjUEAANA855RShhmDAACgeeaVUv7BGAQAAM3SJ53/KHCIUQgAAJpl83T+eqBnBAgAABpmu3Q+KGhzoxAAADTLkCTXlFL6GIUAAKBZPCNAAADQUMNLKecYgwAAoHmmllKONwYBAEDzzC2lHGQMAgCAZumTzn8U6BkBAgCAhtk8nb8euJ1RCAAAmmXbeEaAAACgkXZL5yODPSNAAADQMP+QZJ4xCAAAmmdYKeVLxiAAAGies0spI4xBAADQPHNLKQcbgwAAoFk2SeczAnY3CgEAQLP0S+evB25vFAIAgGYZ1IqALYxCAADQLLvGMwIEAACNdGCS+cYgAABonuNKKecagwAAoHm+WEoZaQwCAIDmmVNK+UdjEAAANMsmSa4upbzfKAQAAM3SL8kSzwgQAAA0j2cECAAAGmrXJNd6RoAAAKB5DkiywBgEAADNc2wp5TxjEAAANM+UUsoJxiAAAGiey0opHzcGAQBAs6x/RsAeRiEAAGiWv0nnMwJ2MAoBAECzbBPPCBAAADTSLkmuK6X0NQoBAECzfCyeESAAAGikY0op04xBAADQPJNLKaOMQQAA0DyXllKGGoMAAKBZNknyPc8IEAAANI9nBAgAABpqmyQ/KqVsaRQCAIBmeV88I0AAANBI+ydZWEopRiEAAGiWzyc5xRgEAADNc24pZStjEAAANMuWSTwpUAAA0ECjSilvNwYBAEDzzrjdjEEAANA8uxqBAAD4s0op7oX1tYsRCACAP2vrrbcekOQPJlFLg4xAAAC81jcA6d+///81iVra1AgEAMBr2nvvvVebAgIAoGGOPPLIt5kCAgCgYUaOHPmh/v3732cSCACABiml5K677toqyQumgQAAaJCddtpp+7POOusek0AAADTMl7/85f1mzZp1RynlWdNAAAA0yLhx4z7ym9/8Zt2OO+54p2kgAAAaZOuttx7w6KOPfnj58uX/NmPGjNsPPPDAW/v3739f3759H3FtvKtXr15Pd/FWVj7NAgBgg+26667vnDBhwkdvueWWj/3ud7/b4+WXXx7s2njX9OnTu/qBTMWnWAAAAAIAAAQAACAAAAABAAAIAABAAAAAAqA29uri/70hRgrwpt8r9zJSAQAACAAAEAAAgAAAAAQAACAAAAABAAAIAABAAAAAAmBjKqVsVUo5rJQyppQyrZSyqJRyUylleSnl0Q25kry9K/9s55xzzqbvfe9743K5XK7Xvs4555xNu/hoePuG3v9bZ8ZNrTNkWutMOayUspUA6F6H/rtLKRNKKf+a5LdJvp/kkiSTkwxPcnCSXZPsuIFXl87mhRde6LtixYq4XC6X67WvF154oe9GOOc29P6/a+vsGN46Sy5pnS1Pl1JuK6X8z1LKOwRAew79t5RS/qmU8oskDyeZkWRfP9IAYCPqnWS/JP8ryb+VUu4rpYwtpfQRABv/4O9VSjm2deh/Jcnf+zwC0CbvTzI7yUOllM+WUooA2DiH/8FJ7knyjSQ7+NwB0E0MTvLtJD8tpRwkALru4O9bSpmX5KZWbQFAd7RnkptLKfNKKX27+x+2WwdAKWVAkluSnOBzBUAPcUKSW1pnmAD4Kw7/IUnuTuc/7gOAnmTfJHe3zjIBsAGH/8FJ7kjnr2MAQE+0Y5I7WmeaAHgdh//7klydZDOfHQB6uM2SXN062wTAXzj8357kB0k295kBoCY2T/KD1hknAP7M4b9J62/+7/JZAaBm3tX6JmATAfD/m5nkAJ8RAGrqgHQ+RVAA/Mnf/vdOcorPBgA1N7Z15gmAlulJis8FADVXWmeeACilfCLJ/j4TADTE/q2zr7kBUErpleRCnwUAGubC1hnY2G8Ajknne5cBoEl2bZ2BjQ4AAGiiZgZAKWWL+Nk/AM21fymlbQ++a+c3AEOT9LH/ADRUn9ZZ2LgAOMzeA9Bwn2xUAJRS2lo9ANBNHNKuxwO36xuA9yTZwr4D0HBbts7ExgTAIHsOAO07EwUAALTXtgIAAHwDIAAAQAAIAAAQAD08AN5mvwGgfWdiL3MHgOYRAAAgAAAAAQAACAAAQAAAAAIAABAAAIAAAAAEAAAgAAAAAQAACAB4o7bbbrvstNNOLtdfvAYPHuw/FgQA1MmSJUvyy1/+0uX6i9d9993nPxYEAAAgAAAAAQAACAAAQAAAAAIAABAAAIAAAAAEAAAgAACAGgVAMXoAaN+Z2K4A2Mx+A0D7zsR2BcBb7DcAtO9MbFcAbGq/AaB9Z6JvAADANwACAAAEgAAAAAHQhXrbbwBo35noQUAA0EACAAAEAAAgAAAAAQAACAAAQAAAAAIAABAAAIAAAAAEAAAgAAAAAQAA1CMA7jF6AGjfmegbAADwDQAAIACgJnbffff06dPH9Qau2bNn1/5zcscdd/iPBQEAsN4RRxyR0aNH13qNTz75ZIYPH26zEQAASTJo0KB8/etfzyabbFLbNXZ0dOS0007Lb3/7WxuOAABIkoULF2abbbap9RovueSSXH/99TYbAQCQJFOnTs2BBx5Y6zXeeeedOeOMM2w2AgAgSQ444IBMmDCh1mtctWqVn/sjAADW69evX+bMmZNNN920tmtcu3ZtzjjjjDzyyCM2HAEAkHT+3P+d73xnrdd4xRVX5Fvf+pbNRgAAJMlpp52Www8/vNZrvP/++zNmzBibjQAASJK99947U6dOrfUan3vuuYwYMSLr1q2z4QgAgD59+mTevHnp169fbddYVVWmTJmSBx54wIYjAIwASJI5c+Zkl112qfUar7766syZM8dmgwAAkmTYsGE55phjar3GFStWZOTIkTYbBACQJDvvvHMuuuii9OpV39vBmjVrctJJJ+Wll16y4SAAgKTzV/769+9f6zVecMEFuf322202CAAgSWbOnJkPfvCDtV7jjTfemIsuushmgwAAkuTwww/PySefXOs1PvbYYxkxYoTNBgEAJM14xe/LL7+csWPH5tlnn7XhIACApPPn/oMGDar1Gr/2ta/lpptustkgAIAkOfvss2v/it+lS5dmypQpNhsEAJB0vuL3jDPOqPUan376aa/4BQEArNevX79cdtlltX7F76uvvprx48fniSeesOEgAIAkWbBgQQYPHlzrNV5++eW55pprbDYIACBJxo0blyOOOKLWa/zpT3+a8ePH22wQAECS7LXXXvnSl75U6zX++7//e4YPH+4VvyAAgCTp3bt3I17xe+aZZ2bFihU2HAQAkCRz587NbrvtVus1XnnllVm0aJHNBgEAJMlxxx2XY489ttZrfPDBBzN69GibDQIASDpf8Tt9+vRav+L3+eefz6hRo9LR0WHDQQAASTNe8XvOOefk7rvvttkgAIAk+epXv1r7V/xed911mTVrls0GAQAkyWGHHZYxY8bUeo0rV67M8ccfb7NBAABJM17x++KLL2b06NFZs2aNDQcBACTJ/Pnzs+2229Z6jdOnT89tt91ms0EAAEkyZcqUHHTQQbVe480335xp06bZbBAAQJLsv//+mThxYq3X+OSTT3rFLwgAYL3NNtssc+bMyVvf+tbarrGjoyOnnHJKVq1aZcNBAABJ5+/7v+td76r1GmfPnp0bbrjBZoMAAJLklFNOyac+9alar/GOO+6o/Y83QAAAr9uee+6Zc889t9ZrXLVqld/3BwEArNe7d+/Mnz+/1q/4Xbt2bSZMmJBHHnnEhoMAAJJkzpw5tX/F7xVXXJFvf/vbNhsEAJAkxx57bI477rhar/FnP/tZ7R9nDAIAeN2a8Irf1atXZ+TIkVm3bp0NBwEAJJ2P+t1qq61qu76qqjJlypQ88MADNhsEAJAkM2bMyIc//OFar/F73/te5s6da7NBAABJ8slPfjJjx46t9RpXrFjhV/5AAADrDRgwILNmzUqfPn1qu8Y1a9Zk1KhR6ejosOEgAICk89fh6v6K3/PPPz933nmnzQYBACTJ5MmTc/DBB9d6jUuWLMmMGTNsNggAIEn222+/TJo0qdZrfOyxx7ziFwQAsF4TXvH78ssvZ8yYMVm9erUNBwEAJMmCBQuy00471XqNM2fOzM0332yzQQAASTJ27NgceeSRtV7jbbfdlrPPPttmgwAAkma84vepp57y+/4gAID1evfunXnz5mXzzTev7RpfffXVjB8/Pk888YQNBwEAJMmll16aIUOG1HqNc+fOzeLFi202CAAgSY455pgMGzas1mtctmxZTj/9dJsNAgBIksGDB2f69Onp3bt3bdf47LPPZsSIEV7xCwIAWG/RokUZMGBAbde3bt26nHXWWVmxYoXNBgEAJMn06dPzkY98pNZr/OY3v5lFixbZbBAAQJIceuihOeWUU2q9xp///Oc58cQTbTYIACDpfMXvxRdfXOtX/D7//PM54YQTsnbtWhsOAgBIOl/xu91229V6jVOnTs0999xjs0EAAEnyhS98ofav+L322mtz8cUX22wQAEDSjFf8rly5MiNGjLDZIACApPMVv5dddlne9ra31XaNL774Yk466aSsWbPGhoMAAJJk3rx5efe7313rNV500UVZunSpzQYBACTJmDFj8ulPf7rWa7z55ptz/vnn22wQAECSDBkyJOeee25KKbVd4+OPP57hw4fbbBAAQJL06tUrCxcuzBZbbFHbNb7yyis59dRTs2rVKhsOAgBIOl/xu/vuu9d6jbNnz84NN9xgs0EAAEny+c9/vvZfi99+++21/7VGEADA6zZ48ODMmDGj1q/4feaZZ/y+PwgA4E/V/RW/a9euzYQJE/LII4/YbBAAQNL5u/B1f8XvokWL8p3vfMdmgwAAkma84ve+++7L2LFjbTYIACDpfMXvrFmz0rdv39qucfXq1Rk5cmTWrVtnw0EAAEmycOHCbL/99rVdX1VVmTx5cpYvX26zQQAASXLmmWfm4x//eK3XeNVVV+Xyyy+32SAAgCTZZ599ctZZZ9V6jQ899JBf+QMBAKy36aabZu7cubV+xe+aNWsyatSodHR02HAQAEDS+XP/nXfeudZrnDZtWu666y6bDQIASJKTTz659q/4/eEPf5ivfOUrNhsEAJB0vuL3vPPOq/Urfh999FE/9wcBAPzHfzANeMXvSy+9lDFjxmT16tU2HAQAkDTjFb8zZ87Mj3/8Y5sNAgBIks997nO1f8XvrbfemqlTp9psEABA0oxX/D711FN+7g8CAPhTCxYsyMCBA2u7vldffTXjx4/PE088YbNBAABJcuGFF2afffap9RrnzJmTxYsX22wQAECSHHLIIRk3blyt17hs2bKMHz/eZoMAAJKkf//+mT17dq1f8fvss8/m+OOPt9kgAID1Fi1aVOtX/K5bty4TJ07Mww8/bLNBAABJ5yt+hw4dWus1XnnllfnGN75hs0EAAEny0Y9+tPav+F2+fHlOOukkmw0CAEg6X/F7+eWX1/oVv88//3xOOOGErF271oaDAACSzt/3r/srfr/4xS/m3nvvtdkgAIAkGT16dI466qhar3Hx4sW55JJLbDYgACBpxit+f/WrX2XkyJE2GxAAkHS+4nf+/PnZcssta7vGP/7xjzn55JOzZs0aGw4I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'
sg.theme('Reddit') 
menu_def = ['My Menu Def', ['&Configurações', '&Zerar Contador','&Sobre','&Ajuda', 'E&xit']]
tray = sg.SystemTray(menu=menu_def, data_base64=logo)
tray.Hide()

frame_layout = [
    [sg.Text('E-mail:', size=(9,1)),sg.InputText(default_text=cfg_email,size=(50,1),key='cfg1')],  
    [sg.Text('Senha:', size=(9,1)),sg.Input(default_text=cfg_senha,size=(50,1),key='cfg2')],
    [sg.Text('Servidor:', size=(9,1)),sg.Input(default_text=cfg_server,size=(50,1),key='cfg3')],  
    [sg.Text('Porta:', size=(9,1)),sg.Input(default_text=cfg_porta,size=(50,1),key='cfg4')],  
    [sg.Text('Dias Antes:', size=(9,1)),sg.Input(default_text=cfg_dias,size=(50,1),key='cfg5')],  
    [sg.Text('Tempo Min.:', size=(9,1)),sg.Input(default_text=cfg_tempo,size=(50,1),key='cfg6')],  
    [sg.Text('Tamanho XML:', size=(9,1)),sg.Input(default_text=cfg_tamanho_xml,size=(50,1),key='cfg7')],  
    [sg.Text('Pasta Temp.:', size=(9,1)),sg.Input(default_text=cfg_pasta_temp,size=(50,1),key='cfg8')],  
    [sg.Text('Pasta:', size=(9,1)),sg.Input(default_text=cfg_pasta,size=(50,1),key='cfg9')],
    [sg.Text('Limpeza Log:', size=(9,1)),sg.Input(default_text=cfg_limpeza_log,size=(50,1),key='cfg10')],  
    [sg.Text('Pasta Erro:', size=(9,1)),sg.Input(default_text=cfg_pasta_erro,size=(50,1),key='cfg11')],  
    [sg.Text('Filtros XML:', size=(9,1)),sg.Input(default_text=cfg_filtros,size=(50,1),key='cfg12')],
    [sg.Button('Salvar'),sg.Button('Processar pasta Erro')]
        ]
layout = [
          [sg.Frame('Configurações', frame_layout, font='Any 12', title_color='blue')],
         ]
window = sg.Window('Z-Notas XML').Layout(layout)
tray_visible = True
window_visible = False
window_closed = True
#----------------------------------------------------------------------------------------------------------

#-------------------------------------------------variaveis globais usadas---------------------------------
quantia_notas = 0
start = time.time_ns()
last_id = 0


#--------------------------------------grava data inicial que programa executou no log---------------------
#cfg_pasta_temp
data_exe = datetime.today()
temp_data = data_exe.strftime('%d-%b-%Y')
#log = open (os.path.join(cfg_pasta_temp, 'log.txt'),'a+')
log = open(cfg_pasta_temp + 'log.txt' , 'a')
log.write('\n'+ str(temp_data) + '\n')
#log.write(str(temp_data + '\n'))
log.close()  
#----------------------------------------------------------------------------------------------------------
 
 

 
 
 #-------------------------------------------------Logica de checagem de emails-------------------------------------------
 #--variaveis da configuracao que serao usadas neste trecho
 #cfg_email cfg_senha cfg_dias cfg_tempo cfg_server cfg_porta cfg_pasta_temp cfg_pasta 
def checa_emails():
    
    global last_id
    email_user = cfg_email.split(',')
    email_pass = cfg_senha.split(',')
    tmp_data = datetime.today()
    tmp_dia = timedelta(int(cfg_dias))
    tmp_data = (tmp_data - tmp_dia)
    date = tmp_data.strftime('%d-%b-%Y')
    for i in range (0, len(email_user)):
        #print(email_user[i])
        mail = imaplib.IMAP4_SSL(cfg_server,cfg_porta) 
        try:
            mail.login(email_user[i], email_pass[i])
        except imaplib.IMAP4.error:
            sg.popup('Informações de Login e senha do e-mail erradas')
            exit()            
        mail.select('Inbox', readonly=False)
        type, data = mail.search(None, '(SENTSINCE {date})'.format(date=date))
        mail_ids = data[0]
        id_list = mail_ids.split()
        if not id_list:
            #print('vazia')
            continue
        else:
            checagem = (email_user[i],re.sub(r"\D", "", str(id_list[-1])))
            user , id_ = checagem
            #print(user)
           # print(int(id_))
            #print(last_id)
            if(int(id_) > last_id) :
                last_id = int(id_)
                for num in data[0].split():
                    typ, response = mail.fetch(num, '(FLAGS)')
                    if ('\\Seen' in str(response)):
                        flags = 1
                        #print('vista')
                    else:
                        flags = 0
                        #print('n vista')
                    typ, data = mail.fetch(num, '(RFC822)')
                    raw_email = data[0][1] 
                    try:
                        raw_email_string = raw_email.decode('utf-8')
                    except:
                        continue
                    email_message = email.message_from_string(raw_email_string)
                    if (flags == 0):
                        mail.store(num, '-FLAGS', '\\SEEN')
                    for part in email_message.walk():
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get('Content-Disposition') is None:
                            continue
                        fileName = part.get_filename()
                        if bool(fileName): 
                            if fileName.lower().endswith('.xml'):
                                nome_original = fileName
                                filePath = os.path.join(cfg_pasta_temp, fileName) # caminho aonde vai salvar
                                retorno = testa_ja_baixadas(nome_original)
                                if retorno == False:
                                    fp = open(filePath, 'wb')
                                    fp.write(part.get_payload(decode=True))
                                    fp.close()
                                    #quantia_notas = quantia_notas+1
                                else:
                                    continue
                            if fileName.endswith('.7z'):
                                nome_original = fileName
                                filePath = os.path.join(cfg_pasta_temp, fileName)
                                retorno = testa_ja_baixadas(nome_original)
                                if retorno == False:
                                    fp = open(filePath, 'wb')
                                    fp.write(part.get_payload(decode=True))
                                    fp.close()
                                    trata_compactada(filePath)
                                else:
                                    continue
                            if fileName.endswith('.zip'):
                                nome_original = fileName
                                filePath = os.path.join(cfg_pasta_temp, fileName)
                                retorno = testa_ja_baixadas(nome_original)
                                if retorno == False:
                                    fp = open(filePath, 'wb')
                                    fp.write(part.get_payload(decode=True))
                                    fp.close()
                                    trata_compactada(filePath)
                                else:
                                    continue
                            if fileName.endswith('.rar'):
                                nome_original = fileName
                                filePath = os.path.join(cfg_pasta_temp, fileName)
                                retorno = testa_ja_baixadas(nome_original)
                                if retorno == False:
                                    fp = open(filePath, 'wb')
                                    fp.write(part.get_payload(decode=True))
                                    fp.close()
                                    trata_compactada(filePath)
                                else:
                                    continue
            else:
                #print('n entrou')
                continue
    mail.close()
    mail.logout()  
    ajuste_final()
    move_tudo()    
                 
#-------------------------------------------------------varre baixadas e baixa somente novas baseado no log-----------------    
#cfg_pasta_temp        
def testa_ja_baixadas(nome_original):
    filePath2 = os.path.join(cfg_pasta_temp, 'log.txt')
    with open(filePath2, 'r+') as file:
        ends_with_newline = True
        for line in file:
            ends_with_newline = line.endswith('\n')
            if line.rstrip('\n\r') == nome_original:
                return True
                break
        else:
            if not ends_with_newline:
                file.write('\n')
            file.write(nome_original + '\n')
            return False  
#---------------------------------------------------------------------------------------------------------------------------


#-----------------------------------TRATAMENTO DAS COMPACTADAS---------------------------------------------

def trata_compactada(filePath):
    if filePath.endswith('.7z'):
        with py7zr.SevenZipFile(filePath, 'r') as zipfile:
            zipfile.extractall(path= cfg_pasta_temp)
    if filePath.endswith('.zip'):
        with ZipFile(filePath, 'r') as zipobj:
            zipobj.extractall(cfg_pasta_temp)
    if filePath.endswith('.rar'):
        patoolib.extract_archive(filePath, outdir=cfg_pasta_temp,verbosity=-1)
    os.remove(filePath)         
            

#-----------------------------------------Limpa e edita o nome baseado na informação dentro do xml da nota------------------
#cfg_pasta_temp cfg_tamanho_xml
def ajuste_final():
    filtros = cfg_filtros.split(',')
    files = os.listdir(cfg_pasta_temp)
    for f2 in files:
        if f2 == 'log.txt':
            continue
       # if f2.lower().endswith('.xml'):
        if len(f2) != (int(cfg_tamanho_xml) + 4):
            try:
                with open(cfg_pasta_temp + f2, 'r') as f:
                    data = f.read()
            except:
               try:
                   with open(cfg_pasta_temp + f2, 'r', encoding='utf-8') as f:
                       data = f.read()
               except:
                   shutil.move(cfg_pasta_temp + f2, cfg_pasta_erro + f2)
                   continue                   
            for elementos in filtros:
                if data.find(elementos) != -1 :
                    a = data.find(elementos)
                    novo_nome2 = (data[(a + len(elementos)):(a + 44 + len(elementos))])
                    break
                else:
                    continue      
            try:
                os.rename(cfg_pasta_temp + f2, cfg_pasta_temp + novo_nome2 + '.xml')
            except:
                shutil.move(cfg_pasta_temp + f2, cfg_pasta_erro + f2)  
           
                
#-------------#------------------------------------------------------------------------------------------------
def ajuste_final_erro():
    filtros = cfg_filtros.split(',')
    files = os.listdir(cfg_pasta_erro)
    for f2 in files:
        
        if f2 == 'erro.txt':
            continue
        if len(f2) != (int(cfg_tamanho_xml) + 4):
            try:
                with open(cfg_pasta_erro+ f2, 'r') as f:
                    data = f.read()
            except:
               try:
                   with open(cfg_pasta_erro + f2, 'r', encoding='utf-8') as f:
                       data = f.read()
               except:
                   continue                   
            for elementos in filtros:
                if data.find(elementos) != -1 :
                    a = data.find(elementos)
                    novo_nome2 = (data[(a + len(elementos)):(a + 44 + len(elementos))])
                    break
                else:
                    continue      
            try:
                os.rename(cfg_pasta_erro + f2, cfg_pasta_temp + novo_nome2 + '.xml')
            except:
                k = 0    







#----------------------------------------------datas no log e limpeza do log----------------------------------
#cfg_pasta_data_temp cfg_limpeza_log

def data_log():
    data_exe = datetime.today()
    tmp_dia2 = timedelta(int(cfg_limpeza_log))
    tmp_data2 = (data_exe - tmp_dia2)
    tmp_data2 = tmp_data2.strftime('%d-%b-%Y')
    temp_data = data_exe.strftime('%d-%b-%Y')
    log = open (os.path.join(cfg_pasta_temp, 'log.txt'),'a+')
    log.write(str(temp_data + '\n'))
    log.close() 
    log2 = open(cfg_pasta_temp + 'log.txt')
    if (tmp_data2 in log2.read()):
        #print('entrou')
        log2.close()
        with open(cfg_pasta_temp + 'log.txt','r') as dados:
            for num, line in enumerate(dados, 1):
                if tmp_data2 in line:
                    break
        with open(cfg_pasta_temp + 'log.txt') as dados2, open(cfg_pasta_temp + 'log2.txt', 'w') as new:
            lines = dados2.readlines()
            new.writelines(lines[num:])        
        os.replace(cfg_pasta_temp + 'log2.txt',cfg_pasta_temp + 'log.txt')
#-------------------------------------------------------------------------------------------------------------

# ------------------------------------------------move arquivos para pasta final------------------------------
def move_tudo():
    global quantia_notas
    arquivos = os.listdir(cfg_pasta_temp)
    for arquivo in arquivos:
        if arquivo.endswith('.xml'):
            shutil.move(cfg_pasta_temp + arquivo, cfg_pasta + arquivo)    
            quantia_notas = quantia_notas+1            
#-----------------------------------------------------Laco Grafico--------------------------------------------
#cfg_tempo
while True: 
    
    tray.Update(tooltip = '{} : {}'.format ("Notas Baixadas",quantia_notas)) 
    if window_visible:
        event, values = window.Read()
        #print(event)
        if event is None:
            tray.UnHide()
            tray_visible = True
            window_visible = False
            window_closed = True
        elif event == 'Salvar':
             cfg_email = values['cfg1']
             cfg_senha = values['cfg2']
             cfg_dias = values['cfg5']
             cfg_porta = values['cfg4']
             cfg_pasta = values['cfg9']
             cfg_pasta_temp = values['cfg8']
             cfg_tamanho_xml = values['cfg7']
             cfg_tempo = (values['cfg6'])
             cfg_server = values['cfg3']
             cfg_limpeza_log = values['cfg10']
             cfg_pasta_erro = values['cfg11']
             cfg_execucao = values['cfg12']
             try:
                 if cfg_email and cfg_senha and cfg_dias and cfg_porta and cfg_pasta and cfg_pasta_temp and cfg_limpeza_log and cfg_tamanho_xml and cfg_tempo and cfg_server != ' ':
                     sg.popup('Alterações feitas execute novamente')
                     configuracoes = open('config2.txt', 'w')
                     configuracoes.write(cfg_email +"\n")
                     configuracoes.write(cfg_senha +"\n")
                     configuracoes.write(cfg_dias +"\n")
                     configuracoes.write(cfg_tempo +"\n")
                     configuracoes.write(cfg_server +"\n")
                     configuracoes.write(cfg_porta +"\n")
                     configuracoes.write(cfg_pasta_temp +"\n")
                     configuracoes.write(cfg_pasta +"\n")
                     configuracoes.write(cfg_tamanho_xml +"\n")
                     configuracoes.write(cfg_limpeza_log + "\n")
                     configuracoes.write(cfg_pasta_erro + "\n")
                     configuracoes.write(cfg_execucao + "\n")
                     configuracoes.close()
                     #os.execv(sys.executable, ['python'] + sys.argv)                      
                     break
                 else:
                     sg.popup('Preencha todas as informações para continuar') 
             except:
                 k = 0  
        elif event == 'Processar pasta Erro':
            ajuste_final_erro()
            move_tudo()            
    if tray_visible:
        menu_item = tray.Read(timeout= 300) 
        #tray.Update(tooltip = '{} : {}'.format ("Notas Baixadas",quantia_notas))
        end = time.time_ns()
        dif = (end - start)
        #print(dif)
        #print(cfg_tempo2)
        cfg_tempo2 = int(cfg_tempo) * 60000000000
        #cfg_tempo2 = int(cfg_tempo2)
        if dif > int(cfg_tempo2) :
            start = time.time_ns()
           # threading.Thread(target=checa_emails, daemon=True).start()
            checa_emails()
            ajuste_final()
            move_tudo()
            if datetime.today().strftime('%d') > data_exe.strftime('%d') :
                data_log()
        if menu_item == 'Configurações' or menu_item == sg.EVENT_SYSTEM_TRAY_ICON_DOUBLE_CLICKED:
            tray.Hide()
            tray_visible = False
            window_visible = True
            if window_closed:
                window_closed = False
                frame_layout = [
                    [sg.Text('E-mail:', size=(9,1)),sg.InputText(default_text=cfg_email,size=(50,1),key='cfg1')],  
                    [sg.Text('Senha:', size=(9,1)),sg.Input(default_text=cfg_senha,size=(50,1),key='cfg2')],
                    [sg.Text('Servidor:', size=(9,1)),sg.Input(default_text=cfg_server,size=(50,1),key='cfg3')],  
                    [sg.Text('Porta:', size=(9,1)),sg.Input(default_text=cfg_porta,size=(50,1),key='cfg4')],  
                    [sg.Text('Dias Antes:', size=(9,1)),sg.Input(default_text=cfg_dias,size=(50,1),key='cfg5')],  
                    [sg.Text('Tempo Min.:', size=(9,1)),sg.Input(default_text=cfg_tempo,size=(50,1),key='cfg6')],  
                    [sg.Text('Tamanho XML:', size=(9,1)),sg.Input(default_text=cfg_tamanho_xml,size=(50,1),key='cfg7')],  
                    [sg.Text('Pasta Temp.:', size=(9,1)),sg.Input(default_text=cfg_pasta_temp,size=(50,1),key='cfg8')],  
                    [sg.Text('Pasta:', size=(9,1)),sg.Input(default_text=cfg_pasta,size=(50,1),key='cfg9')], 
                    [sg.Text('Limpeza Log:', size=(9,1)),sg.Input(default_text=cfg_limpeza_log,size=(50,1),key='cfg10')], 
                    [sg.Text('Pasta Erro:', size=(9,1)),sg.Input(default_text=cfg_pasta_erro,size=(50,1),key='cfg11')],                     
                    [sg.Text('Filtros XML:', size=(9,1)),sg.Input(default_text=cfg_filtros,size=(50,1),key='cfg12')],
                    [sg.Button('Salvar'),sg.Button('Processar pasta Erro')]
                               ]
                layout = [
                    [sg.Frame('Configurações', frame_layout, font='Any 12', title_color='black')],
                         ]
                window = sg.Window('Z-Notas XML').Layout(layout)
            else:
                window.UnHide()
        elif menu_item == 'Ajuda':
            #subprocess.Popen(["notepad.exe","z-notas.html"],shell=False)
            webbrowser.open('z-notas.html')
        elif menu_item == 'Sobre':
            sg.popup('Sobre o Programa', 'Version 1.5 -Beta', 'Reportar problemas no endereco: duarte936@gmail.com')   
        elif menu_item == 'Zerar Contador':
            quantia_notas = 0
        else:           # some other tray command was received
            #print('Menu item %s'%menu_item)
            if menu_item == 'Exit' or menu_item == sg.EVENT_SYSTEM_TRAY_MESSAGE_CLICKED:
                break
        
 #------------------------------------------------------------------------------------------------------------------------           
 
 
 
 
