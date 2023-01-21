import json
import logging
import os
import shutil
import socket
import time
import winreg
from datetime import date, datetime
from winreg import HKEY_LOCAL_MACHINE

# import dateutil
# from dateutil import relativedelta
import pandas as pd
import psutil as ps

import TOPdeskPy


# import dateutil


# COMANDO QUE DEU CERTO NO TERMINAL
# cxfreeze LimpaTempExecutaMulticlubes.py --target-dir LimpaTempExecutaKiosk --base-name=WIN32GUI --icon=icon.ico

# Lendo configs.json
def ConfigsJSON():
    try:
        with open('configs.json', 'r') as json_file:
            configs = json.load(json_file)
    except Exception as error:
        log.error(error)
    else:
        return configs

# Função que verifica quando a/o senha/token da API esta prestes a expirar o prazo de validade
# Esta informação encontra-se no arquivo configs.json e deve ser modificado diretamente no arquivo
# Uma nova data de validade deve ser inserida também no arquivo
def VerificaValidadeSenhaApi():
    dif_tempo = datetime.strptime(ConfigsJSON()['VALIDADE_SENHA_API'],
                                  '%d/%m/%Y %H:%M:%S') - datetime.strptime(datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
                                                                           '%d/%m/%Y %H:%M:%S')
    conteudo_chamado = 'A/O senha/token de API do Operador {} vence em 3 dias! Considere renova-la e atualiza-la no arquivo configs.json.'.format(
        ConfigsJSON()['USUARIO_API'])

    if dif_tempo.days <= 3:
        Abre_Chamado(Conexao_API(ConfigsJSON()), socket.gethostname(), conteudo_chamado,
                     f'{socket.gethostname()} - Token/Senha da API próxima de expirar')


# Função limpa pasta %temp% do terminal
def LimpaTemp(dir):
    for f in os.listdir(dir):
        if os.path.isfile(os.path.join(dir, f)):
            try:
                os.remove(os.path.join(dir, f))
                print('Removeu File: {}'.format(os.path.join(dir, f)))
            except Exception as error:
                print(error)
            finally:
                continue
        elif os.path.isdir(os.path.join(dir, f)):
            # print('Pasta: {}'.format(os.path.join(dir,f)))
            try:
                shutil.rmtree(os.path.join(dir, f))
                print('Removeu Pasta: {}'.format(os.path.join(dir, f)))
            except Exception as error:
                print(error)
            finally:
                continue


# Função que obtem lista de processos em execução no Windows
def Obtem_Lista_Processos():
    lista_de_processos = list()
    # print('lista de processos:')
    for process in ps.process_iter():
        info = process.as_dict(attrs=['pid', 'name'])

        lista_de_processos.append(info['name'])
    return lista_de_processos


# Função que executa a aplicação Kiosk do Multiclubes
def Executa_Multiclubes():
    # Pega appref do Multiclubes para executarmos
    PastaMulticlubes = 'C:/Users/' + os.getlogin() + '/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/Triade Soluções Inteligentes/MultiClubes/MultiClubes Autoatendimento.appref-ms'

    # Execura o appref do multiclubes
    os.startfile(PastaMulticlubes)


# Função que faz uso da bibliotéca TOPdeskPY para uso da API do TopDesk
def Conexao_API(configs):
    try:
        topdesk = TOPdeskPy.connect(configs['URL_API'], configs['USUARIO_API'], configs['SENHA_API'])
    except Exception as error:
        log.error(error)
    else:
        log.info('Conexão com API TopDesk bem sucedida')
    finally:
        return topdesk


# Função que verifica se há chamdo aberto referente ao terminal de autoatendimento
def Busca_Chamado_Aberto(hostname, topdesk):
    # Criação do dicionário para filtragem dos chamados de incidente não resolvidos
    incident_list = {}
    incident_list['status'] = 'firstLine'
    incident_list['completed'] = False
    try:
        # cria uma lista chamada result com todos os chamados retornados na busca com filtro
        result = topdesk.incident.get_list(**incident_list)
    except Exception as error:
        log.error(error)
    else:
        # Aqui é percorrida a lista contendo os chamados de incidente não resolvidos
        for i in range(0, len(result)):
            # São guardados chamado a chamado na variável cl o dicionário atralado a chave 'object'
            cl = result[i]['object']
            # Caso a variável não esteja vazia, significa que o chamado em questão possui um objeto atrelado
            if cl != None:
                # verifica se o objeto atrelado se trata do terminal em questão
                if cl['name'] == hostname:
                    print(result[i]['number'])
                    # Caso positivo: returna True e o número do chamado, indicando para a função Abre_Chamado() que
                    # não será necessário abrir um novo chamado
                    return True, result[i]['number']
    # Caso a função não termine no if acima, a mesma terminará neste return indicando que não há um chamado aberto
    return False, None


def Abre_Chamado(topdesk, hostname, conteudo_chamado, breve_descrição):
    try:
        # Criada a variável verifica para receber o retorno da função que buscou ocorrência de chamado não resolvido
        # para o terminal
        verifica = Busca_Chamado_Aberto(hostname, topdesk)
        if not verifica[0]:
            print(breve_descrição)
            # Aqui preenchemos o chamado no formato esperado pela API
            incident_parm = {
                'status': 'firstLine',
                'briefDescription': breve_descrição,
                'impact': {'name': 'Afeta uma pessoa'},
                'operatorGroup': {'id': '5eb97f94-0831-47e9-bf0c-a2f8a85e7c60'},
                'priority': {'name': 'Incidente - Baixa'},
                'urgency': {'name': 'Consigo trabalhar'},
                'entryType': {'name': 'Portal de Serviços'},
                'category': {'name': 'Terminal de Auto Atendimento'},
                'subcategory': {'name': 'Manutenção no Terminal'},
                'callType': {'name': 'Incidente'},
                'request': conteudo_chamado,
                'object': {'name': hostname,
                           'type': None,
                           'make': None,
                           'model': None,
                           'branch': None,
                           'location': None,
                           'specification': '',
                           'serialNumber': ''},

            }
            # Criação do chamado incidente
            chamado = topdesk.incident.create('servicedeskbot@aviva.com.br', **incident_parm)
            # Registra no arquivo de log o valor do chamado
            log.info(f'Foi aberto o chamado {chamado["number"]}')
        else:
            # chamado = verifica[1]
            # Registra no arquivo de log a já existência de um chamado para o terminal
            log.info(f'Já existe um chamado ({verifica[1]}) aberto para o termina {hostname}')
    except Exception as error:
        log.error(error)


# Função que busca o TeamViewer do terminal
def GetTeamViewer():
    Mensagem = ' TeamViewer não localizado!!'
    try:
        hkey = winreg.OpenKey(HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\TeamViewer", 0, access=winreg.KEY_READ)
        id_TeamViewer = winreg.QueryValueEx(hkey, 'ClientID')
    except Exception as error:
        log.error(error)
    else:
        Mensagem = str(id_TeamViewer[0])
    finally:
        return Mensagem


# Guarda registro de execução
def Guarda_Registro(hostname):
    try:
        if not os.path.exists('Registros_de_execução'):
            os.makedirs('Registros_de_execução')
            df_terminais = pd.DataFrame(columns=['HOSTNAME', 'ID_TEAMVIEWER', 'DATA_HORA_ULTIMA_EXEC'])
            df_terminais = pd.concat([df_terminais, pd.DataFrame([{'HOSTNAME': hostname,
                                                                   'ID_TEAMVIEWER': GetTeamViewer(),
                                                                   'DATA_HORA_ULTIMA_EXEC': datetime.now().strftime(
                                                                       '%d/%m/%Y %H:%M:%S')}])])
            df_terminais.drop_duplicates(subset=['HOSTNAME'], keep='last', inplace=True)
            df_terminais.reset_index(drop=True, inplace=True)
            cwd = os.getcwd()
            print(cwd)
            # df_terminais.to_csv(cwd+'\\Registros_de_execução\\Registros.csv')
            print(df_terminais)
            df_terminais.to_xml(cwd + '\\Registros_de_execução\\Registros.xml')
        else:
            cols = ['HOSTNAME', 'ID_TEAMVIEWER', 'DATA_HORA_ULTIMA_EXEC']
            cwd = os.getcwd()
            print(cwd)
            df_terminais = pd.read_xml(cwd + '\\Registros_de_execução\\Registros.xml')[cols]
            # df_terminais = pd.read_csv(cwd+'\\Registros_de_execução\\Registros.csv', index_col=0)
            df_terminais = pd.concat([df_terminais, pd.DataFrame([{'HOSTNAME': hostname,
                                                                   'ID_TEAMVIEWER': GetTeamViewer(),
                                                                   'DATA_HORA_ULTIMA_EXEC': datetime.now().strftime(
                                                                       '%d/%m/%Y %H:%M:%S')}])])
            print(df_terminais)
            df_terminais.drop_duplicates(subset=['HOSTNAME'], keep='last', inplace=True)
            df_terminais.reset_index(drop=True, inplace=True)
            print(df_terminais)
            # df_terminais.to_csv(cwd+'\\Registros_de_execução\\Registros.csv')
            df_terminais.to_xml(cwd + '\\Registros_de_execução\\Registros.xml')
    except Exception as error:
        log.error(error)


# Função que busca no arquivo de registro algum terminal que esteja a mais de 15 minutos sem executar
def Busca_Terminais_Inativo():
    try:
        # Lê arquivo .xml que guarda os valores
        # cols = [['HOSTNAME', 'ID_TEAMVIEWER', 'DATA_HORA_ULTIMA_EXEC']]
        cwd = os.getcwd()
        df_terminais = pd.read_xml(cwd + '\\Registros_de_execução\\Registros.xml')[
            ['HOSTNAME', 'ID_TEAMVIEWER', 'DATA_HORA_ULTIMA_EXEC']]

        for idx, row in df_terminais.iterrows():
            dif_tempo = datetime.strptime(datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
                                          '%d/%m/%Y %H:%M:%S') - datetime.strptime(row['DATA_HORA_ULTIMA_EXEC'],
                                                                                   '%d/%m/%Y %H:%M:%S')
            if (dif_tempo.seconds > 900) and (dif_tempo.seconds <= 172800):
                print(
                    'Abrindo chamado para o computador {} pois já se passaram 15 minutos desde a '
                    'ultima execução da tarefa'
                    .format(row['HOSTNAME']))

                # Criando variável com conteúdo do chamado
                conteudo_chamado = f'\nID do Team Viewer\r\n- {row["ID_TEAMVIEWER"]} \r\n\r\nDescreva detalhadamente seu ' \
                                   f'chamado\r\n- ' \
                                   f'Terminal desligado e/ou offline há mais de 15 minutos\r\n\r\nO terminal {row["HOSTNAME"]} ' \
                                   f'pode estar desligado ou sem conexão com a internet, pois não há registros de ' \
                                   f'execução com exito nos ultimos 15 minutos.\r\n\r '

                # Argumentos de função (parametros do arquivo configs.json, conteudo do chamado, breve descrição do chamado)
                Abre_Chamado(Conexao_API(ConfigsJSON()), row["HOSTNAME"], conteudo_chamado,
                             f'{row["HOSTNAME"]} - Terminal fora de operação')


    except Exception as error:
        log.error(error)


if __name__ == "__main__":

    # Cria uma pasta para armazenar logs de execução
    DateTodayStr = '{:02}-{:02}-{}'.format(date.today().day, date.today().month, date.today().year)
    if not os.path.exists('Logs'):
        os.makedirs('Logs')
    LOG_FORMAT = '%(levelname)s: %(asctime)s - %(message)s'
    print('~Log/log_' + socket.gethostname() + '_' + DateTodayStr + '.log')

    logging.basicConfig(filename='Logs/log_' + socket.gethostname() + '_' + DateTodayStr + '.log', level=logging.DEBUG,
                        filemode='a', format=LOG_FORMAT)
    log = logging.getLogger()

    # Chamada da função que lê o arquivo de configurações configs.json
    configs = ConfigsJSON()

    # Verifica se quem está rodando o programa é o servidor ou não
    if socket.gethostname() != configs['SERVIDOR']:
        # log.debug('Entrando no IF! Identificou que não se trata do nó servidor')
        if 'MultiClubes.Kiosk.UI.exe' not in Obtem_Lista_Processos():
            # Limpa pasta temp. Do utilizador logado
            dir = "C:/Users/" + os.getlogin() + "/AppData/Local/Temp/"
            print(dir)
            LimpaTemp(dir)

            # Executa Multiclubes
            try:
                Executa_Multiclubes()
            except Exception as error:
                log.error(error)
            else:
                time.sleep(30)
                if 'MultiClubes.Kiosk.UI.exe' not in Obtem_Lista_Processos():
                    log.info('O MULTICLUBES NÃO EXECUTOU NO COMPUTADOR ID TEAM VIEWER: ' + GetTeamViewer() + ' !')
                    # Criada variável que receberá o conteúdo do chamado que será aberto
                    conteudo_chamado = f'\nID do Team Viewer\r\n- {GetTeamViewer()} \r\n\r\nDescreva detalhadamente seu ' \
                                       f'chamado\r\n- ' \
                                       f'Kiosk com problemas de execução\r\n\r\nO Multiclubes Autoatendimento não pôde ser ' \
                                       f'executado no ' \
                                       f'autoatendimento {socket.gethostname()}.\r\n\r '
                    Abre_Chamado(Conexao_API(ConfigsJSON()), socket.gethostname(), conteudo_chamado,
                                 f'{socket.gethostname()} - Multiclubes Kiosk não abre')
                else:
                    # Armazena no log a informação de que o Multiclubes executou corretamente
                    log.info(
                        'O MULTICLUBES EXECUTOU NO COMPUTADOR ID TEAM VIEWER: ' + GetTeamViewer() + ' COM SUCESSO!')
                    Guarda_Registro(socket.gethostname())
        else:
            log.info('O MULTICLUBES já esta em execução ID TEAMVIEWER: ' + GetTeamViewer() + ' !')
            Guarda_Registro(socket.gethostname())
    # Caso identifique que o computador que executa o programa se trata do nó servidor, então ele entra nesta condição elif
    elif socket.gethostname() == configs['SERVIDOR']:
        # Chamada da função que verifica se a senha de API esta próxima de expirar
        VerificaValidadeSenhaApi()
        # Varredura feita pelo nó servidor às últimas execuções dos terminais de autoatendimento
        Busca_Terminais_Inativo()
