import requests
import json
import openpyxl as xl

planilha = input('Gostaria de criar uma planilha com as informações consultadas? [S/N] ').strip().upper()

if planilha == 'S':
    planilha = True
elif planilha == 'N':
    planilha = False
else:
    planilha = False
    print('Resposta não identificada, planilha não será criada.')


def consultar_informacoes(cnpj_para_ser_consultado):
    page = requests.get(f'https://www.receitaws.com.br/v1/cnpj/{cnpj_para_ser_consultado}').content

    print('\nColetando informações...', end='\n\n')

    try:
        informacoes = json.loads(page)
    except json.decoder.JSONDecodeError:
        print('Limite de consulta excedido. (3/min)')
        input('Pressione ENTER para sair ')
        exit()

    return {
        'ATIVIDADE_PRINCIPAL': informacoes.get('atividade_principal'),
        'DATA_SITUACAO': informacoes.get('data_situacao'),
        'COMPLEMENTO': informacoes.get('complemento'),
        'TIPO': informacoes.get('tipo'),
        'NOME': informacoes.get('nome'),
        'UF': informacoes.get('uf'),
        'TELEFONE': informacoes.get('telefone'),
        'EMAIL': informacoes.get('email'),
        'ATIVIDADES_SECUNDARIAS': informacoes.get('atividades_secundarias'),
        'QSA': informacoes.get('qsa'),
        'SITUACAO': informacoes.get('situacao'),
        'BAIRRO': informacoes.get('bairro'),
        'LOGRADOURO': informacoes.get('logradouro'),
        'NUMERO': informacoes.get('numero'),
        'CEP': informacoes.get('cep'),
        'MUNICIPIO': informacoes.get('municipio'),
        'PORTE': informacoes.get('porte'),
        'ABERTURA': informacoes.get('abertura'),
        'NATUREZA_JURIDICA': informacoes.get('natureza_juridica'),
        'FANTASIA': informacoes.get('fantasia'),
        'CNPJ': informacoes.get('cnpj'),
        'STATUS': informacoes.get('status'),
        'EFR': informacoes.get('efr'),
        'MOTIVO_SITUACAO': informacoes.get('motivo_situacao'),
        'SITUACAO_ESPECIAL': informacoes.get('situacao_especial'),
        'DATA_SITUACAO_ESPECIAL': informacoes.get('data_situacao_especial'),
        'CAPITAL_SOCIAL': informacoes.get('capital_social')
    }


consulta = True

while consulta:
    cnpj = input('Digite um CNPJ: ').strip()
    cnpj = cnpj.replace('.', '').replace('/', '').replace('-', '')

    info = consultar_informacoes(cnpj)

    if info['STATUS'] == 'ERROR':
        print('CNPJ inválido.')
        consulta = input('Gostaria de consultar outro CNPJ? [S/N] ').strip().upper()
        if consulta == 'S':
            consulta = True
        elif consulta == 'N':
            consulta = False
            exit()
        else:
            print('Resposta não identificada.')
            input('Pressione ENTER para sair ')
            exit()

    print(f'Data de abertura: {info["ABERTURA"]}')
    print(f'Cnpj: {info["CNPJ"]}')
    print(f'Razão social: {info["NOME"]}')
    print(f'Nome fantasia: {info["FANTASIA"]}')
    print(f'Porte: {info["PORTE"]}')
    print(f'Natureza jurídica: {info["NATUREZA_JURIDICA"]}')
    print(f'Atividade Principal: {info["ATIVIDADE_PRINCIPAL"]}')
    print(f'Atividades secundarias: {info["ATIVIDADES_SECUNDARIAS"]}')
    print(f'Qsa: {info["QSA"]}')
    print(f'Telefone: {info["TELEFONE"]}')
    print(f'E-mail: {info["EMAIL"]}')
    print(f'Cep: {info["CEP"]}')
    print(f'Uf: {info["UF"]}')
    print(f'Município: {info["MUNICIPIO"]}')
    print(f'Bairro: {info["BAIRRO"]}')
    print(f'Logradouro: {info["LOGRADOURO"]}')
    print(f'Número: {info["NUMERO"]}')
    print(f'Complemento: {info["COMPLEMENTO"]}')
    print(f'Tipo: {info["TIPO"]}')
    print(f'Capital Social: {info["CAPITAL_SOCIAL"]}', end='\n\n')

    if planilha:
        try:
            df = xl.load_workbook(r'.\consulta-cnpj.xlsx')

        except FileNotFoundError:
            df = xl.Workbook()
            excel = df['Sheet']
            excel.append(['Data de abertura', 'Cnpj', 'Razão social', 'Nome fantasia', 'Porte', 'Natureza jurídica',
                          'Atividade Principal', 'Atividades secundarias', 'Qsa', 'Telefone', 'E-mail', 'Cep', 'Uf',
                          'Município', 'Bairro', 'Logradouro', 'Número', 'Complemento', 'Tipo', 'Capital Social'])
        finally:
            excel = df['Sheet']
            excel.append([str(info['ABERTURA']),
                          str(info['CNPJ']),
                          str(info['NOME']),
                          str(info['FANTASIA']),
                          str(info['PORTE']),
                          str(info['NATUREZA_JURIDICA']),
                          str(info['ATIVIDADE_PRINCIPAL']),
                          str(info['ATIVIDADES_SECUNDARIAS']),
                          str(info['QSA']),
                          str(info['TELEFONE']),
                          str(info['EMAIL']),
                          str(info['CEP']),
                          str(info['UF']),
                          str(info['MUNICIPIO']),
                          str(info['BAIRRO']),
                          str(info['LOGRADOURO']),
                          str(info['NUMERO']),
                          str(info['COMPLEMENTO']),
                          str(info['TIPO']),
                          str(info['CAPITAL_SOCIAL'])])

            df.save(r'.\consulta-cnpj.xlsx')

    consulta = input('Gostaria de consultar outro CNPJ? [S/N] ').strip().upper()
    if consulta == 'S':
        consulta = True
    elif consulta == 'N':
        consulta = False
    else:
        print('Resposta não identificada.')
        input('Pressione ENTER para sair ')
        exit()
