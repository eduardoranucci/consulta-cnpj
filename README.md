# Consulta CNPJ

Este repositório contém uma aplicação de terminal simples desenvolvida em Python para consultar informações de CNPJ utilizando a API de consulta da [Receita WS](https://www.receitaws.com.br/). A aplicação permite aos usuários inserir um CNPJ e obter informações atualizadas sobre a empresa correspondente. Além disso, também inclui a funcionalidade de criar uma planilha com as informações consultadas usando a biblioteca Openpyxl.

## Funcionalidades

A aplicação de consulta de CNPJ inclui as seguintes funcionalidades:

- Inserir um CNPJ válido para consulta
- Consultar informações atualizadas da Receita Federal do Brasil
- Exibir na tela as informações básicas da empresa, como razão social, nome fantasia, endereço, atividade principal e secundária, entre outros dados disponíveis
- Criar uma planilha Excel (.xlsx) com as informações consultadas

## Requisitos

Para executar a aplicação em sua máquina local, você precisará ter os seguintes requisitos:

- Python 3.6 ou superior
- Biblioteca Openpyxl

## Instalação

Siga as etapas abaixo para configurar e executar a aplicação:

1. Clone este repositório para o seu ambiente local usando o seguinte comando:

   ```
   git clone https://github.com/eduardoranucci/consulta-cnpj.git
   ```

2. Navegue até o diretório raiz do projeto:

   ```
   cd consulta-cnpj
   ```

3. Opcionalmente, crie e ative um ambiente virtual:

   ```
   python3 -m venv venv
   source venv/bin/activate
   ```

4. Instale as dependências do projeto:

   ```
   pip install -r requirements.txt
   ```

5. Execute a aplicação:

   ```
   python main.py
   ```

6. Siga as instruções no terminal para inserir um CNPJ válido e obter as informações correspondentes.

7. Uma planilha Excel (.xlsx) com as informações consultadas será gerada no diretório raiz do projeto.

## Uso

Ao executar a aplicação, você terá a opção de gerar uma planilha com as informações coletadas, após isso será solicitado a inserir um CNPJ para consulta. Digite um CNPJ válido e pressione Enter. A aplicação fará uma chamada à API da Receita Federal e exibirá as informações disponíveis correspondentes ao CNPJ inserido.

Após a exibição das informações, uma planilha Excel será gerada no diretório raiz do projeto. Você poderá encontrar a planilha com o nome "consulta-cnpj.xlsx" contendo as informações consultadas.

Tenha em mente que o uso da aplicação está sujeito às políticas e limitações da API de consulta da Receita WS (3 consultas por minuto). Certifique-se de utilizá-la de acordo com as diretrizes e regulamentos aplicáveis.

## Licença

Este projeto está licenciado sob a [MIT License](LICENSE).
