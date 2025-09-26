# Robô de Automação de Chamadas com Python

![Python](https://img.shields.io/badge/python-3.8+-blue.svg)
![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)
![Status](https://img.shields.io/badge/status-conclu%C3%ADdo-brightgreen)

Um robô de automação de desktop construído com Python, Selenium e Tkinter para automatizar o processo de busca de contatos e realização de chamadas em um sistema web, com uma interface gráfica amigável e funcionalidades de priorização.

![Demonstração da Interface do Robô](https://i.imgur.com/link_para_uma_imagem_ou_gif_da_sua_gui.png)
*(Sugestão: Grave um GIF curto da aplicação funcionando e substitua o link acima)*

---

## 📖 Sobre o Projeto

[cite_start]Este projeto foi desenvolvido para otimizar o fluxo de trabalho de equipes de atendimento ou vendas que precisam realizar um grande volume de chamadas diariamente. A ferramenta automatiza as tarefas repetitivas de:
1.  Ler uma lista de clientes (CODs) de uma planilha.
2.  Buscar o telefone de cada cliente no sistema web da empresa.
3.  Discar para o número encontrado.
4.  Aguardar a ação do operador (registrar observação, agendar retorno).
5.  Gravar o resultado de cada chamada em uma planilha de controle.

[cite_start]O desenvolvimento foi uma jornada iterativa, partindo da fusão de dois scripts simples até chegar a uma aplicação robusta e empacotada como um executável (`.exe`) para fácil distribuição.

## ✨ Funcionalidades Principais

* [cite_start]**Interface Gráfica Amigável:** Painel de controle construído com Tkinter para iniciar, pausar e interagir com a automação.
* [cite_start]**Automação Web com Selenium:** Controla o Google Chrome para realizar login, buscas e discagem de forma autônoma.
* [cite_start]**Integração com Excel:** Utiliza `pandas` e `openpyxl` para ler listas de contatos e gravar resultados detalhados em abas separadas.
* [cite_start]**Fila de Prioridade Inteligente:** Permite adicionar contatos urgentes durante a execução através de um botão na interface ou de uma aba "PRIORIDADE" na planilha, que são atendidos antes da lista principal.
* [cite_start]**Validação de Dados:** Verifica se o resultado da busca corresponde ao cliente pesquisado para evitar ligar para a pessoa errada.
* [cite_start]**Controle de Contatos:** Salta automaticamente clientes que já foram contatados (presentes na aba de resultados).
* [cite_start]**Configuração Externa:** Credenciais de login e URL do sistema são gerenciadas em um arquivo `login.txt` para fácil alteração sem mexer no código.
* [cite_start]**Empacotamento para Distribuição:** O script pode ser compilado em um único arquivo `.exe` com o PyInstaller, incluindo todas as dependências como o `chromedriver`.

## 🚀 Começando

Siga estas instruções para configurar e executar o projeto no seu ambiente de desenvolvimento.

### Pré-requisitos

* Python 3.8 ou superior
* Gerenciador de pacotes `pip`
* Navegador Google Chrome instalado
* Git (opcional, para clonar o repositório)

### Instalação

1.  **Clone o repositório:**
    ```sh
    git clone [URL_DO_SEU_REPOSITORIO_GIT]
    cd [NOME_DA_PASTA_DO_PROJETO]
    ```

2.  **Crie um Ambiente Virtual (Recomendado):**
    ```sh
    python -m venv venv
    ```
    * Para ativar no Windows:
        ```sh
        .\venv\Scripts\activate
        ```

3.  **Instale as dependências:**
    O projeto utiliza um arquivo `requirements.txt` para gerenciar as bibliotecas. Execute o comando:
    ```sh
    pip install -r requirements.txt
    ```

## ⚙️ Configuração

Antes de executar, você precisa configurar dois arquivos na pasta principal do projeto:

1.  **`login.txt`:** Crie este arquivo e preencha com suas credenciais.
    ```
    Usuário=seu_usuario_weon
    Senha=sua_senha_weon
    URL=[https://suaempresa.weon.com.br](https://suaempresa.weon.com.br)
    ```

2.  [cite_start]**`automacao_weon.xlsx`:** O programa cria este arquivo automaticamente na primeira execução. Você precisa preencher as abas conforme necessário:
    * **`contatos`**: Coloque os códigos dos clientes a serem contatados na coluna `COD`.
    * **`PRIORIDADE`**: Coloque aqui os códigos de clientes urgentes. [cite_start]Eles serão lidos e a aba será limpa no início da automação.

## ▶️ Como Usar

Com a configuração pronta, execute o script principal:
```sh
python automacao_completa.py
```
A interface gráfica será aberta. Clique em "Iniciar" para começar o processo. Os botões permitem pausar, continuar, adicionar novos CODs com prioridade e registrar as ações de cada chamada.

## 📦 Gerando o Executável (.exe)

[cite_start]Para distribuir o programa sem que os usuários precisem instalar Python, você pode gerar um arquivo `.exe` usando o PyInstaller.

1.  **Instale o PyInstaller:**
    ```sh
    pip install pyinstaller
    ```

2.  **Execute o Comando de Build:**
    Navegue até a pasta do projeto via prompt de comando e execute:
    ```sh
    pyinstaller --windowed --onefile --add-data "chromedriver.exe;." automacao_completa.py
    ```
    * `--windowed`: Esconde a janela de console preta.
    * `--onefile`: Gera um único arquivo executável.
    * `--add-data`: Inclui o `chromedriver.exe` dentro do pacote.

O arquivo final estará na pasta `dist`. Lembre-se de colocar o `.exe` na mesma pasta que os arquivos `login.txt` e `automacao_weon.xlsx` para que ele funcione.

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.

---
