# 🚀 Gerador de Relatórios de Treinamento

Este projeto é uma aplicação de desktop desenvolvida em Python para automatizar completamente o processo de criação de relatórios de presença em treinamentos. A ferramenta consolida dados de múltiplas fontes, calcula percentuais de presença, extrai informações de NPS de arquivos PDF e gera um relatório final detalhado em Excel, além de automatizar o envio de e-mails.

## 📸 Demonstração

*(**DICA IMPORTANTE:** Grave um GIF curto ou tire screenshots da sua aplicação em funcionamento e coloque aqui. Para um projeto com interface gráfica, a parte visual é fundamental! Você pode usar ferramentas como o ScreenToGIF ou o LICEcap para gravar a tela.)*

**Exemplo de como ficaria:**

![Screenshot da tela principal da aplicação](caminho/para/sua/imagem.png)
*Tela principal da aplicação, com as abas de Configuração, Arquivos e E-mails.*

## ✨ Principais Funcionalidades

-   **Interface Gráfica Intuitiva:** Construído com `CustomTkinter` para uma aparência moderna e amigável.
-   **Múltiplos Modos de Relatório:** Suporte para treinamentos **Digitais**, **Presenciais** e **Híbridos**, cada um com sua própria lógica de processamento.
-   **Consolidação de Dados:** Combina listas de convidados com múltiplas listas de presença (seja por contagem de sessões ou por duração).
-   **Cálculo de Presença:** Calcula automaticamente o percentual de presença com base na carga horária líquida (descontando intervalos).
-   **Extração de NPS:** Lê arquivos PDF de relatórios NPS, extrai os totais de Promotores, Passivos e Detratores e calcula o score final.
-   **Relatórios Detalhados:** Gera um arquivo `.xlsx` estilizado com:
    -   Resumo do treinamento.
    -   Lista de participantes com status (Presente/Falta).
    -   Análise de divergências (convidados ausentes e presentes não convidados).
    -   Feedback consolidado do NPS.
-   **Ferramentas de Preparação:** Inclui utilitários para limpar e formatar arquivos CSV brutos do Microsoft Teams e planilhas específicas.
-   **Automação de E-mails:** Envia automaticamente o relatório gerado via Outlook para uma lista de e-mails, e também permite o envio de e-mails de pós-treinamento (agradecimento ou ausência) para os participantes.
-   **Desempenho Otimizado:** Utiliza `threading` para tarefas demoradas (como a busca de dados em planilhas) para não congelar a interface, e um sistema de `caching` para acelerar a leitura de arquivos grandes.

## 🛠️ Tecnologias Utilizadas

-   **Python 3**
-   **CustomTkinter & Tkinter:** Para a interface gráfica.
-   **Pandas:** Para manipulação e análise de dados.
-   **PDFPlumber:** Para extração de texto de arquivos PDF.
-   **OpenPyXL:** Para a criação e estilização dos relatórios em Excel.
-   **pywin32:** Para integração com o Microsoft Outlook (envio de e-mails).

## ⚙️ Como Usar

Siga os passos abaixo para configurar e executar o projeto em sua máquina local.

### Pré-requisitos

-   [Python 3.8+](https://www.python.org/downloads/) instalado.
-   [Git](https://git-scm.com/) instalado.
-   Microsoft Outlook instalado e configurado (para a funcionalidade de envio de e-mails).
-   O projeto foi desenvolvido para o ambiente Windows, devido à dependência `pywin32` para controlar o Outlook.

### Instalação

1.  **Clone o repositório:**
    ```bash
    git clone https://github.com/jooh0010/automacao-relatorios-python.git
    cd automacao-relatorios-python
    ```

2.  **Crie e ative um ambiente virtual (Recomendado):**
    ```bash
    # Cria o ambiente
    python -m venv venv

    # Ativa no Windows
    .\venv\Scripts\activate

    # Ativa no macOS/Linux
    source venv/bin/activate
    ```

3.  **Instale as dependências:**
    Crie um arquivo chamado `requirements.txt` na pasta do projeto e cole o conteúdo abaixo.

    ```txt
    pandas
    pdfplumber
    customtkinter
    tkcalendar
    openpyxl
    pywin32
    ```
    Depois, instale tudo com um único comando:
    ```bash
    pip install -r requirements.txt
    ```

### Execução

Com o ambiente virtual ativado e as dependências instaladas, execute o seguinte comando no terminal para iniciar a aplicação:

```bash
python app.py 
```
*(Se você salvou o seu arquivo principal com outro nome, como `main.py`, use esse nome no comando).*

## 📄 Estrutura dos Arquivos de Entrada

Para que o programa funcione corretamente, os arquivos de entrada devem seguir uma estrutura mínima:

-   **Lista de Convidados/Base de Profissionais (.xlsx):** Deve conter, no mínimo, colunas de `Nome` e `E-mail`.
-   **Lista de Presença Digital (.xlsx):** Deve conter colunas para `E-mail`, `Data Entrada`, `Hora Entrada`, `Data Saída` e `Hora Saída`.
-   **Lista de Presença Presencial (.xlsx, .csv):** Deve conter uma coluna com um identificador único, preferencialmente `E-mail` ou `Nome`.
-   **Relatório NPS (.pdf):** Deve ser um arquivo PDF que contenha o texto "Promoters", "Passives" e "Detractors" seguido de seus respectivos totais numéricos.

## ✒️ Autor

-   **João Rodrigues**
-   **LinkedIn:** [https://www.linkedin.com/in/joao-rodrigues-junior/](https://www.linkedin.com/in/joao-rodrigues-junior/)
-   **GitHub:** [@jooh0010](https://github.com/jooh0010)

## ⚖️ Licença

Este projeto está sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.
