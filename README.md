```markdown
## DollarReportBot

Este projeto automatiza a captura do valor do dólar e cria um relatório em formato DOCX, que também é convertido para PDF. Utiliza o Selenium para acessar o site de conversão de moeda, PyAutoGUI para capturar uma captura de tela e `python-docx` para criar e formatar o documento.

## Funcionalidades

- Captura o valor atual do dólar em relação ao real brasileiro (BRL) do site [XE](https://www.xe.com/pt/currencyconverter/convert/?Amount=1&From=USD&To=BRL).
- Gera um relatório em formato DOCX com o valor do dólar, a data e uma captura de tela do site.
- Converte o documento DOCX para PDF usando `soffice` (LibreOffice).

## Requisitos

- Python 3.x
- ChromeDriver
- LibreOffice

Para instalar as dependências, execute:

```bash
pip install -r requirements.txt
```

## Configuração

1. **Instale o ChromeDriver**: Baixe o [ChromeDriver](https://sites.google.com/chromium.org/driver/) compatível com a sua versão do Chrome e coloque-o em um diretório acessível no PATH do seu sistema.

2. **Configuração do `soffice`**: Certifique-se de que o LibreOffice está instalado e que o caminho para o executável `soffice` está definido na variável de ambiente `SOFFICE_PATH`.

3. **Crie um arquivo `.env`**: Na raiz do projeto, crie um arquivo `.env` e adicione a linha com o caminho para o `soffice`:
ex:
   ```env
   SOFFICE_PATH = /path/to/soffice.exe
   ```

## Instalação

O instalador para este projeto está disponível na seção de [releases](https://github.com/usuario/projeto/releases). Baixe o instalador adequado para o seu sistema e siga as instruções para instalar e configurar o projeto.

## Como Usar

1. Execute o script principal:

   ```bash
   python app.py
   ```

2. O script acessará o site XE, capturará o valor do dólar, fará uma captura de tela e criará um relatório DOCX na sua área de trabalho. O relatório será convertido para PDF automaticamente.

## Estrutura do Projeto

- `app.py`: Script principal que executa o processo de automação.
- `.env`: Arquivo de configuração para variáveis de ambiente.
- `requirements.txt`: Lista de dependências do projeto.

## Logs

Os logs da execução são gravados no console e fornecem informações sobre o progresso e possíveis erros durante o processo.
