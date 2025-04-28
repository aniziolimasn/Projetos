# Separador de Arquivos e Renomear

Este aplicativo permite dividir um arquivo PDF em múltiplos arquivos menores e renomeá-los automaticamente com base em uma planilha Excel fornecida pelo usuário. Possui uma interface gráfica amigável desenvolvida em Tkinter.

## Funcionalidades
- Divide um PDF em vários arquivos menores.
- Renomeia os arquivos gerados conforme nomes presentes em uma planilha Excel.
- Permite escolher o número de páginas por arquivo.
- Interface gráfica simples e intuitiva.
- Barra de progresso e logs de operação.

## Como usar
1. Execute o arquivo `app.py`.
2. Selecione o arquivo PDF a ser dividido.
3. Selecione a planilha Excel com os nomes para os arquivos.
4. Informe o número de páginas por arquivo.
5. Escolha a pasta de saída.
6. Clique em "Iniciar" para processar.

## Requisitos
- Python 3.8 ou superior
- As dependências estão listadas em `requirements.txt`.

## Instalação das dependências
```bash
pip install -r requirements.txt
```

## Observações
- A primeira coluna da planilha Excel deve conter os nomes desejados para os arquivos PDF gerados.
- O aplicativo gera logs no arquivo `app.log`.

---
Desenvolvido por [Seu Nome].