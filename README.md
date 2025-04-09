# Sistema de Relatório para Escola Bíblica Dominical

## O que este sistema Faz?

Este Sistema gera um sistema de planilhas para excell que substitui o livro de relatório físico. Basicamente o sistema cria um Diretório nomeado **Relatorios_EBD** e dentro deste diretório coloca o ano (ex.: 2025), dentro deste diretório coloca os meses com as planilhas de cada domingo. 

![Planilha](images/planilha.png)

### Sistema de diretórios criados pelo script:
![Diretórios com vários anos](images/diretorios.png)
![Diretório mês](images/formato.png)
![Relatório trimestral](images/Imagem%20colada.png)

## Utilização do sistema

Existem mais de uma forma de utilizar este sistema. O requerimento é ter o Python instalado no sistema. Também será necessário criar um ambiente de variável (.env) e instalar os requerimentos com o comando ```pip install -r requirements.txt```.

1. Utilizando o **Jupyter Notebook** ou o Vs Code com a extensão Jupyter para rodar o arquivo [gerando_diretorios_e_planilhas.ipynb](gerando_diretorios_e_planilhas.ipynb)
2. Utilizando o [Google Colab](colab.research.google.com) para rodar [gerando_diretorios_e_planilhas.ipynb](gerando_diretorios_e_planilhas.ipynb).
3. Rodando os scripts [gera_planilhas](gera_planilhas.py) e [gera_relatorio_ebd](gera_relatorio_ebd.py) após python3 diretamente no terminal.

## Criação do executável

Para melhorar a usabilidade, você pode criar os executáveis para os arquivos ```*.py``` com os comando ```pyinstaller --onefile --console gera_planilhas.py``` e também ```pyinstaller --onefile --console relatorio_ebd.py```. 

## Muito Obrigado e deixe uma estrela se este sistema for útil.