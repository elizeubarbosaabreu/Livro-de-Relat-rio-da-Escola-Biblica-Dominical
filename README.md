# Sistema de Relatório para Escola Bíblica Dominical

**Toda a honra e glória a Deus eternamente!**

## O que este sistema Faz?

Este Sistema gera várias planilhas do excel automaticamente e também gera Relatórios Mensais, Trimestrais e Anual da EBD e tem por finalidade substituir o livro de relatório físico. Basicamente o sistema irá criar um Diretório nomeado **Relatorios_EBD**, – talvez será necessário criar este diretório na mesma pasta de execução do criador de planilhas manualmente – e dentro deste diretório, irá gerar diversos diretórios organizados por ano, mês e planilhas para cada domingo. 


![Planilha](images/planilha.png)


### Sistema de diretórios criados pelo script:
![Diretórios com vários anos](images/diretorios.png)
![Diretório mês](images/formato.png)

### Tela do Software e Relatórios gerados

![printscreen](images/tela_relatorio.png)
![Relatório trimestral](images/Imagem%20colada.png)

## Download e Configuração

### Download

Baixe o repositório >>> [clicando aqui](https://github.com/elizeubarbosaabreu/Livro-de-Relatorio-da-Escola-Biblica-Dominical/archive/refs/heads/main.zip).

### Configurações

Existem mais de uma forma de utilizar este sistema. 

#### Modo usuário Comum

1.  Utilizando o executáveis *.exe (caso esteja usando Windows).

![Atenção](images/warning_24dp_E3E3E3_FILL0_wght400_GRAD0_opsz24.png) **ATENÇÃO**

Modifique o arquivo **classes.txt** com o nome das classes de sua EBD.

#### Modo expert

 Caso queira mudar alguns detalhes no sistema utilizando python, o requerimento é ter o Python instalado no sistema. Também será necessário criar um ambiente de variável (.env) e instalar os requerimentos com o comando ```pip install -r requirements.txt```.

1. Utilizando o **Jupyter Notebook** ou o Vs Code com a extensão Jupyter para rodar o arquivo [gerando_diretorios_e_planilhas.ipynb](gerando_diretorios_e_planilhas.ipynb)

2. Rodando os scripts via terminal: ```python3 gera_planilhas.py```(Gera as planilhas para preenchimento) e ```python3 gerador_relatorio_GUI.py ``` (Para gerar os Relatórios mensais, trimestrais e anuais).

#### Criação do executável

Para melhorar a usabilidade, e os executáveis não funcione, você pode criar os executáveis utilizando o *auto-py-to-exe*.
- Digite ```auto-py-to-exe ``` e siga a instrução no browser.

## Muito Obrigado e deixe uma estrela se este sistema for útil.