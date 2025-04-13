# Sistema de Relatório para Escola Bíblica Dominical

**Toda a honra e glória a Deus eternamente!**

## O que este sistema Faz?

Este Sistema gera várias planilhas do excel automaticamente e também gera Relatórios Mensais, Trimestrais e Anual da EBD e tem por finalidade substituir o livro de relatório físico em igrejas pequenas e que não queiram pagar por sistemas existentes no mercado. Basicamente o sistema irá criar um Diretório nomeado **Relatorios_EBD**, – talvez será necessário criar este diretório na mesma pasta de execução do criador de planilhas manualmente – e dentro deste diretório, irá gerar diversos diretórios organizados por ano, mês e planilhas para cada domingo. 
Agora também gera o PDF para impressão do domingo atual.


![Planilha](images/planilha.png)
![PDF exemplo](images/relatório%20pdf.png)


### Sistema de diretórios criados pelo script:
![Diretórios com vários anos](images/diretorios.png)
![Diretório mês](images/formato.png)

### Tela do Software e Relatórios gerados

![printscreen](images/tela_relatorio.png)
![Relatório trimestral](images/Imagem%20colada.png)

## Download e Configuração

#### Linux Ubuntu

- Faça o download do arquivo **Ubuntu.zip** [neste link](https://github.com/elizeubarbosaabreu/Livro-de-Relatorio-da-Escola-Biblica-Dominical/blob/c4d6651f5e519288902356daa3d516c45b662263/dist/Ubuntu.zip), descompacte e siga as orientações do arquivo LEIAME.TXT para proceder a instalação e configuração do sistema.

#### Windows XP, 7, 10 e 11

- Faça o download do arquivo **Windows.zip** [neste link](https://github.com/elizeubarbosaabreu/Livro-de-Relatorio-da-Escola-Biblica-Dominical/blob/c4d6651f5e519288902356daa3d516c45b662263/dist/Windows.zip), descompacte no diretório do compuatador que irá utilizar para guardar as planilhas e utilize os executáveis como aplicativos portáteis. Inclusive pode usar em um pendrive ou HD externo.

![Atenção](images/warning_24dp_E3E3E3_FILL0_wght400_GRAD0_opsz24.png) **ATENÇÃO**

- Lembre-se de modificar o arquivo **classes.txt** com o nome das classes (salas de aula) de sua EBD e o arquivo **igreja.txt** com o nome da sua igreja.

#### Modo expert

- Baixe o repositório >>> [clicando aqui](https://github.com/elizeubarbosaabreu/Livro-de-Relatorio-da-Escola-Biblica-Dominical/archive/refs/heads/main.zip).


 - Caso queira mudar alguns detalhes no sistema utilizando python, o requerimento é ter o Python instalado no sistema. Também será necessário criar um ambiente de variável (.env) e instalar os requerimentos com o comando ```pip install -r requirements.txt```.

- Execute os scripts via terminal: ```python3 gera_planilhas.py```(Gera as planilhas para preenchimento), ```python3 domingo_atual.py```(Gerar pdf do relatório do domingo atual para impressão ) e ```python3 gerador_relatorio_GUI.py ``` (Para gerar os Relatórios mensais, trimestrais e anuais).

#### Criação do executável

Para melhorar a usabilidade, e caso os executáveis não funcionem, você pode criar os executáveis utilizando o *auto-py-to-exe*.
- Digite ```auto-py-to-exe ``` e siga a instrução no browser.

#### Muito Obrigado e que Deus abençoe a todos!