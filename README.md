# Tratamento das bases fechamento
[![NPM](https://img.shields.io/npm/l/react)](https://github.com/pasjunior/ETL-tratamento-da-base-fechamento-mensal/blob/main/LICENCE)

# Descrição do projeto
Este é um código Python que tem como objetivo manipular dados em um arquivo de fechamento de processos judiciais. Ele lê o arquivo, realiza algumas operações de transformação e filtragem e, em seguida, cria algumas visualizações de dados. Neste documento, será descrito o funcionamento do código e como utilizá-lo.

# Requisitos

Este código requer a instalação das seguintes bibliotecas Python:

* pandas

## Funcionamento do código
O código começa importando a biblioteca pandas e, em seguida, lendo o arquivo de fechamento de processos judiciais. Depois disso, o código realiza algumas operações de transformação nos dados, criando novas colunas e reordenando as colunas existentes.

Em seguida, o código filtra os dados em três bases diferentes: base ativa, encerrado e regulatório. O código também cria uma lista de empresas presentes no arquivo.

Por fim, o código cria uma visualização de dados em formato de tabela para as contingências contábeis e para as bases ativa, encerrada e regulatória. Essas tabelas são salvas em um arquivo de Excel.

## Funcionamento do código
Para utilizar o código, é necessário ter o arquivo de fechamento de processos judiciais em formato CSV e ter as bibliotecas pandas instaladas. Basta rodar o código no ambiente de desenvolvimento Python de sua preferência. O resultado será salvo em um arquivo Excel na pasta onde o código foi executado. Certifique-se de ajustar as variáveis do código para as suas necessidades específicas, como o nome do arquivo de entrada e as pastas utilizadas na filtragem dos dados.

# Contribuições
Contribuições para o projeto são sempre bem-vindas! Caso você queira sugerir melhorias, correções de bugs ou novas funcionalidades, por favor, abra uma issue ou pull request.
