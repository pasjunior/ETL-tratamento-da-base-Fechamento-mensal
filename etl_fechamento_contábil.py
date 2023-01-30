# pip install pandas
# pip install openpyxl
# pip install XlsxWriter

# Importando biblioteca
import pandas as pd


# Alterar o nome do arquivo para o arquivo do fechamento corrente

arquivo_fechamento = '(01.2023) Histórico do Fechamento Corporativo.xlsx'


# Leitura do arquivo

df_fechamento = pd.read_excel(arquivo_fechamento, sheet_name='BASE')


# Substituindo "/" por "." (também poderia utilizar o parâmetro regex=True dentro do replace)
string_substituida = df_fechamento['EMPRESA'].str.replace("/", ".")
df_empresas = pd.DataFrame(string_substituida)
df_fechamento = df_fechamento.drop('EMPRESA', axis=1)
df_fechamento.insert(loc=0, column='EMPRESA', value=df_empresas['EMPRESA'])


# Coluna calculada (ALTERAÇÃO)

df_fechamento.loc[df_fechamento['TIPO'] == 'ALTERADO', 'ALTERAÇÃO'] = df_fechamento['PROVÁVEL ATUALIZADO'] - \
                                                                      df_fechamento['VALOR ATUALIZAÇÃO PROVÁVEL'] - \
                                                                      df_fechamento['PROVÁVEL MÊS ANTERIOR']


# DF criado para alterar a ordem das colunas

df_fechamento_seq_colunas = df_fechamento[['EMPRESA', 'COMPETÊNCIA', 'NÚMERO PROCESSO', 'PASTA', 'NATUREZA', 'DATA DISTRIBUIÇÃO', 'DATA CITAÇÃO', 'DATA CADASTRO', 'PROCESSO RELEVANTE', 'COMARCA', 'CONDIÇÃO EMPRESA DO GRUPO',
                                           'POLO', 'PARTE ADVERSA', 'UC', 'TIPO AÇÃO', 'OBJETO', 'CAUSA PRIMARIA', 'DATA DO EVENTO', 'FASE', 'ADVOGADO INTERNO', 'ADVOGADO EXTERNO', 'DIRETORIA', 'DEPARTAMENTO',
                                           'PROGNÓSTICO PREDOMINANTE', 'PEDIDO ORIGINAL', 'PEDIDO MÊS ANTERIOR', 'ATUALIZAÇÃO DO MÊS', 'PEDIDO ATUALIZADO', 'PROVÁVEL INICIAL', 'PROVÁVEL MÊS ANTERIOR', 'VALOR ATUALIZAÇÃO PROVÁVEL',
                                           'PROVÁVEL ATUALIZADO', 'POSSÍVEL INICIAL', 'POSSÍVEL MÊS ANTERIOR', 'POSSÍVEL ATUALIZAÇÃO', 'POSSÍVEL ATUALIZADO', 'REMOTO INICIAL', 'REMOTO MÊS ANTERIOR', 'REMOTO ATUALIZAÇÃO',
                                           'REMOTO ATUALIZADO', 'SOMA DA DISTRIBUIÇÃO', 'JUSTIFICATIVA DA ALTERAÇÃO', 'MOTIVO DO ENCERRAMENTO', 'DATA DO ENCERRAMENTO', 'TIPO', 'ALTERAÇÃO', 'Parteterceira_Principal']]


# Filtrando as bases em (Encerrados, Base_Ativa e Regulatório)

df_base_ativa = df_fechamento_seq_colunas.loc[df_fechamento_seq_colunas['TIPO']!='ENCERRADO']
df_encerrado = df_fechamento_seq_colunas.loc[df_fechamento_seq_colunas['TIPO']=='ENCERRADO']
df_regularorio = df_fechamento_seq_colunas.loc[(df_fechamento_seq_colunas['PASTA']=='0REG.247733/2021') |
                                   (df_fechamento_seq_colunas['PASTA']=='PASTA1') |
                                   (df_fechamento_seq_colunas['PASTA']=='PASTA2') |
                                   (df_fechamento_seq_colunas['PASTA']=='PASTA3') |
                                   (df_fechamento_seq_colunas['PASTA']=='PASTA4') |
                                   (df_fechamento_seq_colunas['PASTA']=='PASTA5') |
                                   (df_fechamento_seq_colunas['PASTA']=='PASTA6') |
                                   (df_fechamento_seq_colunas['PASTA']=='PASTA7')]


# Criando lista de empresas

empresas = df_fechamento['EMPRESA'].unique().tolist()


# Criando aba CONTÁBIL

df_contabil = pd.DataFrame()
contingencias = ['SALDO INICIAL', 'NOVOS PROCESSOS', 'ALTERAÇÃO PROVISÃO (Aumento)', 'ALTERAÇÃO PROVISÃO (Diminuição)', 'ARQUIVAMENTOS', 'ATUALIZAÇÃO (+)', 'REVERSÃO DE ATUALIZAÇÃO (-)', 'SALDO FINAL'] # nome das linhas no resumo
df_contabil = pd.DataFrame(contingencias)
df_contabil.columns = ['CONTIGÊNCIAS'] # Inserindo o nome da primeira coluna

# Criando

mes = arquivo_fechamento[:10]
df_contabil_temp = pd.DataFrame()
df_base_ativa_temp = pd.DataFrame()
df_encerrado_temp = pd.DataFrame()
df_regularorio_temp = pd.DataFrame()

for emp_temp in empresas:

    # Criando os arquivos .xlsx
    pasta_saida = 'C:\Paulo\Curso\Estudos\Python\Python ETL(Fechamento)\saida\\'
    nome_arquivo_saida = pasta_saida + mes + 'FECHAMENTO CONTENCIOSO ' + emp_temp + '.xlsx'
    writer_temp = pd.ExcelWriter(nome_arquivo_saida, engine='xlsxwriter')


    # Armazendo as abas em cada arquivo
    df_base_ativa_temp = df_base_ativa.loc[df_base_ativa['EMPRESA'] == emp_temp]
    df_encerrado_temp = df_encerrado.loc[df_encerrado['EMPRESA'] == emp_temp]
    df_regularorio_temp = df_regularorio.loc[df_regularorio['EMPRESA'] == emp_temp]
    df_base_ativa_temp.to_excel(writer_temp, sheet_name='BASE ATIVA', index=False)
    df_encerrado_temp.to_excel(writer_temp, sheet_name='ENCERRADOS', index=False)
    df_regularorio_temp.to_excel(writer_temp, sheet_name='REGULATÓRIO', index=False)


    ### Criando valores da aba Contábil
    # Criando os valores da coluna TRABALHISTA
    df_fechamento_seq_colunas_empresas = df_fechamento_seq_colunas.loc[df_fechamento_seq_colunas['EMPRESA'] == emp_temp]

    trabalhista_saldo_inicial = float(df_fechamento_seq_colunas_empresas.loc[df_fechamento_seq_colunas_empresas['NATUREZA'] == 'TRABALHISTA']
                                      ['PROVÁVEL MÊS ANTERIOR'].sum())
    trabalhista_novos_processos = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['TIPO'] == 'NOVO') & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'TRABALHISTA')]
                                        ['PROVÁVEL ATUALIZADO'].sum())
    trabalhista_alt_prov_aumento = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['NATUREZA'] == 'TRABALHISTA') & (df_fechamento_seq_colunas_empresas['ALTERAÇÃO'] > 0)]
                                         ['ALTERAÇÃO'].sum())
    trabalhista_alt_prov_diminuicao = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['NATUREZA'] == 'TRABALHISTA') & (df_fechamento_seq_colunas_empresas['ALTERAÇÃO'] < 0)]
                                            ['ALTERAÇÃO'].sum())
    trabalhista_arquivamento = -float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['NATUREZA'] == 'TRABALHISTA') & (df_fechamento_seq_colunas_empresas['TIPO'] == 'ENCERRADO')]
                                      ['PROVÁVEL MÊS ANTERIOR'].sum())
    trabalhista_atualizacao = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['NATUREZA'] == 'TRABALHISTA') & (df_fechamento_seq_colunas_empresas['VALOR ATUALIZAÇÃO PROVÁVEL'] > 0) &
                                                                           (df_fechamento_seq_colunas_empresas['TIPO'] != 'ENCERRADO')]['VALOR ATUALIZAÇÃO PROVÁVEL'].sum())
    trabalhista_reversao = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['NATUREZA'] == 'TRABALHISTA') & (df_fechamento_seq_colunas_empresas['VALOR ATUALIZAÇÃO PROVÁVEL'] < 0)]
                                 ['VALOR ATUALIZAÇÃO PROVÁVEL'].sum())
    trabalhista_saldo_final = trabalhista_saldo_inicial + trabalhista_novos_processos + trabalhista_alt_prov_aumento + trabalhista_alt_prov_diminuicao + trabalhista_arquivamento + trabalhista_atualizacao + trabalhista_reversao

    trabalhista = ["{:_.2f}".format(trabalhista_saldo_inicial).replace(".",",").replace("_","."),
                   "{:_.2f}".format(trabalhista_novos_processos).replace(".",",").replace("_","."),
                   "{:_.2f}".format(trabalhista_alt_prov_aumento).replace(".",",").replace("_","."),
                   "{:_.2f}".format(trabalhista_alt_prov_diminuicao).replace(".",",").replace("_","."),
                   "{:_.2f}".format(trabalhista_arquivamento).replace(".",",").replace("_","."),
                   "{:_.2f}".format(trabalhista_atualizacao).replace(".",",").replace("_","."),
                   "{:_.2f}".format(trabalhista_reversao).replace(".",",").replace("_","."),
                   "{:_.2f}".format(trabalhista_saldo_final).replace(".",",").replace("_",".")]

    # Criando os valores da coluna CÍVEL
    civel_saldo_inicial = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['NATUREZA'] == 'ADMINISTRATIVA') |
                                                                       (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'CIVEL') |
                                                                       (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'CRIMINAL') |
                                                                       (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'JEC') |
                                                                       (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'RECEBIVEIS') |
                                                                       (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'AMBIENTAL')]
                                                                       ['PROVÁVEL MÊS ANTERIOR'].sum())
    civel_novos_processos = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['TIPO'] == 'NOVO') & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'ADMINISTRATIVA') |
                                                                         (df_fechamento_seq_colunas_empresas['TIPO'] == 'NOVO') & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'CIVEL') |
                                                                         (df_fechamento_seq_colunas_empresas['TIPO'] == 'NOVO') & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'CRIMINAL') |
                                                                         (df_fechamento_seq_colunas_empresas['TIPO'] == 'NOVO') & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'JEC') |
                                                                         (df_fechamento_seq_colunas_empresas['TIPO'] == 'NOVO') & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'RECEBIVEIS') |
                                                                         (df_fechamento_seq_colunas_empresas['TIPO'] == 'NOVO') & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'AMBIENTAL')]
                                                                         ['PROVÁVEL ATUALIZADO'].sum())
    civel_alt_prov_aumento = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['ALTERAÇÃO'] > 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'ADMINISTRATIVA') |
                                                                          (df_fechamento_seq_colunas_empresas['ALTERAÇÃO'] > 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'CIVEL') |
                                                                          (df_fechamento_seq_colunas_empresas['ALTERAÇÃO'] > 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'CRIMINAL') |
                                                                          (df_fechamento_seq_colunas_empresas['ALTERAÇÃO'] > 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'JEC') |
                                                                          (df_fechamento_seq_colunas_empresas['ALTERAÇÃO'] > 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'RECEBIVEIS') |
                                                                          (df_fechamento_seq_colunas_empresas['ALTERAÇÃO'] > 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'AMBIENTAL')]
                                                                          ['ALTERAÇÃO'].sum())
    civel_alt_prov_diminuicao = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['ALTERAÇÃO'] < 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'ADMINISTRATIVA') |
                                                                             (df_fechamento_seq_colunas_empresas['ALTERAÇÃO'] < 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'CIVEL') |
                                                                             (df_fechamento_seq_colunas_empresas['ALTERAÇÃO'] < 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'CRIMINAL') |
                                                                             (df_fechamento_seq_colunas_empresas['ALTERAÇÃO'] < 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'JEC') |
                                                                             (df_fechamento_seq_colunas_empresas['ALTERAÇÃO'] < 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'RECEBIVEIS') |
                                                                             (df_fechamento_seq_colunas_empresas['ALTERAÇÃO'] < 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'AMBIENTAL')]
                                                                             ['ALTERAÇÃO'].sum())
    civel_arquivamento = -float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['TIPO'] == 'ENCERRADO') & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'ADMINISTRATIVA') |
                                                                      (df_fechamento_seq_colunas_empresas['TIPO'] == 'ENCERRADO') & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'CIVEL') |
                                                                      (df_fechamento_seq_colunas_empresas['TIPO'] == 'ENCERRADO') & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'CRIMINAL') |
                                                                      (df_fechamento_seq_colunas_empresas['TIPO'] == 'ENCERRADO') & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'JEC') |
                                                                      (df_fechamento_seq_colunas_empresas['TIPO'] == 'ENCERRADO') & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'RECEBIVEIS') |
                                                                      (df_fechamento_seq_colunas_empresas['TIPO'] == 'ENCERRADO') & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'AMBIENTAL')]
                                                                      ['PROVÁVEL MÊS ANTERIOR'].sum())
    civel_atualizacao = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['VALOR ATUALIZAÇÃO PROVÁVEL'] > 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'ADMINISTRATIVA') &
                                                                     (df_fechamento_seq_colunas_empresas['TIPO'] != 'ENCERRADO') |
                                                                     (df_fechamento_seq_colunas_empresas['VALOR ATUALIZAÇÃO PROVÁVEL'] > 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'CIVEL') &
                                                                     (df_fechamento_seq_colunas_empresas['TIPO'] != 'ENCERRADO') |
                                                                     (df_fechamento_seq_colunas_empresas['VALOR ATUALIZAÇÃO PROVÁVEL'] > 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'CRIMINAL') &
                                                                     (df_fechamento_seq_colunas_empresas['TIPO'] != 'ENCERRADO') |
                                                                     (df_fechamento_seq_colunas_empresas['VALOR ATUALIZAÇÃO PROVÁVEL'] > 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'JEC') &
                                                                     (df_fechamento_seq_colunas_empresas['TIPO'] != 'ENCERRADO') |
                                                                     (df_fechamento_seq_colunas_empresas['VALOR ATUALIZAÇÃO PROVÁVEL'] > 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'RECEBIVEIS') &
                                                                     (df_fechamento_seq_colunas_empresas['TIPO'] != 'ENCERRADO') |
                                                                     (df_fechamento_seq_colunas_empresas['VALOR ATUALIZAÇÃO PROVÁVEL'] > 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'AMBIENTAL') &
                                                                     (df_fechamento_seq_colunas_empresas['TIPO'] != 'ENCERRADO')]
                                                                     ['VALOR ATUALIZAÇÃO PROVÁVEL'].sum())
    civel_reversao = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['VALOR ATUALIZAÇÃO PROVÁVEL'] < 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'ADMINISTRATIVA') |
                                                                  (df_fechamento_seq_colunas_empresas['VALOR ATUALIZAÇÃO PROVÁVEL'] < 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'CIVEL') |
                                                                  (df_fechamento_seq_colunas_empresas['VALOR ATUALIZAÇÃO PROVÁVEL'] < 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'CRIMINAL') |
                                                                  (df_fechamento_seq_colunas_empresas['VALOR ATUALIZAÇÃO PROVÁVEL'] < 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'JEC') |
                                                                  (df_fechamento_seq_colunas_empresas['VALOR ATUALIZAÇÃO PROVÁVEL'] < 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'RECEBIVEIS') |
                                                                  (df_fechamento_seq_colunas_empresas['VALOR ATUALIZAÇÃO PROVÁVEL'] < 0) & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'AMBIENTAL')]
                                                                  ['VALOR ATUALIZAÇÃO PROVÁVEL'].sum())
    civil_saldo_final = civel_saldo_inicial + civel_novos_processos + civel_alt_prov_aumento + civel_alt_prov_diminuicao + civel_arquivamento + civel_atualizacao + civel_reversao

    civel = ["{:_.2f}".format(civel_saldo_inicial).replace(".",",").replace("_","."),
             "{:_.2f}".format(civel_novos_processos).replace(".",",").replace("_","."),
             "{:_.2f}".format(civel_alt_prov_aumento).replace(".",",").replace("_","."),
             "{:_.2f}".format(civel_alt_prov_diminuicao).replace(".",",").replace("_","."),
             "{:_.2f}".format(civel_arquivamento).replace(".",",").replace("_","."),
             "{:_.2f}".format(civel_atualizacao).replace(".",",").replace("_","."),
             "{:_.2f}".format(civel_reversao).replace(".",",").replace("_","."),
             "{:_.2f}".format(civil_saldo_final).replace(".",",").replace("_",".")]

    # Criando os valores da coluna FISCAL
    fiscal_saldo_inicial = float(df_fechamento_seq_colunas_empresas.loc[df_fechamento_seq_colunas_empresas['NATUREZA'] == 'FISCAL']
                                 ['PROVÁVEL MÊS ANTERIOR'].sum())
    fiscal_novos_processos = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['TIPO'] == 'NOVO') & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'FISCAL')]
                                   ['PROVÁVEL ATUALIZADO'].sum())
    fiscal_alt_prov_aumento = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['NATUREZA'] == 'FISCAL') & (df_fechamento_seq_colunas_empresas['ALTERAÇÃO'] > 0)]
                                    ['ALTERAÇÃO'].sum())
    fiscal_alt_prov_diminuicao = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['NATUREZA'] == 'FISCAL') & (df_fechamento_seq_colunas_empresas['ALTERAÇÃO'] < 0)]
                                       ['ALTERAÇÃO'].sum())
    fiscal_arquivamento = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['NATUREZA'] == 'FISCAL') & (df_fechamento_seq_colunas_empresas['TIPO'] == 'ENCERRADO')]
                                ['PROVÁVEL MÊS ANTERIOR'].sum())
    fiscal_atualizacao = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['NATUREZA'] == 'FISCAL') & (df_fechamento_seq_colunas_empresas['VALOR ATUALIZAÇÃO PROVÁVEL'] > 0) &
                                                                      (df_fechamento_seq_colunas_empresas['TIPO'] != 'ENCERRADO')]['VALOR ATUALIZAÇÃO PROVÁVEL'].sum())
    fiscal_reversao = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['NATUREZA'] == 'FISCAL') & (df_fechamento_seq_colunas_empresas['VALOR ATUALIZAÇÃO PROVÁVEL'] < 0)]
                            ['VALOR ATUALIZAÇÃO PROVÁVEL'].sum())
    fiscal_saldo_final = fiscal_saldo_inicial + fiscal_novos_processos + fiscal_alt_prov_aumento + fiscal_alt_prov_diminuicao + fiscal_arquivamento + fiscal_atualizacao + fiscal_reversao

    fiscal = ["{:_.2f}".format(fiscal_saldo_inicial).replace(".",",").replace("_","."),
              "{:_.2f}".format(fiscal_novos_processos).replace(".",",").replace("_","."),
              "{:_.2f}".format(fiscal_alt_prov_aumento).replace(".",",").replace("_","."),
              "{:_.2f}".format(fiscal_alt_prov_diminuicao).replace(".",",").replace("_","."),
              "{:_.2f}".format(fiscal_arquivamento).replace(".",",").replace("_","."),
              "{:_.2f}".format(fiscal_atualizacao).replace(".",",").replace("_","."),
              "{:_.2f}".format(fiscal_reversao).replace(".",",").replace("_","."),
              "{:_.2f}".format(fiscal_saldo_final).replace(".",",").replace("_",".")]

    # Criando os valores da coluna REGULATÓRIO

    regulatorio_saldo_inicial = float(df_fechamento_seq_colunas_empresas.loc[df_fechamento_seq_colunas_empresas['NATUREZA'] == 'REGULATORIO']
                                      ['PROVÁVEL MÊS ANTERIOR'].sum())
    regulatorio_novos_processos = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['TIPO'] == 'NOVO') & (df_fechamento_seq_colunas_empresas['NATUREZA'] == 'REGULATORIO')]
                                        ['PROVÁVEL ATUALIZADO'].sum())
    regulatorio_alt_prov_aumento = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['NATUREZA'] == 'REGULATORIO') & (df_fechamento_seq_colunas_empresas['ALTERAÇÃO'] > 0)]
                                         ['ALTERAÇÃO'].sum())
    regulatorio_alt_prov_diminuicao = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['NATUREZA'] == 'REGULATORIO') & (df_fechamento_seq_colunas_empresas['ALTERAÇÃO'] < 0)]
                                            ['ALTERAÇÃO'].sum())
    regulatorio_arquivamento = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['NATUREZA'] == 'REGULATORIO') & (df_fechamento_seq_colunas_empresas['TIPO'] == 'ENCERRADO')]
                                     ['PROVÁVEL MÊS ANTERIOR'].sum())
    regulatorio_atualizacao = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['NATUREZA'] == 'REGULATORIO') & (df_fechamento_seq_colunas_empresas['VALOR ATUALIZAÇÃO PROVÁVEL'] > 0) &
                                                                           (df_fechamento_seq_colunas_empresas['TIPO'] != 'ENCERRADO')]['VALOR ATUALIZAÇÃO PROVÁVEL'].sum())
    regulatorio_reversao = float(df_fechamento_seq_colunas_empresas.loc[(df_fechamento_seq_colunas_empresas['NATUREZA'] == 'REGULATORIO') & (df_fechamento_seq_colunas_empresas['VALOR ATUALIZAÇÃO PROVÁVEL'] < 0)]
                                 ['VALOR ATUALIZAÇÃO PROVÁVEL'].sum())
    regulatorio_saldo_final = regulatorio_saldo_inicial + regulatorio_novos_processos + regulatorio_alt_prov_aumento + regulatorio_alt_prov_diminuicao + regulatorio_arquivamento + regulatorio_atualizacao + regulatorio_reversao

    regulatorio = ["{:_.2f}".format(regulatorio_saldo_inicial).replace(".",",").replace("_","."),
                   "{:_.2f}".format(regulatorio_novos_processos).replace(".",",").replace("_","."),
                   "{:_.2f}".format(regulatorio_alt_prov_aumento).replace(".",",").replace("_","."),
                   "{:_.2f}".format(regulatorio_alt_prov_diminuicao).replace(".",",").replace("_","."),
                   "{:_.2f}".format(regulatorio_arquivamento).replace(".",",").replace("_","."),
                   "{:_.2f}".format(regulatorio_atualizacao).replace(".",",").replace("_","."),
                   "{:_.2f}".format(regulatorio_reversao).replace(".",",").replace("_","."),
                   "{:_.2f}".format(regulatorio_saldo_final).replace(".",",").replace("_",".")]

    # Criando os valores da coluna SALDO
    saldo_saldo_inicial = float(trabalhista_saldo_inicial + civel_saldo_inicial + fiscal_saldo_inicial + regulatorio_saldo_inicial)
    saldo_novos_processos = float(trabalhista_novos_processos + civel_novos_processos + fiscal_novos_processos + regulatorio_novos_processos)
    saldo_alt_prov_aumento = float(trabalhista_alt_prov_aumento + civel_alt_prov_aumento + fiscal_alt_prov_aumento + regulatorio_alt_prov_aumento)
    saldo_alt_prov_diminuicao = float(trabalhista_alt_prov_diminuicao + civel_alt_prov_diminuicao + fiscal_alt_prov_diminuicao + regulatorio_alt_prov_diminuicao)
    saldo_arquivamento = float(trabalhista_arquivamento + civel_arquivamento + fiscal_arquivamento + regulatorio_arquivamento)
    saldo_atualizacao = float(trabalhista_atualizacao + civel_atualizacao + fiscal_atualizacao + regulatorio_atualizacao)
    saldo_reversao = float(trabalhista_reversao + civel_reversao + fiscal_reversao + regulatorio_reversao)
    saldo_saldo_final = saldo_saldo_inicial + saldo_novos_processos + saldo_alt_prov_aumento + saldo_alt_prov_diminuicao + saldo_arquivamento + saldo_atualizacao + saldo_reversao

    saldo = ["{:_.2f}".format(saldo_saldo_inicial).replace(".",",").replace("_","."),
             "{:_.2f}".format(saldo_novos_processos).replace(".",",").replace("_","."),
             "{:_.2f}".format(saldo_alt_prov_aumento).replace(".",",").replace("_","."),
             "{:_.2f}".format(saldo_alt_prov_diminuicao).replace(".",",").replace("_","."),
             "{:_.2f}".format(saldo_arquivamento).replace(".",",").replace("_","."),
             "{:_.2f}".format(saldo_atualizacao).replace(".",",").replace("_","."),
             "{:_.2f}".format(saldo_reversao).replace(".",",").replace("_","."),
             "{:_.2f}".format(saldo_saldo_final).replace(".",",").replace("_",".")]

    # Criando a tabela CONTÁBIL
    df_contabil_temp = df_contabil.assign(TRABALHISTA = trabalhista,
                                     CÍVEL = civel,
                                     FISCAL = fiscal,
                                     REGULATÓRIO = regulatorio,
                                     SALDO = saldo)

    df_contabil_temp.to_excel(writer_temp, sheet_name='CONTÁBIL', index=False)
    writer_temp.save()
