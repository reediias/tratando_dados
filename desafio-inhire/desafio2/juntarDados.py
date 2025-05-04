import pandas as pd
import os
import re
from datetime import datetime

#função para formatar o telefone no padrão brasileiro com +55
def formatarTelefone(telefone):
    #remove tudo que não for número
    telefone = re.sub(r'\D', '', str(telefone))
    #verifica se o telefone tem mais de 1 dígito
    if len(telefone) > 1:
        return f"+55{telefone}"
    #caso não tenha informado o telefone retorna 'Não informado'
    return 'Não informado'

#função para verificar se as colunas updatedAt e createdAt contém apenas '#' ou valores inválidos e substituir por 'Não informado'
def verificarDatas(valor):
    #verifica se é string e contém #
    if isinstance(valor, str) and '#' in valor:
        return 'Não informado'
    #verifica se é NaN ou NaT
    if pd.isna(valor):
        return 'Não informado'
    #se for datetime converte para string formatada
    if isinstance(valor, (pd.Timestamp, datetime)):
        return valor.strftime('%Y-%m-%d %H:%M:%S')
    return valor

#caminho para a pasta onde os arquivos das candidaturas estão
#caminho dos arquivos utilizando a bilioteca os para garantir que não quebre dependendo de onde abra o código
basePath = os.path.dirname(os.path.abspath(__file__))
arquivoCandidaturas =  os.path.join(basePath, "candidaturas")

#lista para armazenar todos as candidaturas
armazenarCandidaturas = []

#loop para cada arquivo de candidatura do 1 ao 50
for i in range(1, 51):
    file_path = os.path.join(arquivoCandidaturas, f'candidates_{i:02d}.xlsx')
    df = pd.read_excel(file_path)
    
    #substitui valores de "#", NaN por 'Não informado'
    df = df.replace("########", 'Não informado')
    #preenche valores em branco com 'Não informado'
    df = df.fillna('Não informado')
    #ajusta a coluna de telefone para o formato correto com +55
    df['phone'] = df['phone'].apply(formatarTelefone)
    #ajusta as colunas de datas (createdAt, updatedAt) para 'Não informado' se estiver vazio
    df['createdAt'] = df['createdAt'].apply(verificarDatas)
    df['updatedAt'] = df['updatedAt'].apply(verificarDatas)
    #adiciona a coluna de origem para saber de qual arquivo cada candidatura veio
    df['Arquivo Origem'] = f'candidates_{i:02d}.xlsx'
    
    #adicionar o arquivo formatado na lista
    armazenarCandidaturas.append(df)

#junta todos os arquivo formatado em um único arquivo
arquivoConsolidado = pd.concat(armazenarCandidaturas, ignore_index=True)
#caminho onde o arquivo final será salvo
arquivoFinal = os.path.join(basePath, "arquivo_final", "dados_consolidados.xlsx")
#salva o arquivo final consolidado em um arquivo Excel
arquivoConsolidado.to_excel(arquivoFinal, index=False)

print(f'Candidaturas consolidadas salvas em: {arquivoFinal}')