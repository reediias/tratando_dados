import pandas as pd
import re
from fuzzywuzzy import fuzz
import os

#caminho dos arquivos utilizando a bilioteca os para garantir que não quebre dependendo de onde abra o código
basePath = os.path.dirname(os.path.abspath(__file__))
arquivoOrigem =  os.path.join(basePath, "dados", "vagas_e_candidaturas.xlsx")
arquivoFinal = os.path.join(basePath, "dados", "inhire_template.xlsx")

df = pd.read_excel(arquivoOrigem)
template = pd.read_excel(arquivoFinal)

#colunas obrigatórias na aba job e application
colunasJob = [
    'code', 'name', 'locationCity', 'locationState', 'locationCountry',
    'reason', 'salaryMin', 'salaryMax', 'createdAt', 'updatedAt', 'status'
]
colunasApplication = [
    'jobCode', 'jobName', 'applicationId', 'name', 'email', 'phone', 'linkedInURL',
    'tags', 'talentLocationCity', 'expectedSalary', 'status', 'reasonDeclined', 
    'reasonRejected', 'createdAt', 'updatedAt', 'Triagem RH', 'Entrevista RH', 
    'Validação Área', 'Entrevista Gestor'
]

#função para tratar TODOS os campos em branco dependendo do tipo (genérica)
def tratarCampo(valor, tipo='texto'):
    if pd.isna(valor) or str(valor).strip() == '':
        return 'Não informado'
    
    valor = str(valor).strip()
    
    if tipo == 'numero':
        try:
            return float(valor.replace(',', '.'))
        except:
            return 'Não informado'
        
    elif tipo == 'data':
        try:
            return pd.to_datetime(valor).strftime('%Y-%m-%d %H:%M:%S')
        except:
            return 'Não informado'
        
    elif tipo == 'telefone':
        numeroLimpo = re.sub(r'[^\d]', '', valor)
        if numeroLimpo.startswith('55') and len(numeroLimpo) >= 12:
            return f'+{numeroLimpo}'
        elif len(numeroLimpo) == 11:  
            return f'+55{numeroLimpo[1:]}'
        elif len(numeroLimpo) == 10:  
            return f'+55{numeroLimpo}'
        else:
            return 'Não informado'
        
    else:
        return valor if valor else 'Não informado'

#função para tratar salários em min e max
def ajustarSalarios(valor):
    if isinstance(valor, str):
        valor = valor.lower().replace('r$', '').replace(' ', '').replace(',', '').strip()

        if 'entre' in valor:
            numeros = re.findall(r'\d+(?:\.\d+)?', valor)

            if len(numeros) == 2:
                return pd.Series([float(numeros[0]), float(numeros[1])])
            else:
                return pd.Series([None, None])
            
        else:
            try:
                valorFloat = float(valor)
                return pd.Series([valorFloat, valorFloat])
            except:
                return pd.Series([None, None])
            
    elif isinstance(valor, (int, float)):
        return pd.Series([valor, valor])
    
    else:
        return pd.Series([None, None])

#função para encontrar código da vaga e colocar na planilha final
def encontrarCodigoVaga(nomeVaga, dfVagas):
    nomeNormalizado = nomeVaga.strip().lower()

    #tentativa de match com nome exato
    for _, row in dfVagas.iterrows():
        nomeDb = row['Cargo'].strip().lower()
        if nomeNormalizado == nomeDb:
            return int(float(row['Código']) )   
        
    #tentativa de match por similaridade
    melhoresResultados = []
    for _, row in dfVagas.iterrows():
        nomeDb = row['Cargo'].strip().lower()
        similaridade = fuzz.token_sort_ratio(nomeNormalizado, nomeDb)
        melhoresResultados.append((similaridade, row['Código'], nomeDb))

    melhoresResultados.sort(reverse=True)
    melhorMatch = melhoresResultados[0]

    if melhorMatch[0] > 80:
        print(f"Match fuzzy: '{nomeVaga}' ≈ '{melhorMatch[2]}' (sim: {melhorMatch[0]}) → Código: {melhorMatch[1]}")
        return melhorMatch[1]

#leitura e tratamento das vagas
dfVagas = pd.read_excel(arquivoOrigem, sheet_name='Listagem de Vagas')

#carrega a planilha de vagas
for coluna in dfVagas.columns:
    if coluna == 'Valor Proposto':
        salarioAjustado = dfVagas[coluna].apply(ajustarSalarios)
        dfVagas['salaryMin'] = salarioAjustado[0]
        dfVagas['salaryMax'] = salarioAjustado[1]

    elif coluna in ['Inicio Recrut.', 'Fim Recrut.']:
        dfVagas[coluna] = dfVagas[coluna].apply(lambda x: tratarCampo(x, 'data'))

    else:
        dfVagas[coluna] = dfVagas[coluna].apply(tratarCampo)

#mapeia  os status para os nomes padrões esperados em job
statusMapeamento = {
    'encerrada': 'closed',
    'cancelada': 'cancelled',
    'stand by': 'standby'
}

dfVagas['status'] = dfVagas['Status'].str.lower().map(statusMapeamento).fillna(dfVagas['Status'])

#conversão e construção do arquivo final de vagas
def converterCodigoVaga(valor):
    try:
        return int(float(valor)) if valor != 'Não informado' else 'Não informado'
    except ValueError:
        return 'Não informado'
    
#cria o dataframe final de vagas
jobsFinal = pd.DataFrame({
    'code': dfVagas['Código'].apply(converterCodigoVaga),
    'name': dfVagas['Cargo'],
    'locationCity': 'Não informado', #como eu não vi na planilha original (candidaturas_e_vagas) as cidades das vagas, coloquei como 'Não informado'
    'locationState': dfVagas['Escritório'].str[:2],
    'locationCountry': 'Brasil',
    'reason': dfVagas['Motivo'],
    'salaryMin': dfVagas['salaryMin'],
    'salaryMax': dfVagas['salaryMax'],
    'createdAt': dfVagas['Inicio Recrut.'],
    'updatedAt': dfVagas['Fim Recrut.'],
    'status': dfVagas['status']
})

#função para garantir colunas fobrigatórias
def preencheColunasFaltantes(df, colunas):
    for col in colunas:
        if col not in df.columns:
            df[col] = 'Não informado'
    return df

#mapeia  os status para os nomes padrões esperados em application
statusApplication = {
    'reprovado': 'rejected',
    'desistente': 'declined',
    'ativo': 'active',
    'encerrada': 'closed', 
    'cancelada': 'cancelled',
    'stand by': 'standby'
}

#processamento das candidaturas
abas_candidaturas = [
    '1 - Cientista de Dados Sênior',
    '2 - Engenheiro de Software - PL',
    '3 - Cientista de Dados - Júnior',
    '4 - Engenheiro de Software - SR'
]

candidaturas = []
#lista o nome das abas que contém as candidaturas
for aba in abas_candidaturas:
    df = pd.read_excel(arquivoOrigem, sheet_name=aba)
    df.columns = [col.strip() for col in df.columns]
    
    #dicionário de mapeamento de tipos por coluna
    tiposDeColunas = {
        'Código Candidato': 'texto',
        'Nome Candidato': 'texto',
        'email': 'texto',
        'Telefone': 'telefone',
        'Linkedin': 'texto',
        'tags': 'texto',
        'Localização': 'texto',
        'Pretensão Salarial': 'numero',
        'Status do Candidato': 'texto',
        'Motivo': 'texto',
        'Data Inscrição': 'data',
        'Data Etapa': 'data'
    }
    
    #trata TODOS os campos das candidaturas em branco
    for coluna in df.columns:
        tipo = tiposDeColunas.get(coluna, 'texto')
        df[coluna] = df[coluna].apply(lambda x: tratarCampo(x, tipo))
    
    #renomear colunas para o formato final
    df = df.rename(columns={ 
        'Código Candidato': 'applicationId',
        'Nome Candidato': 'name',
        'email': 'email',
        'Telefone': 'phone',
        'Linkedin': 'linkedInURL',
        'tags': 'tags',
        'Localização': 'talentLocationCity',
        'Pretensão Salarial': 'expectedSalary',
        'Status do Candidato': 'status',
        'Motivo': 'reasonDeclined',
        'Data Inscrição': 'createdAt',
        'Data Etapa': 'updatedAt'
    })

    #trata tags como lista tipo ["estagio", "salario"]
    if 'tags' in df.columns:
        def formatarTags(valor):
            valor = tratarCampo(valor, 'texto')  # já trata vazio, None, espaço em branco
            if valor == 'Não informado':
                return valor
            tags = [tag.strip().capitalize() for tag in re.split(r'[;,/]', valor) if tag.strip()]
            if not tags:
                return 'Não informado'
            return str(tags).replace("'", '"') 

        df['tags'] = df['tags'].apply(formatarTags)

    #garante que a cidade (talentLocationCity) seja preenchida corretamente
    df['talentLocationCity'] = df['talentLocationCity'].apply(
        lambda x: 'Não informado' if pd.isna(x) or str(x).strip() in ['', '-', 'Não informado'] else x
    )

    #trata as etapas do processo
    etapas = ['Triagem RH', 'Entrevista RH', 'Validação Área', 'Entrevista Gestor']
    for etapa in etapas:
        df[etapa] = 'Não informado'

    for i, row in df.iterrows():
        nomeEtapa = row.get('Etapa', '').strip()
        dataEtapa = row.get('updatedAt', 'Não informado')
        if nomeEtapa in etapas:
            df.at[i, nomeEtapa] = dataEtapa
    
    #associa o nome da aba à vaga correspondente e encontra seu código
    nomeVaga = ' '.join(aba.split(' - ')[1:]).strip()
    df['jobName'] = nomeVaga
    df['jobCode'] = df['jobName'].apply(lambda x: encontrarCodigoVaga(x, dfVagas))

    #garante que todos os códigos de vagas na aba 'applications' sejam inteiros
    df['jobCode'] = df['jobCode'].apply(
        lambda x: int(float(x)) if str(x).replace('.', '', 1).isdigit() else 'Não informado'
    )
    
    candidaturas.append(df)

#junta todas as candidaturas
applicationFinal = pd.concat(candidaturas, ignore_index=True)
applicationFinal['status'] = applicationFinal['status'].str.lower().replace(statusApplication)

#garante que todas as colunas necessárias estejam presentes.
jobsFinal = preencheColunasFaltantes(jobsFinal, colunasJob)
applicationFinal = preencheColunasFaltantes(applicationFinal, colunasApplication)

#ajusta salários para inteiros.
for col in ['salaryMin', 'salaryMax']:
    jobsFinal[col] = jobsFinal[col].apply(
        lambda x: int(float(x)) if str(x).replace('.', '', 1).isdigit() else x
    )
applicationFinal['expectedSalary'] = applicationFinal['expectedSalary'].apply(
    lambda x: int(float(x)) if str(x).replace('.', '', 1).isdigit() else x
)

#seleciona apenas as colunas desejadas
jobsFinal = jobsFinal[colunasJob]
applicationFinal = applicationFinal[colunasApplication]

#salva no arquivo final
with pd.ExcelWriter(arquivoFinal) as writer:
    jobsFinal.to_excel(writer, sheet_name='jobs', index=False)
    applicationFinal.to_excel(writer, sheet_name='applications', index=False)

print("Processamento concluído com sucesso! Todos os campos vazios foram tratados.")