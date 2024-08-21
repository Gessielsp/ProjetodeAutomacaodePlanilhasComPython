import pandas as pd
import os

# Caminhos dos arquivos
caminho_consulta = r"C:\Users\gessiel.passos\Documents\planilhas\CONSULTA LIBERAÇÕES.xlsx"
caminho_liberacoes = r"C:\Users\gessiel.passos\Documents\planilhas\LIBERAÇÕES.xlsx"

# Ler os dados da planilha CONSULTA LIBERAÇÕES
df_consulta = pd.read_excel(caminho_consulta, sheet_name='CONSULTA LIBERAÇÕES', dtype=str)

# Converter as colunas de data para o formato datetime
df_consulta['Emissão - Mês Sigla Completa (MMM/AAAA)'] = pd.to_datetime(df_consulta['Emissão - Mês Sigla Completa (MMM/AAAA)'], format='%b/%Y', errors='coerce')
df_consulta['Emissão - Dia Data Completa'] = pd.to_datetime(df_consulta['Emissão - Dia Data Completa'], errors='coerce')

# Definir a data final para agosto
data_fim_agosto = pd.Timestamp(year=2024, month=8, day=31)

# Filtrar os dados até o fim de agosto
df_filtrado = df_consulta[
    (df_consulta['Emissão - Dia Data Completa'] <= data_fim_agosto) |
    (df_consulta['Emissão - Mês Sigla Completa (MMM/AAAA)'] <= data_fim_agosto)
]

colunas_necessarias = {
    'PF - Ação Nome': 'PF - Ação Nome',
    'Emitente - UG Código': 'Emitente - UG Código',
    'Favorecido Doc. Número': 'Favorecido Doc. Número',
    'Favorecido Doc. Nome': 'Favorecido Doc. Nome',
    'PF Número': 'PF Número',
    'Emissão - Mês Sigla Completa (MMM/AAAA)': 'Emissão - Mês Sigla Completa (MMM/AAAA)',
    'Emissão - Dia Data Completa': 'Emissão - Dia Data Completa',
    'Doc - Observação Texto': 'Doc - Observação Texto',
    'PF - Recurso Código': 'PF - Recurso Código',
    'PF - Fonte Recursos Código': 'PF - Fonte Recursos Código',
    
}
df_filtrado = df_filtrado[list(colunas_necessarias.keys())]
df_filtrado = df_filtrado.rename(columns=colunas_necessarias)

if os.path.exists(caminho_liberacoes):
    df_liberacoes = pd.read_excel(caminho_liberacoes, sheet_name='Liberações')
else:
    df_liberacoes = pd.DataFrame(columns=list(colunas_necessarias.values()))

df_filtrado = df_filtrado.dropna(how='all')

df_liberacoes_atualizado = pd.concat([df_liberacoes, df_filtrado], ignore_index=True)

with pd.ExcelWriter(caminho_liberacoes, engine='openpyxl', mode='w') as writer:
    df_liberacoes_atualizado.to_excel(writer, sheet_name='Liberações', index=False)

print("Planilha 'Liberações' atualizada com sucesso.")
