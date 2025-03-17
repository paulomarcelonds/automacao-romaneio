import pandas as pd
import re
from fpdf import FPDF
import warnings

# Ignorar warnings
warnings.simplefilter(action='ignore', category=UserWarning)

# Passo 1: Ler o arquivo Excel
df = pd.read_excel('marcelo-entrega-devolucao.xlsx')

# Substituir NaN na coluna 'qtd' por espaço vazio
df['qtd'] = pd.to_numeric(df['qtd'], errors='coerce').fillna('').astype(str).str.replace('.0', '')

# Converter coluna de data para formato datetime
df['data'] = pd.to_datetime(df['data'], errors='coerce')

# Passo 2: Selecionar as colunas desejadas
df = df[['sku', 'qtd', 'descricao', 'movimento', 'tecnico', 'data']]

# Agrupar os dados por movimento e técnico
grouped = df.groupby(['movimento', 'tecnico'])

# Função para criar PDF
def criar_pdf(data, movimento, tecnico, data_formatada):
    pdf = FPDF(orientation='L')  # Modo paisagem
    pdf.add_page()
    pdf.set_font('Arial', 'B', size=10)  # Fonte maior para títulos

    colunas = data.columns.tolist()
    col_widths = [40, 20, 60, 30, 40, 30]  # Largura fixa para cada coluna
    row_height = 8  # Altura fixa para cada linha

    # Adicionar cabeçalhos
    for i, col in enumerate(colunas):
        pdf.cell(col_widths[i], row_height, str(col), border=1, align='C')
    pdf.ln()

    pdf.set_font('Arial', size=8)  # Fonte menor para os dados

    # Adicionar dados
    for _, row in data.iterrows():
        for i, col in enumerate(colunas):
            valor = str(row[col])
            if col == 'data' and pd.notna(row[col]):
                valor = row[col].strftime('%d-%m-%Y')
            pdf.cell(col_widths[i], row_height, valor, border=1, align='C')
        pdf.ln()

    pdf.ln(10)
    pdf.set_font('Arial', size=10)
    mensagem = ""
    if movimento.upper() == 'ENTREGA':
        mensagem = ("Por meio desta, declaro formalmente o recebimento dos materiais "
                    "discriminados acima e assumo total responsabilidade pela guarda "
                    "e preservação dos mesmos.\n\n"
                    "Assinatura do Responsável: _________________________")
    elif movimento.upper() == 'DEVOLUCAO':
        mensagem = ("Por meio deste documento, o signatário abaixo identificado "
                    "formaliza a devolução dos materiais especificados a seguir "
                    "à parte interessada, reconhecendo o término de qualquer "
                    "responsabilidade pela sua guarda a partir da data de devolução.\n\n"
                    "Assinatura do Responsável: _________________________")
    pdf.multi_cell(0, 10, mensagem)
    pdf.ln(10)

    # Geração do nome do arquivo
    nome_tecnico = re.sub(r'\W+', '_', tecnico.lower())
    data_formatada_final = data_formatada.strftime('%d%m%Y')
    filename = f"{movimento.lower()}-{nome_tecnico}-{data_formatada_final}-normal.pdf"
    pdf.output(filename)

# Passo 3: Gerar PDFs para cada grupo
for (movimento, tecnico), group in grouped:
    if not group['data'].dropna().empty:
        data_formatada = group['data'].dropna().iloc[0]
    else:
        data_formatada = pd.Timestamp.today()
    criar_pdf(group, movimento, tecnico, data_formatada)

print("Processo concluído! Arquivos PDF gerados com sucesso.")