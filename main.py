import pandas as pd
from fpdf import FPDF

# Passo 1: Ler o arquivo Excel
df = pd.read_excel('marcelo-entrega-devolucao.xlsx')

# Substituir NaN na coluna 'qtd' por espaço vazio
df['qtd'] = df['qtd'].fillna('')

# Converter coluna de data para formato datetime
df['data'] = pd.to_datetime(df['data'])
df['qtd'] = df['qtd'].apply(lambda x: int(x) if pd.notna(x) and str(x).replace('.0', '').isdigit() else x)


# Passo 2: Selecionar as colunas desejadas
df = df[['sku', 'qtd', 'descricao', 'movimento', 'tecnico', 'data']]

# Agrupar os dados por movimento e técnico
grouped = df.groupby(['movimento', 'tecnico'])

# Função para criar PDF
def criar_pdf(data, movimento, tecnico, data_formatada):
    pdf = FPDF(orientation='L')  # Modo paisagem
    pdf.add_page()
    pdf.set_font('Arial', size=8)  # Reduzir o tamanho da fonte

    # Definir cabeçalhos e largura das colunas
    colunas = data.columns.tolist()
    
    # Calcular a largura das células dinamicamente
    col_widths = [pdf.get_string_width(col) + 10 for col in colunas]  # Largura mínima para cada coluna
    
    # Ajuste da largura das colunas com base no maior conteúdo da coluna
    for i, col in enumerate(colunas):
        max_width = max([pdf.get_string_width(str(val)) for val in data[col]])  # Maior largura entre os valores
        col_widths[i] = max(col_widths[i], max_width + 10)  # Adicionar margem

    # Adicionar cabeçalhos
    for i, col in enumerate(colunas):
        pdf.cell(col_widths[i], 10, str(col), border=1)
    pdf.ln()

    # Adicionar dados
    for _, row in data.iterrows():
        for i, col in enumerate(colunas):
            valor = str(row[col])
            
            # Formatando a data para o formato dd-mm-aaaa
            if col == 'data':
                valor = row['data'].strftime('%d-%m-%Y')
            
            # Quebrar texto em várias linhas se necessário
            if len(valor) > 30:  # Ajuste para permitir um limite maior de caracteres
                valor = valor[:30] + "..."  # Limitar o texto
            pdf.cell(col_widths[i], 10, valor, border=1)
        pdf.ln()

    # Se o movimento for "ENTREGA", adicionar a mensagem de recebimento
    if movimento.upper() == 'ENTREGA':
        pdf.ln(10)
        pdf.set_font('Arial', size=10)
        mensagem = ("Por meio desta, declaro formalmente o recebimento dos materiais "
                    "discriminados a acima e assumo integral responsabilidade pela guarda"
                    "e preservação dos mesmos:\n\n"
                    "Assinatura do Responsável: _________________________")
        pdf.multi_cell(0, 10, mensagem)
        pdf.ln(10)

 # Se o movimento for "DEVOLUCAO", adicionar a mensagem de devolução
    if movimento.upper() == 'DEVOLUCAO': 
        pdf.ln(10)
        pdf.set_font('Arial', size=10)
        mensagem = ("Por meio deste documento, o signatário abaixo identificado " 
                    "formaliza a devolução dos materiais especificados a seguir"
                    "à parte interessada, reconhecendo o término de qualquer " 
                    "responsabilidade pela sua guarda a partir da data de devolução:\n\n"
                    "Assinatura do Responsável: _________________________")
        pdf.multi_cell(0, 10, mensagem)
        pdf.ln(10)

    # Nomear e salvar arquivo
    nome_movimento = movimento.lower()
    nome_tecnico = tecnico.lower().replace(' ', '_')
    
    # Formatar a data como ddmmaaaa
    data_formatada_final = data_formatada.strftime('%d%m%Y')

    filename = f"{nome_movimento}-{nome_tecnico}-{data_formatada_final}.pdf"
    pdf.output(filename)

# Passo 3: Gerar PDFs para cada grupo
for (movimento, tecnico), group in grouped:
    # Formatar data para o nome do arquivo (ddmmaaaa)
    data_formatada = group['data'].iloc[0]

    # Criar PDF para o grupo atual
    criar_pdf(group, movimento, tecnico, data_formatada)

print("Processo concluído! Arquivos PDF gerados com sucesso.")
