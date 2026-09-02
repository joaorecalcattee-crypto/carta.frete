import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Auditoria: Diário vs Relatório CF", layout="wide")

# ==========================================
# FUNÇÕES DE TRATAMENTO
# ==========================================

def limpar_valor_excel(valor_str):
    """Converte valores das planilhas para número (float) de forma segura."""
    if pd.isna(valor_str) or valor_str == '':
        return 0.0
    if isinstance(valor_str, (int, float)):
        return float(valor_str)
    
    valor_str = str(valor_str).strip()
    try:
        # Se o formato já estiver com ponto como decimal (ex: 1500.50)
        if '.' in valor_str and ',' not in valor_str:
            return float(valor_str)
        # Se for formato brasileiro (ex: 1.500,50)
        valor_str = valor_str.replace('.', '').replace(',', '.')
        return float(valor_str)
    except:
        return 0.0

def processar_livro_diario(arquivo):
    """Lê o Livro Diário e extrai os lançamentos baseados no Histórico."""
    try:
        if arquivo.name.endswith('.csv'):
            df = pd.read_csv(arquivo, sep=None, engine='python')
        else:
            df = pd.read_excel(arquivo)
    except Exception as e:
        st.error(f"Erro ao ler Livro Diário: {e}")
        return pd.DataFrame()

    # Identificar colunas do Livro Diário
    col_hist = next((col for col in df.columns if 'Hist' in str(col)), None)
    col_debito = next((col for col in df.columns if 'Debito' in str(col) or 'Débito' in str(col)), None)
    col_credito = next((col for col in df.columns if 'Credito' in str(col) or 'Crédito' in str(col)), None)
    col_conta = next((col for col in df.columns if 'Conta' in str(col)), None)

    if not col_hist or not col_debito or not col_credito or not col_conta:
        st.error("Erro no Livro Diário: Não achei as colunas Histórico, Débito, Crédito ou Conta.")
        return pd.DataFrame()

    lancamentos_agrupados = {}

    for index, row in df.iterrows():
        historico = str(row[col_hist])
        # Puxa o número da CF que está no histórico
        match_cf = re.search(r'PAGO\s*C[FE]\s*(\d+)(?:\s*-\s*(.+))?', historico, re.IGNORECASE)
        
        if match_cf:
            numero_titulo = match_cf.group(1)
            prestador_excel = match_cf.group(2).strip() if match_cf.group(2) else ""
            conta = str(row[col_conta]).replace('.', '')
            
            if numero_titulo not in lancamentos_agrupados:
                lancamentos_agrupados[numero_titulo] = {
                    'Valor (Diário)': 0.0,
                    'tem_credito_banco': False,
                    'Prestador (Diário)': prestador_excel
                }
            
            if not lancamentos_agrupados[numero_titulo]['Prestador (Diário)'] and prestador_excel:
                lancamentos_agrupados[numero_titulo]['Prestador (Diário)'] = prestador_excel
            
            # Conta de Despesa de Frete (Ajuste conforme seu plano de contas)
            if conta.startswith('211010300001'):
                valor_deb = limpar_valor_excel(row[col_debito])
                if valor_deb > 0:
                    lancamentos_agrupados[numero_titulo]['Valor (Diário)'] += valor_deb
                    
            # Conta de Banco
            if conta.startswith('11102'):
                valor_cred = limpar_valor_excel(row[col_credito])
                if valor_cred > 0:
                    lancamentos_agrupados[numero_titulo]['tem_credito_banco'] = True

    lancamentos_finais = []
    for titulo, dados in lancamentos_agrupados.items():
        lancamentos_finais.append({
            'Número do Título': titulo,
            'Prestador (Diário)': dados['Prestador (Diário)'],
            'Valor (Diário)': dados['Valor (Diário)'],
            'Estrutura OK': dados['tem_credito_banco'] and (dados['Valor (Diário)'] > 0)
        })
            
    return pd.DataFrame(lancamentos_finais)

def processar_relatorio_cf(arquivo):
    """Lê o relatório das Cartas Frete e extrai número, prestador e valor."""
    try:
        if arquivo.name.endswith('.csv'):
            df = pd.read_csv(arquivo, sep=None, engine='python')
        else:
            df = pd.read_excel(arquivo)
    except Exception as e:
        st.error(f"Erro ao ler Relatório CF: {e}")
        return pd.DataFrame()

    # Tenta achar a coluna que contém o Número da Carta Frete
    col_num = next((c for c in df.columns if any(x in str(c).lower() for x in ['número', 'numero', 'título', 'titulo', 'documento', 'cf'])), None)
    # Tenta achar a coluna de Valor
    col_valor = next((c for c in df.columns if any(x in str(c).lower() for x in ['valor', 'líquido', 'liquido', 'saldo', 'receber', 'total'])), None)
    # Tenta achar a coluna de Prestador/Motorista
    col_prestador = next((c for c in df.columns if any(x in str(c).lower() for x in ['prestador', 'motorista', 'nome', 'contratado', 'favorecido'])), None)

    if not col_num or not col_valor:
        st.error("Erro no Relatório CF: Não consegui identificar as colunas de 'Número' ou 'Valor'. Verifique o cabeçalho da planilha.")
        st.write("Colunas encontradas no arquivo:", list(df.columns))
        return pd.DataFrame()

    # Extrai apenas os números da coluna de Título (caso tenha letras junto)
    df['Número do Título'] = df[col_num].astype(str).str.extract(r'(\d+)')
    df['Valor (Relatório CF)'] = df[col_valor].apply(limpar_valor_excel)
    df['Prestador (Relatório CF)'] = df[col_prestador].astype(str) if col_prestador else "N/A"

    # Remove linhas vazias
    df = df.dropna(subset=['Número do Título'])
    
    # Agrupa caso o mesmo título apareça em mais de uma linha no relatório (soma os valores)
    df_agrupado = df.groupby('Número do Título').agg({
        'Valor (Relatório CF)': 'sum',
        'Prestador (Relatório CF)': 'first'
    }).reset_index()

    return df_agrupado

# ==========================================
# INTERFACE DO USUÁRIO
# ==========================================

st.title("📊 Comparador: Livro Diário vs Relatório de Carta Frete")
st.write("Cruza os dados de dois relatórios Excel/CSV para encontrar divergências de valores ou lançamentos faltantes.")

col1, col2 = st.columns(2)
with col1:
    arquivo_diario = st.file_uploader("1️⃣ Upload do Livro Diário (Excel/CSV)", type=['xlsx', 'xls', 'csv'])
with col2:
    arquivo_relatorio_cf = st.file_uploader("2️⃣ Upload do Relatório Carta Frete (Excel/CSV)", type=['xlsx', 'xls', 'csv'])

if arquivo_diario and arquivo_relatorio_cf:
    if st.button("Iniciar Comparação", type="primary"):
        with st.spinner("Processando planilhas e cruzando dados..."):
            
            df_diario = processar_livro_diario(arquivo_diario)
            df_cf = processar_relatorio_cf(arquivo_relatorio_cf)
            
            if df_diario.empty:
                st.warning("O processamento do Livro Diário não retornou resultados válidos.")
            elif df_cf.empty:
                st.warning("O processamento do Relatório de Carta Frete não retornou resultados válidos.")
            else:
                # Cruza as duas tabelas usando o Número do Título
                df_cruzamento = pd.merge(df_cf, df_diario, on='Número do Título', how='outer')
                
                resultados = []
                for index, row in df_cruzamento.iterrows():
                    titulo = row['Número do Título']
                    
                    valor_cf = row.get('Valor (Relatório CF)', 0.0)
                    valor_diario = row.get('Valor (Diário)', 0.0)
                    if pd.isna(valor_cf): valor_cf = 0.0
                    if pd.isna(valor_diario): valor_diario = 0.0

                    prestador = row.get('Prestador (Relatório CF)', row.get('Prestador (Diário)', 'NÃO IDENTIFICADO'))
                    if pd.isna(prestador) or prestador == 'nan': 
                        prestador = row.get('Prestador (Diário)', 'NÃO IDENTIFICADO')

                    # Regras de Status
                    if pd.isna(row.get('Valor (Diário)')):
                        status = '❌ ERRO: Presente no Relatório CF, mas não achado no Diário'
                    elif pd.isna(row.get('Valor (Relatório CF)')):
                        status = '❌ ERRO: Lançado no Diário, mas não está no Relatório CF'
                    elif not row.get('Estrutura OK', True):
                        status = '⚠️ ALERTA CONTÁBIL: Falta Débito (21101) ou Crédito (Banco 11102)'
                    elif abs(valor_cf - valor_diario) > 0.02: # Margem de erro de 2 centavos
                        status = f'⚠️ DIVERGÊNCIA DE VALOR: CF (R$ {valor_cf:.2f}) x Diário (R$ {valor_diario:.2f})'
                    else:
                        status = '✅ OK: Valores batem perfeitamente'
                        
                    resultados.append({
                        'Número da CF': titulo,
                        'Prestador': prestador,
                        'Valor Relatório CF': valor_cf,
                        'Valor Contabilidade (Diário)': valor_diario,
                        'Diferença (R$)': abs(valor_cf - valor_diario),
                        'Status': status
                    })
                
                df_resultado = pd.DataFrame(resultados)
                
                # Exibir as métricas de resumo
                st.subheader("Resumo da Auditoria")
                erros = df_resultado[~df_resultado['Status'].str.contains('✅ OK')]
                
                metrica1, metrica2, metrica3 = st.columns(3)
                metrica1.metric("Total de CFs Analisadas", len(df_resultado))
                metrica2.metric("Lançamentos Corretos", len(df_resultado) - len(erros))
                metrica3.metric("Divergências Encontradas", len(erros))

                # Exibir tabela na tela
                if not erros.empty:
                    st.error("As seguintes inconsistências foram encontradas:")
                    # Destacar visualmente no dataframe
                    st.dataframe(erros, use_container_width=True)
                else:
                    st.success("Tudo certo! As duas planilhas estão perfeitamente alinhadas.")
                
                # Aba de download do Excel final
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name='Cruzamento')
                    # Configurar largura das colunas no Excel
                    worksheet = writer.sheets['Cruzamento']
                    worksheet.set_column('A:A', 15)
                    worksheet.set_column('B:B', 30)
                    worksheet.set_column('C:E', 20)
                    worksheet.set_column('F:F', 50)
                
                st.write("---")
                st.download_button(
                    label="📥 Baixar Relatório do Cruzamento (Excel)",
                    data=buffer.getvalue(),
                    file_name="cruzamento_diario_vs_relatorioCF.xlsx",
                    mime="application/vnd.ms-excel",
                    type="primary"
                )
