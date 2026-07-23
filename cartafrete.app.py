import streamlit as st
import pdfplumber
import re
import pandas as pd
import io

st.set_page_config(page_title="Conferencia Contabil - Cartas Frete", layout="wide")

def extrair_texto_pdf(arquivo_upload):
    texto_completo = ""
    try:
        with pdfplumber.open(arquivo_upload) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text()
                if texto:
                    texto_completo += texto + "\n"
    except Exception as e:
        st.error(f"Erro ao ler PDF: {e}")
    return texto_completo

def limpar_valor_pdf(valor_str):
    if not valor_str: return 0.0
    return float(valor_str.replace('.', '').replace(',', '.'))

def limpar_valor_excel(valor_str):
    valor_str = str(valor_str).strip()
    if valor_str.lower() in ['nan', 'none', ''] or not valor_str:
        return 0.0
    try:
        return float(valor_str.replace('.', '').replace(',', '.'))
    except:
        return 0.0

def processar_livro_diario_excel(arquivo_excel):
    try:
        df = pd.read_excel(arquivo_excel)
    except Exception as e:
        st.error(f"Erro ao ler Excel: {e}")
        return pd.DataFrame()

    col_hist = next((col for col in df.columns if 'Hist' in str(col)), None)
    col_debito = next((col for col in df.columns if 'Debito' in str(col)), None)
    col_credito = next((col for col in df.columns if 'Credito' in str(col)), None)
    col_conta = next((col for col in df.columns if 'Conta' in str(col)), None)

    if not col_hist or not col_debito or not col_credito or not col_conta:
        st.error("As colunas de Histórico, Débito, Crédito ou Conta Contábil não foram localizadas.")
        return pd.DataFrame()

    lancamentos_agrupados = {}

    for index, row in df.iterrows():
        historico = str(row[col_hist])
        match_cf = re.search(r'PAGO\s*C[FE]\s*(\d+)', historico, re.IGNORECASE)
        
        if match_cf:
            numero_titulo = match_cf.group(1)
            conta = str(row[col_conta]).replace('.', '')
            
            if numero_titulo not in lancamentos_agrupados:
                lancamentos_agrupados[numero_titulo] = {
                    'debito_frete': 0.0,
                    'tem_credito_banco': False
                }
            
            if conta.startswith('211010300001'):
                valor_deb = limpar_valor_excel(row[col_debito])
                if valor_deb > 0:
                    lancamentos_agrupados[numero_titulo]['debito_frete'] += valor_deb
                    
            if conta.startswith('11102'):
                valor_cred = limpar_valor_excel(row[col_credito])
                if valor_cred > 0:
                    lancamentos_agrupados[numero_titulo]['tem_credito_banco'] = True

    lancamentos_finais = []
    for titulo, dados in lancamentos_agrupados.items():
        lancamentos_finais.append({
            'Número do Título': titulo,
            'Valor (Diário)': dados['debito_frete'],
            'Estrutura OK': dados['tem_credito_banco'] and (dados['debito_frete'] > 0)
        })
            
    return pd.DataFrame(lancamentos_finais)

def processar_cartas_frete(texto):
    cartas = []
    blocos = re.split(r'CONTRATO DE TRANSPORTES', texto, flags=re.IGNORECASE)
    
    for bloco in blocos:
        match_numero = re.search(r'N[UÚ]MERO:\s*(\d+)', bloco, re.IGNORECASE)
        if not match_numero:
            continue
        numero_titulo = match_numero.group(1)
        
        # Nova Extração: Nome do Prestador
        match_prestador = re.search(r'CONTRATADO\s+NOME\s*:\s*([^\n]+)', bloco, re.IGNORECASE)
        prestador = match_prestador.group(1).strip() if match_prestador else "NÃO IDENTIFICADO"
        
        match_valor = re.search(r'SALDO A RECEBER:\s*R\$\s*([\d\.,]+)', bloco, re.IGNORECASE)
        valor_liquido = limpar_valor_pdf(match_valor.group(1)) if match_valor else 0.0
        
        cartas.append({
            'Número do Título': numero_titulo,
            'Prestador': prestador,
            'Valor (Carta Frete)': valor_liquido
        })
        
    return pd.DataFrame(cartas)

st.title("📊 Conferencia Contabil: Livro Diário vs. Cartas Frete")
st.write("Auditoria de Partidas Dobradas, Valores e Prestadores.")

col1, col2 = st.columns(2)
with col1:
    arquivo_diario = st.file_uploader("Upload do Livro Diário (Excel)", type=['xlsx', 'xls'])
with col2:
    arquivo_cartas = st.file_uploader("Upload das Cartas Frete (PDF)", type=['pdf'])

if arquivo_diario and arquivo_cartas:
    if st.button("Iniciar Auditoria", type="primary"):
        with st.spinner("Realizando enfrentamento contábil..."):
            
            df_diario = processar_livro_diario_excel(arquivo_diario)
            texto_cartas = extrair_texto_pdf(arquivo_cartas)
            df_cartas = processar_cartas_frete(texto_cartas)
            
            if df_diario.empty:
                st.error("Nenhum lançamento válido encontrado no Excel do Livro Diário.")
            elif df_cartas.empty:
                st.error("Nenhuma Carta Frete identificada nos PDFs.")
            else:
                df_cruzamento = pd.merge(df_cartas, df_diario, on='Número do Título', how='outer')
                
                resultados = []
                for index, row in df_cruzamento.iterrows():
                    titulo = row['Número do Título']
                    prestador = row.get('Prestador', 'NÃO IDENTIFICADO (Falta Carta Frete)')
                    
                    if pd.isna(row['Valor (Diário)']):
                        status = 'Erro: Presente na Carta Frete, ausente no Diário'
                    elif pd.isna(row['Valor (Carta Frete)']):
                        status = 'Erro: Presente no Diário, Carta Frete ausente'
                    elif not row['Estrutura OK']:
                        status = 'ERRO CONTÁBIL: Falta Débito (21101) ou Crédito (Banco 11102)'
                    elif abs(row['Valor (Carta Frete)'] - row['Valor (Diário)']) > 0.01:
                        status = 'Divergência de Valor'
                    else:
                        status = 'OK: Lançamento validado (Partidas Dobradas Corretas)'
                        
                    resultados.append({
                        'Título': titulo,
                        'Prestador': prestador,
                        'Valor Carta Frete': row.get('Valor (Carta Frete)', 0.0),
                        'Valor Diário': row.get('Valor (Diário)', 0.0),
                        'Status': status
                    })
                
                df_resultado = pd.DataFrame(resultados)
                
                st.subheader("Resultado da Validação")
                erros = df_resultado[df_resultado['Status'] != 'OK: Lançamento validado (Partidas Dobradas Corretas)']
                sucessos = df_resultado[df_resultado['Status'] == 'OK: Lançamento validado (Partidas Dobradas Corretas)']
                
                if not erros.empty:
                    st.warning(f"Atenção: Encontramos {len(erros)} inconsistência(s) que precisam de revisão.")
                    st.dataframe(erros, use_container_width=True)
                else:
                    st.success("Auditoria perfeita! Todos os valores e contas contábeis estão batendo.")
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name='Auditoria')
                
                st.download_button(
                    label="Baixar Relatório Completo (Excel)",
                    data=buffer.getvalue(),
                    file_name="relatorio_auditoria_fretes.xlsx",
                    mime="application/vnd.ms-excel"
                )
