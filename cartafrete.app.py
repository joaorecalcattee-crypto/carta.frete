import streamlit as st
import pdfplumber
import re
import pandas as pd
import io

st.set_page_config(page_title="Conferencia Diario - Cartas Frete", layout="wide")

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

def limpar_valor(valor_str):
    if not valor_str: return 0.0
    return float(valor_str.replace('.', '').replace(',', '.'))

def processar_livro_diario_excel(arquivo_excel):
    try:
        # Lê o Excel sem focar em cabeçalhos, para varrer todas as células
        df = pd.read_excel(arquivo_excel, header=None)
    except Exception as e:
        st.error(f"Erro ao ler Excel: {e}")
        return pd.DataFrame()

    lancamentos = []
    
    for index, row in df.iterrows():
        # Junta o texto da linha inteira para achar o histórico "PAGO CF"
        linha_str = ' '.join([str(val) for val in row if pd.notna(val)])
        
        match_cf = re.search(r'PAGO\s*C[FE]\s*(\d+)', linha_str, re.IGNORECASE)
        
        if match_cf:
            numero_titulo = match_cf.group(1)
            
            # Encontra automaticamente as colunas que possuem valores financeiros (Débito/Crédito)
            valores_numericos = [val for val in row if isinstance(val, (int, float))]
            
            if valores_numericos:
                # Captura o maior valor absoluto daquela linha
                valor = max([abs(v) for v in valores_numericos])
            else:
                valor = 0.0
                
            lancamentos.append({
                'Número do Título': numero_titulo,
                'Valor (Diário)': float(valor)
            })
            
    return pd.DataFrame(lancamentos).drop_duplicates('Número do Título')

def processar_cartas_frete(texto):
    cartas = []
    blocos = re.split(r'CONTRATO DE TRANSPORTES', texto, flags=re.IGNORECASE)
    
    for bloco in blocos:
        match_numero = re.search(r'N[UÚ]MERO:\s*(\d+)', bloco, re.IGNORECASE)
        if not match_numero:
            continue
        numero_titulo = match_numero.group(1)
        
        match_valor = re.search(r'SALDO A RECEBER:\s*R\$\s*([\d\.,]+)', bloco, re.IGNORECASE)
        valor_liquido = limpar_valor(match_valor.group(1)) if match_valor else 0.0
        
        cartas.append({
            'Número do Título': numero_titulo,
            'Valor (Carta Frete)': valor_liquido
        })
        
    return pd.DataFrame(cartas)

st.title("📊 Conferencia Contábil: Livro Diário vs. Cartas Frete")
st.write("Faça o upload do Livro Diário (Excel) e das Cartas Frete (PDF).")

col1, col2 = st.columns(2)
with col1:
    # Atualizado para aceitar formatos do Excel
    arquivo_diario = st.file_uploader("Upload do Livro Diário (Excel)", type=['xlsx', 'xls'])
with col2:
    arquivo_cartas = st.file_uploader("Upload das Cartas Frete (PDF)", type=['pdf'])

if arquivo_diario and arquivo_cartas:
    if st.button("Iniciar Conferência", type="primary"):
        with st.spinner("Cruzando dados fiscais..."):
            
            df_diario = processar_livro_diario_excel(arquivo_diario)
            
            texto_cartas = extrair_texto_pdf(arquivo_cartas)
            df_cartas = processar_cartas_frete(texto_cartas)
            
            if df_diario.empty:
                st.error("Não foi possível localizar lançamentos 'PAGO CF' no arquivo Excel enviado.")
            elif df_cartas.empty:
                st.error("Falha ao extrair dados dos PDFs das Cartas Frete.")
            else:
                df_cruzamento = pd.merge(df_cartas, df_diario, on='Número do Título', how='outer')
                
                resultados = []
                for index, row in df_cruzamento.iterrows():
                    titulo = row['Número do Título']
                    if pd.isna(row['Valor (Diário)']):
                        status = 'Erro: Presente na Carta Frete, ausente no Diário'
                    elif pd.isna(row['Valor (Carta Frete)']):
                        status = 'Erro: Presente no Diário, Carta Frete ausente'
                    elif abs(row['Valor (Carta Frete)'] - row['Valor (Diário)']) > 0.01:
                        status = 'Divergência de Valor'
                    else:
                        status = 'OK: Lançamento validado'
                        
                    resultados.append({
                        'Título': titulo,
                        'Valor Carta Frete': row.get('Valor (Carta Frete)', 0.0),
                        'Valor Diário': row.get('Valor (Diário)', 0.0),
                        'Status': status
                    })
                
                df_resultado = pd.DataFrame(resultados)
                
                st.subheader("Resultado da Validação")
                erros = df_resultado[df_resultado['Status'] != 'OK: Lançamento validado']
                sucessos = df_resultado[df_resultado['Status'] == 'OK: Lançamento validado']
                
                if not erros.empty:
                    st.warning(f"Foram encontradas {len(erros)} inconsistências.")
                    st.dataframe(erros, use_container_width=True)
                else:
                    st.success("Conferência concluída! Nenhuma divergência encontrada.")
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name='Conciliacao')
                
                st.download_button(
                    label="Baixar Relatório (Excel)",
                    data=buffer.getvalue(),
                    file_name="relatorio_conferencia.xlsx",
                    mime="application/vnd.ms-excel"
                )
