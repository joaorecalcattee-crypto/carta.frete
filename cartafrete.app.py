import streamlit as st
import pdfplumber
import re
import pandas as pd
import io

st.set_page_config(page_title="Validador Fiscal - Cartas Frete", layout="wide")

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

def processar_livro_diario(texto):
    lancamentos = []
    linhas = texto.split('\n')
    
    for i, linha in enumerate(linhas):
        # Regex mais flexível (ignora maiúsculas/minúsculas, acentos e espaços extras)
        match_cf = re.search(r'PAGO\s*C[FE]\s*(\d+)', linha, re.IGNORECASE)
        
        if match_cf:
            numero_titulo = match_cf.group(1)
            
            # Ampliamos a busca do valor para a linha atual, a anterior e a próxima
            linhas_busca = [linha]
            if i < len(linhas) - 1: linhas_busca.append(linhas[i+1])
            if i > 0: linhas_busca.append(linhas[i-1])
            
            valor = 0.0
            for linha_b in linhas_busca:
                match_valor = re.search(r'(\d{1,3}(?:\.\d{3})*,\d{2})', linha_b)
                if match_valor:
                    valor = limpar_valor(match_valor.group(1))
                    break
                    
            prestador = re.sub(r'PAGO\s*C[FE]\s*\d+', '', linha, flags=re.IGNORECASE).strip()
            lancamentos.append({
                'Número do Título': numero_titulo,
                'Prestador (Diário)': prestador,
                'Valor (Diário)': valor
            })
            
    return pd.DataFrame(lancamentos).drop_duplicates('Número do Título')

def processar_cartas_frete(texto):
    cartas = []
    # Usando split mais seguro
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

st.title("📊 Validador Fiscal: Livro Diário vs. Cartas Frete")
st.write("Faça o upload dos documentos para iniciar a conciliação.")

col1, col2 = st.columns(2)
with col1:
    arquivo_diario = st.file_uploader("Upload do Livro Diário (PDF)", type=['pdf'])
with col2:
    arquivo_cartas = st.file_uploader("Upload das Cartas Frete (PDF)", type=['pdf'])

# Nova ferramenta de diagnóstico visual
mostrar_debug = st.checkbox("Modo de Depuração (Mostrar texto extraído dos PDFs)")

if arquivo_diario and arquivo_cartas:
    if st.button("Iniciar Conferência", type="primary"):
        with st.spinner("Analisando documentos..."):
            
            texto_diario = extrair_texto_pdf(arquivo_diario)
            texto_cartas = extrair_texto_pdf(arquivo_cartas)
            
            if mostrar_debug:
                with st.expander("Ver texto extraído do Livro Diário"):
                    st.text(texto_diario[:2000]) # Mostra os primeiros 2000 caracteres
                with st.expander("Ver texto extraído das Cartas Frete"):
                    st.text(texto_cartas[:2000])

            df_diario = processar_livro_diario(texto_diario)
            df_cartas = processar_cartas_frete(texto_cartas)
            
            if df_diario.empty and df_cartas.empty:
                st.error("Não foi possível extrair dados de NENHUM dos documentos. Ative o Modo de Depuração acima para ver como o sistema está lendo os arquivos.")
            elif df_diario.empty:
                st.error("Dados extraídos das Cartas Frete, mas o LIVRO DIÁRIO falhou. Ative o Modo de Depuração.")
            elif df_cartas.empty:
                st.error("Dados extraídos do Livro Diário, mas as CARTAS FRETE falharam. Ative o Modo de Depuração.")
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
