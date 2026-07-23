import streamlit as st
import pdfplumber
import re
import pandas as pd
import io

# Configuração da página da interface web
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
        st.error(f"Erro ao processar o documento: {e}")
    return texto_completo

def limpar_valor(valor_str):
    if not valor_str: return 0.0
    return float(valor_str.replace('.', '').replace(',', '.'))

def processar_livro_diario(texto):
    lancamentos = []
    linhas = texto.split('\n')
    for i, linha in enumerate(linhas):
        match_cf = re.search(r'PAGO C[FE]\s*(\d+)', linha)
        if match_cf:
            numero_titulo = match_cf.group(1)
            match_valor = re.search(r'(\d{1,3}(?:\.\d{3})*,\d{2})', linha)
            if not match_valor and i + 1 < len(linhas):
                match_valor = re.search(r'(\d{1,3}(?:\.\d{3})*,\d{2})', linhas[i+1])
            valor = limpar_valor(match_valor.group(1)) if match_valor else 0.0
            prestador = re.sub(r'PAGO C[FE]\s*\d+', '', linha).strip()
            lancamentos.append({
                'Número do Título': numero_titulo,
                'Prestador (Diário)': prestador,
                'Valor (Diário)': valor
            })
    return pd.DataFrame(lancamentos).drop_duplicates('Número do Título')

def processar_cartas_frete(texto):
    cartas = []
    blocos = texto.split('CONTRATO DE TRANSPORTES')
    for bloco in blocos:
        if 'NÚMERO:' not in bloco:
            continue
        match_numero = re.search(r'NÚMERO:\s*(\d+)', bloco)
        if not match_numero:
            continue
        numero_titulo = match_numero.group(1)
        match_valor = re.search(r'SALDO A RECEBER:\s*R\$\s*([\d\.,]+)', bloco)
        valor_liquido = limpar_valor(match_valor.group(1)) if match_valor else 0.0
        cartas.append({
            'Número do Título': numero_titulo,
            'Valor (Carta Frete)': valor_liquido
        })
    return pd.DataFrame(cartas)

# Construção da Interface Web
st.title("📊 Validador Fiscal: Livro Diário vs. Cartas Frete")
st.write("Faça o upload dos documentos em PDF para realizar a conciliação automática dos lançamentos.")

# Áreas de Upload
col1, col2 = st.columns(2)
with col1:
    arquivo_diario = st.file_uploader("Upload do Livro Diário (PDF)", type=['pdf'])
with col2:
    arquivo_cartas = st.file_uploader("Upload das Cartas Frete (PDF)", type=['pdf'])

if arquivo_diario and arquivo_cartas:
    if st.button("Iniciar Conferência", type="primary"):
        with st.spinner("Analisando documentos..."):
            
            texto_diario = extrair_texto_pdf(arquivo_diario)
            texto_cartas = extrair_texto_pdf(arquivo_cartas)
            
            df_diario = processar_livro_diario(texto_diario)
            df_cartas = processar_cartas_frete(texto_cartas)
            
            if df_diario.empty or df_cartas.empty:
                st.error("Não foi possível extrair dados estruturados dos documentos fornecidos.")
            else:
                # Cruzamento
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
                
                # Exibição Visual dos Resultados
                st.subheader("Resultado da Validação")
                
                erros = df_resultado[df_resultado['Status'] != 'OK: Lançamento validado']
                sucessos = df_resultado[df_resultado['Status'] == 'OK: Lançamento validado']
                
                if not erros.empty:
                    st.warning(f"Foram encontradas {len(erros)} inconsistências que exigem revisão.")
                    st.dataframe(erros, use_container_width=True)
                else:
                    st.success("Conferência concluída com sucesso. Nenhuma divergência encontrada.")
                
                st.write(f"**Total de registros validados com sucesso:** {len(sucessos)}")
                
                # Botão para exportar relatório
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name='Conciliacao')
                
                st.download_button(
                    label="Baixar Relatório Completo (Excel)",
                    data=buffer.getvalue(),
                    file_name="relatorio_conferencia.xlsx",
                    mime="application/vnd.ms-excel"
                )
