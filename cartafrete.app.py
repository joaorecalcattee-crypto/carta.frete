import streamlit as st
import pdfplumber
import re
import pandas as pd
import io

st.set_page_config(page_title="Conferencia Contábil - Cartas Frete", layout="wide")

# ==========================================
# 1. SISTEMA DE LOGIN (CONTROLE DE ACESSO)
# ==========================================

# Aqui você cadastra os usuários e senhas (formato -> "usuario": "senha")
USUARIOS_CADASTRADOS = {
    "joao.recalcatte": "alfa2026",
    "edson.reis": "mudar123",
    "admin": "admin"
}

def check_password():
    """Retorna True se o usuário inseriu as credenciais corretas."""
    def password_entered():
        # Verifica se o usuário existe e se a senha está correta
        if st.session_state["username"] in USUARIOS_CADASTRADOS and st.session_state["password"] == USUARIOS_CADASTRADOS[st.session_state["username"]]:
            st.session_state["password_correct"] = True
            
            # CORREÇÃO: Salva o nome em uma variável que não será apagada quando o campo sumir
            st.session_state["usuario_logado"] = st.session_state["username"] 
            
            del st.session_state["password"]  # Apaga a senha da memória por segurança
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # Primeira tela: Mostra formulário de login
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.title("🔒 Acesso Restrito")
            st.write("Por favor, faça o login para acessar o Validador Fiscal.")
            st.text_input("Usuário", key="username")
            st.text_input("Senha", type="password", key="password")
            st.button("Entrar", on_click=password_entered, type="primary")
        return False
    
    elif not st.session_state["password_correct"]:
        # Erro de senha: Mostra o formulário novamente com aviso
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.title("🔒 Acesso Restrito")
            st.text_input("Usuário", key="username")
            st.text_input("Senha", type="password", key="password")
            st.button("Entrar", on_click=password_entered, type="primary")
            st.error("Usuário ou senha incorretos.")
        return False
    
    else:
        # Tudo certo, deixa o código seguir
        return True

# SE O USUÁRIO NÃO PASSAR NO TESTE DE SENHA, O CÓDIGO PARA AQUI
if not check_password():
    st.stop()

# Adiciona um botão de Sair na barra lateral usando a nova variável salva
st.sidebar.success(f"Logado como: {st.session_state['usuario_logado']}")
if st.sidebar.button("Sair"):
    del st.session_state["password_correct"]
    del st.session_state["usuario_logado"]
    st.rerun()

# ==========================================
# 2. SISTEMA DO VALIDADOR (COMEÇA AQUI O RESTO DO SEU CÓDIGO)
# ==========================================

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
        match_cf = re.search(r'PAGO\s*C[FE]\s*(\d+)(?:\s*-\s*(.+))?', historico, re.IGNORECASE)
        
        if match_cf:
            numero_titulo = match_cf.group(1)
            prestador_excel = match_cf.group(2).strip() if match_cf.group(2) else ""
            conta = str(row[col_conta]).replace('.', '')
            
            if numero_titulo not in lancamentos_agrupados:
                lancamentos_agrupados[numero_titulo] = {
                    'debito_frete': 0.0,
                    'tem_credito_banco': False,
                    'prestador': prestador_excel
                }
            
            if not lancamentos_agrupados[numero_titulo]['prestador'] and prestador_excel:
                lancamentos_agrupados[numero_titulo]['prestador'] = prestador_excel
            
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
            'Prestador_Diario': dados['prestador'],
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
        
        match_prestador = re.search(r'CONTRATADO[\s\S]*?NOME[\s\S]*?:\s*([^\n]+)', bloco, re.IGNORECASE)
        prestador_pdf = match_prestador.group(1).strip() if match_prestador else ""
        
        match_bruto = re.search(r'VALOR BRUTO:\s*R\$\s*([\d\.,]+)', bloco, re.IGNORECASE)
        valor_bruto = limpar_valor_pdf(match_bruto.group(1)) if match_bruto else 0.0
        
        match_inss = re.search(r'INSS:\s*R\$\s*([\d\.,]+)', bloco, re.IGNORECASE)
        inss = limpar_valor_pdf(match_inss.group(1)) if match_inss else 0.0
        
        match_ir = re.search(r'IR:\s*R\$\s*([\d\.,]+)', bloco, re.IGNORECASE)
        ir = limpar_valor_pdf(match_ir.group(1)) if match_ir else 0.0
        
        match_sest = re.search(r'SEST:\s*R\$\s*([\d\.,]+)', bloco, re.IGNORECASE)
        sest = limpar_valor_pdf(match_sest.group(1)) if match_sest else 0.0
        
        match_saldo = re.search(r'SALDO A RECEBER:\s*R\$\s*([\d\.,]+)', bloco, re.IGNORECASE)
        saldo_receber = limpar_valor_pdf(match_saldo.group(1)) if match_saldo else 0.0
        
        cartas.append({
            'Número do Título': numero_titulo,
            'Prestador_Carta': prestador_pdf,
            'Valor Bruto (CF)': valor_bruto,
            'Total Impostos (CF)': inss + ir + sest,
            'INSS': inss,
            'IR': ir,
            'SEST': sest,
            'Saldo a Receber (CF)': saldo_receber
        })
        
    return pd.DataFrame(cartas)

st.title("📊 Conferencia Contabil: Livro Diário vs. Cartas Frete")
st.write("Auditoria de Partidas Dobradas, Valores, Impostos e Prestadores.")

col1, col2 = st.columns(2)
with col1:
    arquivo_diario = st.file_uploader("Upload do Livro Diário (Excel)", type=['xlsx', 'xls'])
with col2:
    arquivo_cartas = st.file_uploader("Upload das Cartas Frete (PDF)", type=['pdf'])

if arquivo_diario and arquivo_cartas:
    if st.button("Iniciar Auditoria", type="primary"):
        with st.spinner("Realizando auditoria completa..."):
            
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
                    
                    prestador_diario = str(row.get('Prestador_Diario', '')).strip()
                    prestador_carta = str(row.get('Prestador_Carta', '')).strip()
                    
                    if prestador_diario and prestador_diario.lower() != 'nan':
                        prestador = prestador_diario
                    elif prestador_carta and prestador_carta.lower() != 'nan':
                        prestador = prestador_carta
                    else:
                        prestador = 'NÃO IDENTIFICADO'
                    
                    bruto = row.get('Valor Bruto (CF)', 0.0)
                    impostos = row.get('Total Impostos (CF)', 0.0)
                    saldo_cf = row.get('Saldo a Receber (CF)', 0.0)
                    valor_diario = row.get('Valor (Diário)', 0.0)
                    
                    matematica_cf_ok = abs((bruto - impostos) - saldo_cf) <= 0.01

                    if pd.isna(row['Valor (Diário)']):
                        status = 'ERRO: Presente na Carta Frete, ausente no Diário'
                    elif pd.isna(row['Saldo a Receber (CF)']):
                        status = 'ERRO: Presente no Diário, Carta Frete ausente'
                    elif not row['Estrutura OK']:
                        status = 'ERRO CONTÁBIL: Falta Débito (21101) ou Crédito (Banco 11102)'
                    elif not matematica_cf_ok:
                        status = f'ERRO NA CF: Bruto (R$ {bruto:.2f}) - Impostos (R$ {impostos:.2f}) não bate com Saldo CF (R$ {saldo_cf:.2f})'
                    elif abs(saldo_cf - valor_diario) > 0.01:
                        status = f'DIVERGÊNCIA: Saldo CF (R$ {saldo_cf:.2f}) diferente do Livro Diário (R$ {valor_diario:.2f})'
                    else:
                        status = 'OK: Lançamento Validado (Matemática e Contabilidade Exatas)'
                        
                    resultados.append({
                        'Título': titulo,
                        'Prestador': prestador,
                        'Valor Bruto (CF)': bruto,
                        'Total Impostos (CF)': impostos,
                        'Saldo a Receber (CF)': saldo_cf,
                        'Valor Diário (Excel)': valor_diario,
                        'Status': status
                    })
                
                df_resultado = pd.DataFrame(resultados)
                
                st.subheader("Resultado da Validação")
                erros = df_resultado[df_resultado['Status'] != 'OK: Lançamento Validado (Matemática e Contabilidade Exatas)']
                
                if not erros.empty:
                    st.warning(f"Atenção: Encontramos {len(erros)} inconsistência(s) que exigem verificação.")
                    st.dataframe(erros, use_container_width=True)
                else:
                    st.success("Auditoria Perfeita! Todos os valores, matemáticas e contas contábeis estão exatos.")
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name='Auditoria')
                
                st.download_button(
                    label="Baixar Relatório Completo (Excel)",
                    data=buffer.getvalue(),
                    file_name="relatorio_auditoria_fretes.xlsx",
                    mime="application/vnd.ms-excel"
                )
