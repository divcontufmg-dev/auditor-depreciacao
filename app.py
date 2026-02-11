import streamlit as st
import pandas as pd
import pdfplumber
import re
from fpdf import FPDF, XPos, YPos
import io
import os

# ==========================================
# CONFIGURAÇÃO DA PÁGINA (VISUAL IDÊNTICO)
# ==========================================
st.set_page_config(
    page_title="Conciliador de Depreciação",
    page_icon="📉",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilos CSS para limpar a interface (Igual ao original)
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ==========================================
# FUNÇÕES DE LÓGICA (DEPRECIAÇÃO)
# ==========================================

def formatar_real(valor):
    """Formata float para moeda BR (R$ 1.234,56)"""
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_moeda_pdf(valor_str):
    if not valor_str: return 0.0
    try:
        limpo = valor_str.replace('.', '').replace(',', '.')
        return float(limpo)
    except:
        return 0.0

def converter_valor_excel(valor):
    if pd.isna(valor): return 0.0
    if isinstance(valor, (int, float)): return float(valor)
    v_str = str(valor).strip().replace('R$', '').replace(' ', '')
    if ',' in v_str: v_str = v_str.replace('.', '').replace(',', '.')
    try: return float(v_str)
    except: return 0.0

def extrair_codigo_grupo(valor_nat_desp):
    try:
        if isinstance(valor_nat_desp, float): valor_nat_desp = int(valor_nat_desp)
        s_val = re.sub(r'\D', '', str(valor_nat_desp).strip())
        if len(s_val) < 5: return None
        return int(s_val[-2:])
    except: return None

def extrair_id_unidade(nome_arquivo):
    match = re.match(r"^(\d+)", nome_arquivo)
    return match.group(1) if match else None

# --- PROCESSAMENTO PDF ---
def processar_pdf(arquivo_obj):
    dados_pdf = {}
    texto_completo = ""
    try:
        with pdfplumber.open(arquivo_obj) as pdf:
            for page in pdf.pages: texto_completo += page.extract_text() + "\n"
    except Exception as e:
        return {}

    # Regex para identificar blocos de Grupos (Ex: "4- APARELHOS...")
    regex_cabecalho = re.compile(r"(?m)^\s*(\d+)\s*-\s*[A-Z]")
    matches = list(regex_cabecalho.finditer(texto_completo))
    
    for i, match in enumerate(matches):
        grupo_id = int(match.group(1))
        start_idx = match.start()
        # Define o fim do bloco
        end_idx = matches[i+1].start() if i + 1 < len(matches) else len(texto_completo)
        bloco_texto = texto_completo[start_idx:end_idx]
        
        # Busca Saldo Atual
        regex_saldo = re.compile(r"\(\*\)\s*SALDO[\s\S]*?ATUAL[\s\S]*?((?:\d{1,3}(?:\.\d{3})*,\d{2}))")
        match_saldo = regex_saldo.search(bloco_texto)
        
        if match_saldo:
            dados_pdf[grupo_id] = formatar_moeda_pdf(match_saldo.group(1))
        else:
            dados_pdf[grupo_id] = 0.0
            
    return dados_pdf

# --- PROCESSAMENTO EXCEL ---
def processar_excel(arquivo_obj):
    try:
        df = pd.read_csv(arquivo_obj, sep=',', encoding='latin1', header=None)
    except:
        try: 
            arquivo_obj.seek(0)
            df = pd.read_excel(arquivo_obj, header=None)
        except: return {}

    linha_cabecalho = -1
    for i, row in df.iterrows():
        if "Nat Desp" in " ".join([str(x) for x in row.values]):
            linha_cabecalho = i; break
            
    if linha_cabecalho == -1: return {}
    
    arquivo_obj.seek(0)
    try:
        if arquivo_obj.name.lower().endswith('.csv'):
             df = pd.read_csv(arquivo_obj, sep=',', encoding='latin1', header=linha_cabecalho)
        else:
             df = pd.read_excel(arquivo_obj, header=linha_cabecalho)
    except: return {}

    # Identificação flexível das colunas
    col_nat_desp = df.columns[0]
    col_saldo = df.columns[-1] # Última coluna é saldo acumulado
    
    dados_excel = {}
    
    for _, row in df.iterrows():
        codigo = extrair_codigo_grupo(row[col_nat_desp])
        if codigo is not None:
            val = abs(converter_valor_excel(row[col_saldo]))
            dados_excel[codigo] = dados_excel.get(codigo, 0.0) + val
            
    return dados_excel

# ==========================================
# CLASSE DE RELATÓRIO PDF (Estilo FPDF Atualizado)
# ==========================================
class PDFRelatorio(FPDF):
    def header(self):
        self.set_font('Helvetica', 'B', 12)
        self.cell(0, 10, 'Relatório de Conciliação - Depreciação Acumulada', 0, 1, 'C')
        self.ln(5)
        
    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

# ==========================================
# INTERFACE (SIDEBAR E MAIN)
# ==========================================

# --- Sidebar (Idêntica ao app original) ---
with st.sidebar:
    st.header("Instruções")
    st.markdown("""
    **Como usar:**
    1.  **Faça o upload** dos arquivos PDF (Relatórios de Depreciação).
    2.  **Faça o upload** dos arquivos Excel/CSV (Razão SIAFI).
    3.  **Clique em Processar**.
    
    **Regra de Cruzamento:**
    O sistema identifica automaticamente os pares pelo **código da unidade** no início do nome do arquivo (ex: `153289.pdf` cruza com `153289...xlsx`).
    
    **Lógica da Análise:**
    * **PDF:** Busca o valor de `(*) SALDO ATUAL` dentro de cada grupo.
    * **Excel:** Soma os saldos das contas contábeis baseando-se nos 2 últimos dígitos da Natureza de Despesa.
    """)
    st.markdown("---")
    st.markdown("**Versão:** 2.0 (Depreciação)")

# --- Área Principal ---
st.title("📉 Conciliador Automático de Depreciação")
st.markdown("Faça o upload dos arquivos abaixo para iniciar a comparação.")

col1, col2 = st.columns(2)

with col1:
    st.subheader("📄 Relatórios (PDF)")
    arquivos_pdf = st.file_uploader("Arraste os PDFs aqui", type=["pdf"], accept_multiple_files=True, key="pdfs")

with col2:
    st.subheader("📊 Razão SIAFI (Excel/CSV)")
    arquivos_excel = st.file_uploader("Arraste os Excel/CSV aqui", type=["xlsx", "csv"], accept_multiple_files=True, key="excels")

# ==========================================
# BOTÃO DE PROCESSAMENTO E LÓGICA FINAL
# ==========================================

if st.button("🚀 Processar Conciliação", type="primary"):
    if not arquivos_pdf or not arquivos_excel:
        st.warning("⚠️ Por favor, carregue arquivos em ambos os lados.")
    else:
        # 1. Agrupar Arquivos
        unidades = {}
        
        for f in arquivos_pdf:
            uid = extrair_id_unidade(f.name)
            if uid:
                if uid not in unidades: unidades[uid] = {}
                unidades[uid]['pdf'] = f
                
        for f in arquivos_excel:
            uid = extrair_id_unidade(f.name)
            if uid:
                if uid not in unidades: unidades[uid] = {}
                unidades[uid]['excel'] = f

        pares_validos = [u for u, arqs in unidades.items() if 'pdf' in arqs and 'excel' in arqs]

        if not pares_validos:
            st.error("❌ Nenhum par correspondente encontrado. Verifique se os nomes dos arquivos começam com o mesmo código (ex: 153289).")
        else:
            # Inicializa Interface de Progresso
            progresso = st.progress(0)
            status_text = st.empty()
            
            # Inicializa PDF Consolidado
            pdf_out = PDFRelatorio()
            pdf_out.set_auto_page_break(auto=True, margin=15)
            pdf_out.add_page()

            resumo_geral = [] # Para mostrar na tela
            
            # Loop de Processamento
            for idx, uid in enumerate(sorted(pares_validos)):
                status_text.text(f"Processando Unidade: {uid}...")
                
                arqs = unidades[uid]
                
                # Reinicia ponteiro para leitura
                arqs['pdf'].seek(0)
                arqs['excel'].seek(0)
                
                # Extração
                d_pdf = processar_pdf(arqs['pdf'])
                d_excel = processar_excel(arqs['excel'])
                
                # Consolidação dos Grupos
                grupos = sorted(list(set(d_pdf.keys()) | set(d_excel.keys())))
                
                divergencias = []
                total_pdf = 0.0
                total_excel = 0.0
                
                for g in grupos:
                    v_p = d_pdf.get(g, 0.0)
                    v_e = d_excel.get(g, 0.0)
                    total_pdf += v_p
                    total_excel += v_e
                    
                    diff = v_p - v_e
                    if abs(diff) > 0.10: # Tolerância 10 centavos
                        divergencias.append({'grupo': g, 'pdf': v_p, 'excel': v_e, 'diff': diff})

                # --- ESCREVENDO NO PDF (LAYOUT CORRIDO) ---
                
                # Verifica se cabe na página
                if pdf_out.get_y() > 240: pdf_out.add_page()

                # Cabeçalho da Unidade (Fundo Cinza claro)
                pdf_out.set_font("Helvetica", 'B', 11)
                pdf_out.set_fill_color(240, 240, 240)
                pdf_out.cell(0, 8, f"Unidade Gestora: {uid}", 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                pdf_out.ln(2)

                # Totais (Estilo Tabela Azulada como no app original)
                pdf_out.set_font("Helvetica", 'B', 9)
                pdf_out.set_fill_color(220, 230, 241) # Azul claro padrão
                
                pdf_out.cell(63, 7, "Total Relatório", 1, fill=True)
                pdf_out.cell(63, 7, "Total SIAFI", 1, fill=True)
                pdf_out.cell(63, 7, "Diferença", 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                
                pdf_out.set_font("Helvetica", '', 9)
                pdf_out.cell(63, 7, f"R$ {formatar_real(total_pdf)}", 1)
                pdf_out.cell(63, 7, f"R$ {formatar_real(total_excel)}", 1)
                
                dif_total = total_pdf - total_excel
                if abs(dif_total) > 0.10: pdf_out.set_text_color(200, 0, 0)
                else: pdf_out.set_text_color(0, 100, 0)
                
                pdf_out.cell(63, 7, f"R$ {formatar_real(dif_total)}", 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                pdf_out.set_text_color(0, 0, 0) # Reseta cor
                pdf_out.ln(3)

                # Divergências
                status_tela = "✅ OK"
                if not divergencias:
                    # Caixa Verde
                    pdf_out.set_fill_color(220, 255, 220)
                    pdf_out.set_font("Helvetica", 'B', 9)
                    pdf_out.cell(0, 8, "CONCILIADO", 1, fill=True, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                else:
                    status_tela = f"❌ {len(divergencias)} Erros"
                    # Caixa Vermelha
                    pdf_out.set_fill_color(255, 220, 220)
                    pdf_out.set_font("Helvetica", 'B', 9)
                    pdf_out.cell(0, 8, "DIVERGÊNCIAS ENCONTRADAS:", 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    
                    # Cabeçalho da Tabela de Erros
                    pdf_out.set_fill_color(250, 250, 250)
                    pdf_out.set_font("Helvetica", 'B', 8)
                    pdf_out.cell(20, 6, "Grupo", 1, fill=True, align='C')
                    pdf_out.cell(56, 6, "Saldo Relatório", 1, fill=True, align='C')
                    pdf_out.cell(56, 6, "Saldo SIAFI", 1, fill=True, align='C')
                    pdf_out.cell(57, 6, "Diferença", 1, fill=True, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    
                    pdf_out.set_font("Helvetica", '', 8)
                    for d in divergencias:
                        pdf_out.cell(20, 6, str(d['grupo']), 1, align='C')
                        pdf_out.cell(56, 6, formatar_real(d['pdf']), 1, align='R')
                        pdf_out.cell(56, 6, formatar_real(d['excel']), 1, align='R')
                        pdf_out.set_text_color(200, 0, 0)
                        pdf_out.cell(57, 6, formatar_real(d['diff']), 1, align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                        pdf_out.set_text_color(0, 0, 0)

                pdf_out.ln(5)
                # Linha separadora tracejada visual
                pdf_out.cell(0, 0, "", "B", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                pdf_out.ln(5)
                
                # Adiciona ao resumo da tela
                resumo_geral.append({
                    "Unidade": uid,
                    "Status": status_tela,
                    "Diferença Total": f"R$ {formatar_real(dif_total)}"
                })
                
                progresso.progress((idx + 1) / len(pares_validos))

            # Finalização
            progresso.empty()
            status_text.success("Processamento concluído com sucesso!")
            
            # 1. Mostra Tabela Resumo na Tela
            st.markdown("### Resumo da Análise")
            st.dataframe(pd.DataFrame(resumo_geral), use_container_width=True)
            
            # 2. Botão de Download do PDF
            pdf_bytes = pdf_out.output(dest='S').encode('latin-1')
            st.download_button(
                label="📥 Baixar Relatório Consolidado (PDF)",
                data=pdf_bytes,
                file_name="Relatorio_Depreciacao_Consolidado.pdf",
                mime="application/pdf",
                type="primary"
            )
