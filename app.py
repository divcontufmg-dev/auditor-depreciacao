import streamlit as st
import pandas as pd
import pdfplumber
import re
from fpdf import FPDF, XPos, YPos
import io
import os

# ==========================================
# CONFIGURAÇÃO INICIAL (Idêntica à ferramenta RMB)
# ==========================================
st.set_page_config(
    page_title="Conciliador Depreciação x SIAFI",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilos CSS
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ==========================================
# FUNÇÕES AUXILIARES
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

# ==========================================
# MOTORES DE EXTRAÇÃO (Lógica de Depreciação)
# ==========================================

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
        end_idx = matches[i+1].start() if i + 1 < len(matches) else len(texto_completo)
        bloco_texto = texto_completo[start_idx:end_idx]
        
        # Busca Saldo Atual no bloco
        regex_saldo = re.compile(r"\(\*\)\s*SALDO[\s\S]*?ATUAL[\s\S]*?((?:\d{1,3}(?:\.\d{3})*,\d{2}))")
        match_saldo = regex_saldo.search(bloco_texto)
        
        if match_saldo:
            dados_pdf[grupo_id] = formatar_moeda_pdf(match_saldo.group(1))
        else:
            dados_pdf[grupo_id] = 0.0
            
    return dados_pdf

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

    col_nat_desp = df.columns[0]
    col_saldo = df.columns[-1]
    
    dados_excel = {}
    
    for _, row in df.iterrows():
        codigo = extrair_codigo_grupo(row[col_nat_desp])
        if codigo is not None:
            val = abs(converter_valor_excel(row[col_saldo]))
            dados_excel[codigo] = dados_excel.get(codigo, 0.0) + val
            
    return dados_excel

# ==========================================
# CLASSE PDF (Visual do Relatório)
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
# INTERFACE DO USUÁRIO (LAYOUT PADRONIZADO)
# ==========================================

# --- Sidebar ---
with st.sidebar:
    st.header("Instruções")
    st.markdown("""
    **Passo a Passo:**
    1.  **Arraste todos os arquivos** para a área de upload (PDFs e Planilhas misturados).
    2.  O sistema identificará os pares automaticamente pelo **código da unidade** (início do nome).
    3.  Clique em **Processar Arquivos**.
    
    **Arquivos Aceitos:**
    * **PDF:** Relatório de Depreciação (contendo "(*) SALDO ATUAL").
    * **Excel/CSV:** Razão Auxiliar do SIAFI.
    """)
    st.markdown("---")
    st.markdown("**Versão:** 2.2 (Layout Unificado)")

# --- Área Principal ---
st.title("📊 Conciliador Automático: Depreciação (PDF) x Razão SIAFI (Excel)")
st.markdown("Carregue seus arquivos abaixo para iniciar a conferência automática.")

uploaded_files = st.file_uploader(
    "Carregue seus arquivos aqui (PDFs e Excels misturados)", 
    type=['pdf', 'xlsx', 'csv'], 
    accept_multiple_files=True
)

if st.button("Processar Arquivos", type="primary"):
    if not uploaded_files:
        st.warning("⚠️ Por favor, insira os arquivos para processar.")
    else:
        # Separação
        pdfs = [f for f in uploaded_files if f.name.lower().endswith('.pdf')]
        excels = [f for f in uploaded_files if f.name.lower().endswith(('.xlsx', '.csv'))]
        
        st.info(f"Arquivos identificados: {len(pdfs)} PDFs e {len(excels)} Planilhas.")
        
        if not pdfs or not excels:
            st.error("❌ Necessário pelo menos 1 PDF e 1 Excel/CSV para cruzar os dados.")
        else:
            # Pareamento
            unidades = {}
            for f in pdfs:
                uid = extrair_id_unidade(f.name)
                if uid:
                    if uid not in unidades: unidades[uid] = {}
                    unidades[uid]['pdf'] = f
            
            for f in excels:
                uid = extrair_id_unidade(f.name)
                if uid:
                    if uid not in unidades: unidades[uid] = {}
                    unidades[uid]['excel'] = f
            
            # Filtra pares completos
            pares = [uid for uid, docs in unidades.items() if 'pdf' in docs and 'excel' in docs]
            
            if not pares:
                st.error("❌ Nenhum par de arquivos correspondente encontrado (Verifique os nomes).")
            else:
                # Início do Processamento
                progresso = st.progress(0)
                status_box = st.empty()
                
                # Setup PDF
                pdf_out = PDFRelatorio()
                pdf_out.set_auto_page_break(auto=True, margin=15)
                pdf_out.add_page()
                
                lista_resumo = []
                
                for idx, uid in enumerate(sorted(pares)):
                    status_box.text(f"Processando Unidade: {uid}...")
                    
                    docs = unidades[uid]
                    # Reset ponteiros
                    docs['pdf'].seek(0)
                    docs['excel'].seek(0)
                    
                    # Extração
                    d_pdf = processar_pdf(docs['pdf'])
                    d_excel = processar_excel(docs['excel'])
                    
                    # Consolidação
                    grupos = sorted(list(set(d_pdf.keys()) | set(d_excel.keys())))
                    divergencias = []
                    soma_pdf = 0.0
                    soma_excel = 0.0
                    
                    for g in grupos:
                        vp = d_pdf.get(g, 0.0)
                        ve = d_excel.get(g, 0.0)
                        soma_pdf += vp
                        soma_excel += ve
                        
                        dif = vp - ve
                        if abs(dif) > 0.10: # Tolerância 10 centavos
                            divergencias.append({'grupo': g, 'pdf': vp, 'excel': ve, 'diff': dif})

                    # --- GERAÇÃO PDF ---
                    if pdf_out.get_y() > 250: pdf_out.add_page()
                    
                    # Título Unidade
                    pdf_out.set_font("helvetica", 'B', 11)
                    pdf_out.set_fill_color(240, 240, 240)
                    pdf_out.cell(0, 8, f"Unidade Gestora: {uid}", 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    pdf_out.ln(2)
                    
                    # Tabela Totais
                    pdf_out.set_font("helvetica", 'B', 9)
                    pdf_out.set_fill_color(220, 230, 241)
                    pdf_out.cell(63, 7, "Total Relatório", 1, fill=True)
                    pdf_out.cell(63, 7, "Total SIAFI", 1, fill=True)
                    pdf_out.cell(63, 7, "Diferença", 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    
                    pdf_out.set_font("helvetica", '', 9)
                    pdf_out.cell(63, 7, f"R$ {formatar_real(soma_pdf)}", 1)
                    pdf_out.cell(63, 7, f"R$ {formatar_real(soma_excel)}", 1)
                    
                    dif_total = soma_pdf - soma_excel
                    if abs(dif_total) > 0.10: pdf_out.set_text_color(200, 0, 0)
                    else: pdf_out.set_text_color(0, 100, 0)
                    
                    pdf_out.cell(63, 7, f"R$ {formatar_real(dif_total)}", 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    pdf_out.set_text_color(0, 0, 0)
                    pdf_out.ln(3)
                    
                    # Status
                    status_str = "✅ Conciliado"
                    if not divergencias:
                        pdf_out.set_fill_color(220, 255, 220) # Verde
                        pdf_out.set_font("helvetica", 'B', 9)
                        pdf_out.cell(0, 8, "CONCILIADO", 1, fill=True, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    else:
                        status_str = f"❌ {len(divergencias)} Divergência(s)"
                        pdf_out.set_fill_color(255, 220, 220) # Vermelho
                        pdf_out.set_font("helvetica", 'B', 9)
                        pdf_out.cell(0, 8, "DIVERGÊNCIAS ENCONTRADAS:", 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                        
                        # Tabela Erros
                        pdf_out.set_fill_color(250, 250, 250)
                        pdf_out.set_font("helvetica", 'B', 8)
                        pdf_out.cell(20, 6, "Grupo", 1, fill=True, align='C')
                        pdf_out.cell(56, 6, "Saldo Relat.", 1, fill=True, align='C')
                        pdf_out.cell(56, 6, "Saldo SIAFI", 1, fill=True, align='C')
                        pdf_out.cell(57, 6, "Diferença", 1, fill=True, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                        
                        pdf_out.set_font("helvetica", '', 8)
                        for d in divergencias:
                            pdf_out.cell(20, 6, str(d['grupo']), 1, align='C')
                            pdf_out.cell(56, 6, formatar_real(d['pdf']), 1, align='R')
                            pdf_out.cell(56, 6, formatar_real(d['excel']), 1, align='R')
                            pdf_out.set_text_color(200, 0, 0)
                            pdf_out.cell(57, 6, formatar_real(d['diff']), 1, align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf_out.set_text_color(0, 0, 0)
                            
                    pdf_out.ln(5)
                    pdf_out.cell(0, 0, "", "B", new_x=XPos.LMARGIN, new_y=YPos.NEXT) # Linha divisória
                    pdf_out.ln(5)
                    
                    lista_resumo.append({
                        "Unidade": uid,
                        "Status": status_str,
                        "Diferença Total": f"R$ {formatar_real(dif_total)}"
                    })
                    
                    progresso.progress((idx + 1) / len(pares))
                
                progresso.empty()
                status_box.success("Processamento Finalizado com Sucesso!")
                
                # --- RESULTADOS NA TELA ---
                st.markdown("### Resumo da Conferência")
                st.dataframe(pd.DataFrame(lista_resumo), use_container_width=True)
                
                # Botão Download
                pdf_bytes = pdf_out.output(dest='S').encode('latin-1')
                st.download_button(
                    label="📥 Baixar Relatório Consolidado (PDF)",
                    data=pdf_bytes,
                    file_name="Relatorio_Depreciacao_Consolidado.pdf",
                    mime="application/pdf",
                    type="primary"
                )
