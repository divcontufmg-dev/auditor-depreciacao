import streamlit as st
import pandas as pd
import pdfplumber
import re
from fpdf import FPDF, XPos, YPos
import io
import os

# ==========================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Conciliador de Depreciação",
    page_icon="📉",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilos CSS para manter a identidade visual limpa
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            .stFileUploader {
                padding-top: 2rem;
            }
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ==========================================
# FUNÇÕES DE LÓGICA (EXTRAÇÃO E CONVERSÃO)
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
        # Tenta ler como CSV primeiro
        df = pd.read_csv(arquivo_obj, sep=',', encoding='latin1', header=None)
    except:
        try: 
            arquivo_obj.seek(0)
            df = pd.read_excel(arquivo_obj, header=None)
        except: return {}

    # Localiza linha de cabeçalho
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
# CLASSE PDF (Visual Ajustado)
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
# INTERFACE PRINCIPAL
# ==========================================

# Sidebar com instruções
with st.sidebar:
    st.header("Instruções")
    st.markdown("""
    **Como usar:**
    1.  Arraste **todos os arquivos** (PDFs e Excel/CSVs) para a área de upload.
    2.  O sistema separará automaticamente os tipos de arquivo.
    3.  O cruzamento é feito pelo **código da unidade** no início do nome.
    
    **Exemplo:**
    * `153289.pdf` (Relatório)
    * `153289_SIAFI.xlsx` (Dados)
    * O sistema identifica "153289" e cruza os dados.
    """)
    st.markdown("---")
    st.markdown("**Versão:** 2.1 (Upload Unificado)")

# Área Central
st.title("📉 Conciliador Automático de Depreciação")
st.markdown("Faça o upload de todos os arquivos (Relatórios PDF e Planilhas SIAFI) de uma só vez.")

# --- UPLOAD UNIFICADO ---
arquivos_upload = st.file_uploader(
    "Arraste todos os arquivos aqui (PDF, Excel, CSV)", 
    type=["pdf", "xlsx", "csv"], 
    accept_multiple_files=True
)

if st.button("🚀 Processar Conciliação", type="primary"):
    if not arquivos_upload:
        st.warning("⚠️ Nenhum arquivo carregado.")
    else:
        # Separação automática dos arquivos
        arquivos_pdf = [f for f in arquivos_upload if f.name.lower().endswith('.pdf')]
        arquivos_excel = [f for f in arquivos_upload if f.name.lower().endswith(('.xlsx', '.csv'))]
        
        st.info(f"Arquivos identificados: {len(arquivos_pdf)} Relatórios PDF e {len(arquivos_excel)} Planilhas SIAFI.")
        
        if not arquivos_pdf or not arquivos_excel:
            st.error("❌ É necessário pelo menos 1 PDF e 1 Excel/CSV para prosseguir.")
        else:
            # Agrupamento por Unidade
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

            # Filtra apenas os pares completos
            pares_validos = [u for u, arqs in unidades.items() if 'pdf' in arqs and 'excel' in arqs]

            if not pares_validos:
                st.error("❌ Nenhum par correspondente encontrado (Nomes não batem).")
            else:
                progresso = st.progress(0)
                status_text = st.empty()
                
                # Setup do PDF
                pdf_out = PDFRelatorio()
                pdf_out.set_auto_page_break(auto=True, margin=15)
                pdf_out.add_page()
                
                resumo_geral = []

                for idx, uid in enumerate(sorted(pares_validos)):
                    status_text.text(f"Analisando Unidade: {uid}...")
                    arqs = unidades[uid]
                    
                    # Reset ponteiros
                    arqs['pdf'].seek(0)
                    arqs['excel'].seek(0)
                    
                    # Extração
                    d_pdf = processar_pdf(arqs['pdf'])
                    d_excel = processar_excel(arqs['excel'])
                    
                    # Comparação
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
                        if abs(diff) > 0.10:
                            divergencias.append({'grupo': g, 'pdf': v_p, 'excel': v_e, 'diff': diff})

                    # --- GERAÇÃO DO RELATÓRIO PDF ---
                    if pdf_out.get_y() > 240: pdf_out.add_page()

                    # Cabeçalho Unidade
                    pdf_out.set_font("Helvetica", 'B', 11)
                    pdf_out.set_fill_color(240, 240, 240)
                    pdf_out.cell(0, 8, f"Unidade Gestora: {uid}", 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    pdf_out.ln(2)

                    # Tabela de Totais
                    pdf_out.set_font("Helvetica", 'B', 9)
                    pdf_out.set_fill_color(220, 230, 241) # Azul Identidade Visual
                    
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
                    pdf_out.set_text_color(0, 0, 0)
                    pdf_out.ln(3)

                    # Divergências
                    status_tela = "✅ OK"
                    if not divergencias:
                        pdf_out.set_fill_color(220, 255, 220)
                        pdf_out.set_font("Helvetica", 'B', 9)
                        pdf_out.cell(0, 8, "CONCILIADO", 1, fill=True, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    else:
                        status_tela = f"❌ {len(divergencias)} Erros"
                        pdf_out.set_fill_color(255, 220, 220)
                        pdf_out.set_font("Helvetica", 'B', 9)
                        pdf_out.cell(0, 8, "DIVERGÊNCIAS ENCONTRADAS:", 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                        
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
                    pdf_out.cell(0, 0, "", "B", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    pdf_out.ln(5)
                    
                    resumo_geral.append({
                        "Unidade": uid,
                        "Status": status_tela,
                        "Diferença Global": f"R$ {formatar_real(dif_total)}"
                    })
                    
                    progresso.progress((idx + 1) / len(pares_validos))

                progresso.empty()
                status_text.success("Processamento finalizado!")
                
                # Resumo e Download
                st.markdown("### Resumo da Análise")
                st.dataframe(pd.DataFrame(resumo_geral), use_container_width=True)
                
                pdf_bytes = pdf_out.output(dest='S').encode('latin-1')
                st.download_button(
                    label="📥 Baixar Relatório Consolidado (PDF)",
                    data=pdf_bytes,
                    file_name="Relatorio_Depreciacao_Consolidado.pdf",
                    mime="application/pdf",
                    type="primary"
                )
