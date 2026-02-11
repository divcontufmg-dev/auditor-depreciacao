import streamlit as st
import pandas as pd
import pdfplumber
import re
from fpdf import FPDF
import io

# ==========================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Conciliador de Depreciação",
    page_icon="📉",
    layout="wide"
)

st.title("📉 Conciliador Automático de Depreciação")
st.markdown("""
**Instruções:**
1. Faça o upload de **todos** os arquivos PDF (Relatórios de Depreciação).
2. Faça o upload de **todos** os arquivos Excel/CSV (SIAFI).
3. O sistema cruzará automaticamente os arquivos pelo código da Unidade (início do nome do arquivo).
""")

# ==========================================
# 1. FUNÇÕES DE LIMPEZA E EXTRAÇÃO
# ==========================================

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

def fmt_br(valor):
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ==========================================
# 2. PROCESSAMENTO DE ARQUIVOS (ADAPTADO PARA STREAMLIT)
# ==========================================

def processar_pdf(arquivo_obj):
    dados_pdf = {}
    texto_completo = ""
    try:
        with pdfplumber.open(arquivo_obj) as pdf:
            for page in pdf.pages: texto_completo += page.extract_text() + "\n"
    except Exception as e:
        st.error(f"Erro ao ler PDF: {e}")
        return {}

    regex_cabecalho = re.compile(r"(?m)^\s*(\d+)\s*-\s*[A-Z]")
    matches = list(regex_cabecalho.finditer(texto_completo))
    
    for i, match in enumerate(matches):
        grupo_id = int(match.group(1))
        start_idx = match.start()
        end_idx = matches[i+1].start() if i + 1 < len(matches) else len(texto_completo)
        bloco_texto = texto_completo[start_idx:end_idx]
        
        regex_saldo = re.compile(r"\(\*\)\s*SALDO[\s\S]*?ATUAL[\s\S]*?((?:\d{1,3}(?:\.\d{3})*,\d{2}))")
        match_saldo = regex_saldo.search(bloco_texto)
        dados_pdf[grupo_id] = formatar_moeda_pdf(match_saldo.group(1)) if match_saldo else 0.0
            
    return dados_pdf

def processar_excel(arquivo_obj):
    # Pandas lê direto do objeto file do Streamlit
    try:
        df = pd.read_csv(arquivo_obj, sep=',', encoding='latin1', header=None)
    except:
        try: 
            arquivo_obj.seek(0) # Reseta ponteiro se a leitura anterior falhou
            df = pd.read_excel(arquivo_obj, header=None)
        except: return {}

    linha_cabecalho = -1
    for i, row in df.iterrows():
        if "Nat Desp" in " ".join([str(x) for x in row.values]):
            linha_cabecalho = i; break
            
    if linha_cabecalho == -1: return {}
    
    # Recarregar com header correto
    arquivo_obj.seek(0)
    try:
        # Verifica extensão pelo nome do arquivo original
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
# 3. GERAÇÃO DE PDF
# ==========================================

class PDFConsolidado(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 10)
        self.cell(0, 10, 'Relatório de Conciliação - Depreciação Acumulada', 0, 1, 'C')
        self.line(10, 20, 200, 20)
        self.ln(10)
        
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

# ==========================================
# 4. INTERFACE DO USUÁRIO
# ==========================================

col1, col2 = st.columns(2)

with col1:
    arquivos_pdf = st.file_uploader("📂 Carregar PDFs (Relatórios)", type=["pdf"], accept_multiple_files=True)

with col2:
    arquivos_excel = st.file_uploader("📊 Carregar Excel/CSV (SIAFI)", type=["xlsx", "csv"], accept_multiple_files=True)

if st.button("🚀 Processar Conciliação"):
    if not arquivos_pdf or not arquivos_excel:
        st.warning("Por favor, carregue pelo menos um arquivo PDF e um arquivo Excel/CSV.")
    else:
        # Mapeamento
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

        if not unidades:
            st.error("Nenhum par de arquivos correspondente foi encontrado. Verifique se os nomes dos arquivos começam com o código da Unidade.")
        else:
            # Barra de progresso
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Inicializa PDF
            pdf = PDFConsolidado()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()
            
            lista_unidades = sorted(unidades.keys())
            
            # Dados para mostrar na tela (preview)
            resumo_tela = []

            for idx, uid in enumerate(lista_unidades):
                status_text.text(f"Processando Unidade {uid}...")
                arquivos = unidades[uid]
                
                if 'pdf' in arquivos and 'excel' in arquivos:
                    # Resetar ponteiros de arquivo (boa prática no streamlit)
                    arquivos['pdf'].seek(0)
                    arquivos['excel'].seek(0)
                    
                    d_pdf = processar_pdf(arquivos['pdf'])
                    d_excel = processar_excel(arquivos['excel'])
                    
                    grupos = sorted(list(set(d_pdf.keys()) | set(d_excel.keys())))
                    lista_divergencias = []
                    t_pdf, t_excel = 0.0, 0.0
                    
                    for g in grupos:
                        vp = d_pdf.get(g, 0.0)
                        ve = d_excel.get(g, 0.0)
                        t_pdf += vp
                        t_excel += ve
                        diff = vp - ve
                        if abs(diff) > 0.10: 
                            lista_divergencias.append({'grupo': g, 'v_pdf': vp, 'v_excel': ve, 'diff': diff})

                    # --- ESCRITA NO PDF ---
                    if pdf.get_y() > 240: pdf.add_page()

                    pdf.set_font("Arial", 'B', 12)
                    pdf.set_fill_color(230, 230, 230)
                    pdf.cell(0, 8, f"Unidade Gestora: {uid}", 0, 1, 'L', fill=True)
                    pdf.ln(2)
                    
                    # Totais
                    pdf.set_font("Arial", 'B', 9)
                    pdf.cell(60, 6, "Total Relatório", 1, 0, 'C')
                    pdf.cell(60, 6, "Total SIAFI", 1, 0, 'C')
                    pdf.cell(60, 6, "Diferença", 1, 1, 'C')
                    
                    pdf.set_font("Arial", size=9)
                    pdf.cell(60, 6, f"R$ {fmt_br(t_pdf)}", 1, 0, 'C')
                    pdf.cell(60, 6, f"R$ {fmt_br(t_excel)}", 1, 0, 'C')
                    
                    dif_total = t_pdf - t_excel
                    if abs(dif_total) > 0.10: pdf.set_text_color(200, 0, 0)
                    else: pdf.set_text_color(0, 100, 0)
                    
                    pdf.cell(60, 6, f"R$ {fmt_br(dif_total)}", 1, 1, 'C')
                    pdf.set_text_color(0, 0, 0)
                    pdf.ln(4)
                    
                    # Divergências
                    status_str = "✅ Conciliado"
                    if not lista_divergencias:
                        pdf.set_fill_color(220, 255, 220)
                        pdf.set_font("Arial", 'B', 9)
                        pdf.cell(0, 8, "CONCILIADO", 1, 1, 'C', fill=True)
                    else:
                        status_str = f"❌ {len(lista_divergencias)} divergência(s)"
                        pdf.set_fill_color(255, 220, 220)
                        pdf.set_font("Arial", 'B', 9)
                        pdf.cell(0, 8, "DIVERGÊNCIAS ENCONTRADAS:", 1, 1, 'L', fill=True)
                        
                        pdf.set_font("Arial", 'B', 8)
                        pdf.cell(20, 6, "Grupo", 1, 0, 'C')
                        pdf.cell(45, 6, "Relatório", 1, 0, 'C')
                        pdf.cell(45, 6, "SIAFI", 1, 0, 'C')
                        pdf.cell(40, 6, "Diferença", 1, 1, 'C')
                        
                        pdf.set_font("Arial", size=8)
                        for d in lista_divergencias:
                            pdf.cell(20, 6, str(d['grupo']), 1, 0, 'C')
                            pdf.cell(45, 6, fmt_br(d['v_pdf']), 1, 0, 'R')
                            pdf.cell(45, 6, fmt_br(d['v_excel']), 1, 0, 'R')
                            pdf.set_text_color(200, 0, 0)
                            pdf.cell(40, 6, fmt_br(d['diff']), 1, 1, 'R')
                            pdf.set_text_color(0, 0, 0)

                    pdf.ln(8)
                    pdf.cell(0, 0, "", "B", 1, 'C') # Linha separadora
                    pdf.ln(8)
                    
                    # Adicionar ao resumo da tela
                    resumo_tela.append({
                        "Unidade": uid,
                        "Status": status_str,
                        "Diferença Total": f"R$ {fmt_br(dif_total)}"
                    })

                progress_bar.progress((idx + 1) / len(lista_unidades))
            
            # Finalização
            progress_bar.empty()
            status_text.success("Processamento concluído!")
            
            # Mostrar tabela resumo na tela
            st.subheader("Resumo da Análise")
            df_resumo = pd.DataFrame(resumo_tela)
            st.dataframe(df_resumo, use_container_width=True)
            
            # Gerar binário do PDF para download
            pdf_bytes = pdf.output(dest='S').encode('latin-1')
            
            st.download_button(
                label="📥 Baixar Relatório Consolidado (PDF)",
                data=pdf_bytes,
                file_name="Relatorio_Depreciacao_Consolidado.pdf",
                mime="application/pdf"
            )
