import streamlit as st
import pandas as pd
from docx import Document
import re
import io
import zipfile
import requests
import json
import time

# --- CONFIGURAÇÕES PADRÃO ---
DEFAULT_EXCEL = "../dadosNov.xlsx"
DEFAULT_TXT = "../prompt_dupar.txt"

# --- FUNÇÕES AUXILIARES ---
def ler_e_substituir_template(template_text, dados):
    """Substitui variáveis no texto do template."""
    def substituir(match):
        variavel = match.group(1).strip()
        valor = dados.get(variavel)
        return str(valor).strip() if valor is not None else ""
    return re.sub(r'\{\{([^{}]+)\}\}', substituir, template_text)

def formatar_paragrafo_com_negrito(paragrafo, texto):
    """
    Processa o texto procurando por marcadores de negrito markdown (**texto**).
    Adiciona runs ao parágrafo com a formatação correta.
    """
    # Divide o texto pelos marcadores de negrito
    partes = re.split(r'(\*\*.*?\*\*)', texto)
    for parte in partes:
        if parte.startswith('**') and parte.endswith('**'):
            # Remove os asteriscos e adiciona como negrito
            run = paragrafo.add_run(parte[2:-2])
            run.bold = True
        else:
            # Adiciona texto normal
            paragrafo.add_run(parte)

def criar_docx_bytes(texto_resposta):
    """
    Converte a string de resposta (Markdown do GPT) em um arquivo .docx binário.
    Tenta interpretar títulos (#), listas (-) e negrito (**).
    """
    doc = Document()
    
    # Divide o texto por linhas para processar formatação
    linhas = texto_resposta.split('\n')
    
    for linha in linhas:
        linha = linha.strip()
        if not linha:
            continue
            
        # Verifica cabeçalhos (Markdown #, ##, etc)
        if linha.startswith('#'):
            nivel = linha.count('#')
            texto_limpo = linha.lstrip('#').strip()
            # O Word suporta níveis de 1 a 9
            doc.add_heading(texto_limpo, level=min(nivel, 9))
            
        # Verifica listas (Markdown - ou *)
        elif linha.startswith('- ') or linha.startswith('* '):
            p = doc.add_paragraph(style='List Bullet')
            texto_limpo = linha[2:]
            formatar_paragrafo_com_negrito(p, texto_limpo)
            
        # Parágrafo normal
        else:
            p = doc.add_paragraph()
            formatar_paragrafo_com_negrito(p, linha)
    
    # Salva o documento em memória
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.read()

def chamar_gpt(api_key, prompt_text, modelo="gpt-5.2"):
    """Envia o prompt para a API da OpenAI e retorna a resposta."""
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    data = {
        "model": modelo,
        "messages": [{"role": "user", "content": prompt_text}],
        "temperature": 0.7
    }
    
    try:
        # Retry simples em caso de erro de conexão momentâneo
        for tentativa in range(3):
            try:
                response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=data, timeout=60)
                response.raise_for_status() 
                return response.json()['choices'][0]['message']['content']
            except requests.exceptions.HTTPError as e:
                if response.status_code == 429: # Rate limit
                    time.sleep(2 * (tentativa + 1))
                    continue
                return f"Erro na API (HTTP {response.status_code}): {response.text}"
            except Exception:
                if tentativa == 2: raise
                time.sleep(1)
                
    except Exception as e:
        return f"Erro fatal na conexão: {e}"

# --- INTERFACE STREAMLIT ---
st.set_page_config(page_title="Sistema DUPAR + GPT", page_icon="🤖", layout="wide")

st.title("🤖 Automação de Relatórios em Lote")
st.markdown("### 1. Preparar Prompts -> 2. Executar Fila GPT")

# --- ESTADO (SESSION STATE) ---
if 'processamento_concluido' not in st.session_state:
    st.session_state.processamento_concluido = False
if 'zip_prompts' not in st.session_state:
    st.session_state.zip_prompts = None
if 'todos_prompts' not in st.session_state:
    st.session_state.todos_prompts = [] 
if 'respostas_geradas' not in st.session_state:
    st.session_state.respostas_geradas = [] # Lista para guardar as respostas
if 'zip_respostas' not in st.session_state:
    st.session_state.zip_respostas = None

# 1. SIDEBAR - Configuração e API KEY
st.sidebar.header("🔑 Configuração")
api_key = st.sidebar.text_input("OpenAI API Key", type="password", help="Cole sua chave sk-... aqui")

modelo_gpt = st.sidebar.selectbox("Modelo GPT", ["gpt-3.5-turbo", "gpt-4", "gpt-4o", "gpt-5.2"])

st.sidebar.divider()
st.sidebar.header("📁 Arquivos")
uploaded_excel = st.sidebar.file_uploader("Carregar Planilha (.xlsx)", type=["xlsx"])
uploaded_template = st.sidebar.file_uploader("Carregar Template (.txt)", type=["txt"])

df = None
template_content = ""

# Carregamento de Arquivos
try:
    if uploaded_excel:
        df = pd.read_excel(uploaded_excel, sheet_name="dupar")
    else:
        try:
            df = pd.read_excel(DEFAULT_EXCEL, sheet_name="dupar")
            st.sidebar.info(f"Usando planilha padrão.")
        except:
            pass 

    if uploaded_template:
        template_content = uploaded_template.read().decode("utf-8")
    else:
        try:
            with open(DEFAULT_TXT, "r", encoding="utf-8") as f:
                template_content = f.read()
            st.sidebar.info(f"Usando template padrão.")
        except:
            pass

except Exception as e:
    st.error(f"Erro ao carregar: {e}")

# 2. ÁREA PRINCIPAL
if df is not None and template_content:
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.subheader("⚙️ Seleção de Dados")
        todas_colunas = df.columns.tolist()
        
        colunas_selecionadas = st.multiselect(
            "Colunas para incluir no prompt:",
            options=todas_colunas,
            default=todas_colunas 
        )
        
        st.info(f"Registros encontrados: {len(df)}")
        
        # Botão: Preparar Dados
        if st.button("📝 1. Preparar Prompts", type="primary"):
            if not colunas_selecionadas:
                st.error("Selecione ao menos uma coluna.")
            else:
                zip_buffer = io.BytesIO()
                prompts_gerados = [] 

                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    progress_bar = st.progress(0)
                    
                    for i, row in df.iterrows():
                        dados_filtrados = {col: row[col] for col in colunas_selecionadas}
                        
                        conteudo_prompt = ler_e_substituir_template(template_content, dados_filtrados)
                        
                        nome_pessoa = str(row.get('nome_completo', f'Registro {i+1}'))
                        nome_limpo = re.sub(r'[^a-zA-Z0-9\s]', '', nome_pessoa).replace(' ', '_')
                        
                        zf.writestr(f"{nome_limpo}_prompt.txt", conteudo_prompt)
                        
                        prompts_gerados.append({
                            "id": i,
                            "nome": nome_pessoa,
                            "nome_arquivo": nome_limpo,
                            "conteudo": conteudo_prompt
                        })
                        
                        progress_bar.progress((i + 1) / len(df))

                zip_buffer.seek(0)
                
                # Salva no estado
                st.session_state.processamento_concluido = True
                st.session_state.zip_prompts = zip_buffer
                st.session_state.todos_prompts = prompts_gerados
                st.session_state.respostas_geradas = [] 
                st.session_state.zip_respostas = None
                st.rerun()

    with col2:
        st.subheader("📊 Visualização da Planilha")
        
        if colunas_selecionadas:
            df_visualizacao = df[colunas_selecionadas]
        else:
            df_visualizacao = df
            
        st.dataframe(
            df_visualizacao, 
            use_container_width=True, 
            height=300, 
            hide_index=True 
        )

    # --- ÁREA DE EXECUÇÃO ---
    if st.session_state.processamento_concluido:
        st.divider()
        
        col_prompts, col_gpt = st.columns([1, 1])
        
        # --- COLUNA DA ESQUERDA: PROMPTS ---
        with col_prompts:
            st.subheader("1. Fila de Prompts")
            st.caption(f"Total: {len(st.session_state.todos_prompts)} itens")
            
            st.download_button(
                "⬇️ Baixar Prompts (.zip)",
                data=st.session_state.zip_prompts,
                file_name="prompts.zip",
                mime="application/zip"
            )
            
            with st.container(height=600):
                for item in st.session_state.todos_prompts:
                    with st.expander(f"📄 {item['nome']}", expanded=False):
                        st.text_area("Prompt:", value=item['conteudo'], height=150, key=f"p_{item['id']}")

        # --- COLUNA DA DIREITA: RESPOSTAS (PROCESSAMENTO) ---
        with col_gpt:
            st.subheader("2. Respostas da IA")
            
            # Botão de Ação Principal
            if st.button("🚀 2. Processar Fila (Enviar Todos)", type="primary"):
                if not api_key:
                    st.warning("⚠️ Insira a API Key na barra lateral antes de começar.")
                else:
                    st.session_state.respostas_geradas = [] 
                    
                    status_bar = st.progress(0)
                    status_text = st.empty()
                    
                    zip_respostas_io = io.BytesIO()
                    
                    with zipfile.ZipFile(zip_respostas_io, "w") as zf_resp:
                        
                        total = len(st.session_state.todos_prompts)
                        
                        for idx, item in enumerate(st.session_state.todos_prompts):
                            status_text.text(f"Processando {idx+1}/{total}: {item['nome']}...")
                            
                            # 1. Obter resposta
                            resposta_texto = chamar_gpt(api_key, item['conteudo'], modelo_gpt)
                            
                            # 2. Salvar na lista visual
                            st.session_state.respostas_geradas.append({
                                "nome": item['nome'],
                                "resposta": resposta_texto
                            })
                            
                            # 3. Converter para DOCX (COM FORMATAÇÃO)
                            arquivo_docx = criar_docx_bytes(resposta_texto)
                            
                            # 4. Escrever no ZIP como .docx
                            zf_resp.writestr(f"RESPOSTA_{item['nome_arquivo']}.docx", arquivo_docx)
                            
                            status_bar.progress((idx + 1) / total)
                    
                    zip_respostas_io.seek(0)
                    st.session_state.zip_respostas = zip_respostas_io
                    status_text.success("Processamento concluído!")
                    st.rerun()

            # --- EXIBIÇÃO DAS RESPOSTAS ---
            if st.session_state.zip_respostas:
                st.download_button(
                    "⬇️ Baixar Todas as Respostas (.zip)",
                    data=st.session_state.zip_respostas,
                    file_name="respostas_gpt_docx.zip", # Alterado nome para indicar docx
                    mime="application/zip",
                    type="primary"
                )
            
            st.caption(f"Respostas geradas: {len(st.session_state.respostas_geradas)}")
            with st.container(height=600):
                if not st.session_state.respostas_geradas:
                    st.info("Nenhuma resposta gerada ainda. Clique no botão acima.")
                else:
                    for resp in st.session_state.respostas_geradas:
                        icone = "✅"
                        if "Erro" in resp['resposta']:
                            icone = "❌"
                            
                        with st.expander(f"{icone} {resp['nome']}"):
                            st.markdown(resp['resposta']) # Markdown visual na tela

elif df is None:
    st.info("Aguardando arquivos...")