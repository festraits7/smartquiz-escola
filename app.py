import streamlit as st
from docx import Document
import pandas as pd
import random
import os
import time
from streamlit_gsheets import GSheetsConnection

# Configuração da página para o tema Dark/Cyberpunk
st.set_page_config(page_title="SmartQuiz Escola", page_icon="⚡", layout="centered")

# --- TODOS OS EFEITOS VISUAIS E ESTILIZAÇÃO HACKER (CSS/HTML) ---
st.markdown("""
    <style>
    /* Importando fonte digital/terminal */
    @import url('https://fonts.googleapis.com/css2?family=Fira+Code:wght@400;700&display=swap');

    /* Fundo preto puro e animação de scanline (linhas de monitor antigo) */
    .stApp {
        background: radial-gradient(circle, #0d1f10 0%, #050a05 100%) !important;
        color: #39FF14 !important;
        font-family: 'Fira Code', monospace !important;
    }
    
    .stApp::before {
        content: " ";
        display: block;
        position: fixed;
        top: 0; left: 0; bottom: 0; right: 0;
        background: linear-gradient(rgba(18, 16, 16, 0) 50%, rgba(0, 0, 0, 0.25) 50%), linear-gradient(90deg, rgba(255, 0, 0, 0.06), rgba(0, 255, 0, 0.02), rgba(0, 0, 255, 0.06));
        z-index: 99999;
        opacity: 0.8;
        pointer-events: none;
        background-size: 100% 4px, 6px 100%;
    }

    /* Brilho Neon (Glow) nos Títulos e Textos */
    h1, h2, h3, p, label, .stMarkdown {
        color: #39FF14 !important;
        text-shadow: 0 0 8px rgba(57, 255, 20, 0.6);
    }
    
    /* Moldura estilo Terminal Hacker para os blocos */
    .stAlert, div[data-testid="stVerticalBlock"] > div {
        border-radius: 4px;
    }

    /* Caixas de Texto Estilizadas */
    .stTextInput>div>div>input {
        background-color: #000000 !important;
        color: #39FF14 !important;
        border: 1px solid #39FF14 !important;
        box-shadow: 0 0 5px rgba(57, 255, 20, 0.3);
        font-family: 'Fira Code', monospace !important;
    }
    
    /* Botões Digitais com Efeito Hover Retroiluminado */
    .stButton>button {
        background-color: #050a05 !important;
        color: #39FF14 !important;
        border: 1px solid #39FF14 !important;
        border-radius: 5px !important;
        font-weight: bold !important;
        font-family: 'Fira Code', monospace !important;
        letter-spacing: 2px;
        transition: all 0.4s ease;
        width: 100%;
    }
    
    .stButton>button:hover {
        background-color: #39FF14 !important;
        color: #000000 !important;
        box-shadow: 0 0 20px #39FF14, 0 0 40px rgba(57, 255, 20, 0.5) !important;
        transform: translateY(-2px);
    }

    /* Customização dos seletores de resposta (Radio Buttons) */
    div[data-testid="stRadio"] label {
        background-color: rgba(0, 0, 0, 0.6) !important;
        padding: 12px 20px !important;
        border: 1px solid rgba(57, 255, 20, 0.2) !important;
        border-radius: 5px !important;
        margin-bottom: 8px !important;
        transition: 0.3s;
        cursor: pointer;
    }

    div[data-testid="stRadio"] label:hover {
        border-color: #39FF14 !important;
        background-color: rgba(57, 255, 20, 0.1) !important;
        box-shadow: 0 0 10px rgba(57, 255, 20, 0.2);
    }

    /* Barra de progresso customizada verde neon */
    .stProgress > div > div > div > div {
        background-color: #39FF14 !important;
    }
    </style>
""", unsafe_allow_html=True)

# Conexão segura com o Google Sheets
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception:
    conn = None

# Carregamento inteligente das perguntas do arquivo Word
def carregar_perguntas():
    if not os.path.exists("perguntas.docx"):
        return []
        
    doc = Document("perguntas.docx")
    perguntas = []
    
    for paragrafo in doc.paragraphs:
        texto = paragrafo.text.strip()
        if texto and " - " in texto:
            partes = texto.split(" - ")
            if len(partes) >= 3:
                pergunta = partes[0].strip()
                opcoes = [o.strip() for o in partes[1:-1]]
                correta = partes[-1].strip()
                
                # Criar uma cópia embaralhada mantendo a referência correta
                opcoes_misturadas = opcoes.copy()
                random.shuffle(opcoes_misturadas)
                
                perguntas.append({
                    "pergunta": pergunta,
                    "opcoes": opcoes_misturadas,
                    "correta": correta
                })
    return perguntas

# Buscar banco de dados de notas para o ranking em tempo real
def buscar_ranking():
    if conn:
        try:
            df = conn.read(spreadsheet=st.secrets.get("connections", {}).get("gsheets", {}).get("spreadsheet"), ttl="5s")
            if not df.empty:
                df_sorted = df.sort_values(by="Pontuacao", ascending=False).head(5)
                return df_sorted
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()

# Inicialização dos estados do Quiz
if "logado" not in st.session_state:
    st.session_state.logado = False
if "usuario_nome" not in st.session_state:
    st.session_state.usuario_nome = ""
if "perguntas" not in st.session_state:
    st.session_state.perguntas = carregar_perguntas()
if "indice_pergunta" not in st.session_state:
    st.session_state.indice_pergunta = 0
if "pontuacao" not in st.session_state:
    st.session_state.pontuacao = 0
if "tempo_inicial" not in st.session_state:
    st.session_state.tempo_inicial = time.time()

# --- TELA DE LOGIN (ESTILO TERMINAL DE DEFESA) ---
if not st.session_state.logado:
    st.title("⚡ SYSTEM INTERFACE: ONLINE QUIZ")
    st.write("➔ Estabelecendo conexão segura... Insira sua identificação para descriptografar os dados.")
    
    usuario = st.text_input("ENTER USERNAME:")
    
    if st.button("INITIALIZE CORE SYSTEM"):
        if usuario.strip() != "":
            with st.spinner("Descriptografando banco de dados..."):
                time.sleep(1.2) # Efeito de carregamento hacker
            st.session_state.logado = True
            st.session_state.usuario_nome = usuario.strip()
            st.session_state.tempo_inicial = time.time()
            st.success("SISTEMA AUTENTICADO COM SUCESSO.")
            st.rerun()
        else:
            st.error("ERRO: IDENTIFICAÇÃO INVÁLIDA.")

# --- TELA PRINCIPAL DO QUIZ ANIMADO ---
else:
    if st.session_state.indice_pergunta < len(st.session_state.perguntas):
        st.title("📟 TERMINAL OPERACIONAL")
        st.write(f"Operador Ativo: `[ {st.session_state.usuario_nome} ]` ➔ Conexão Estável.")
        st.markdown("---")
        
        if not st.session_state.perguntas:
            st.warning("ALERTA: O arquivo 'perguntas.docx' não possui dados válidos mapeados.")
        else:
            q_atual = st.session_state.perguntas[st.session_state.indice_pergunta]
            
            st.subheader(f"Módulo de Questão {st.session_state.indice_pergunta + 1} de {len(st.session_state.perguntas)}")
            st.write(f"**>> {q_atual['pergunta']}**")
            
            # Cronômetro visual regressivo por questão (30 segundos)
            tempo_limite = 30
            tempo_passado = time.time() - st.session_state.tempo_inicial
            tempo_restante = max(0, int(tempo_limite - tempo_passado))
            
            progresso = tempo_restante / tempo_limite
            st.progress(progresso)
            st.write(f"⏱ Tempo restante para sincronização: `{tempo_restante}s`")
            
            if tempo_restante <= 0:
                st.error("🚨 TEMPO ESGOTADO! O sistema avançou automaticamente.")
                time.sleep(1.5)
                st.session_state.indice_pergunta += 1
                st.session_state.tempo_inicial = time.time()
                st.rerun()
            
            resposta = st.radio("Selecione o vetor de resposta correto:", q_atual["opcoes"], key=f"rad_{st.session_state.indice_pergunta}")
            
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("PROCESSAR VETOR DE RESPOSTA"):
                if resposta == q_atual["correta"]:
                    st.session_state.pontuacao += 1
                    st.success("✔ COMPILADO COM SUCESSO: Resposta Correta.")
                else:
                    st.error(f"❌ COMPILAÇÃO FALHOU: Resposta incorreta. Gabarito mapeado: {q_atual['correta']}")
                
                time.sleep(1.5) # Delay estratégico para o aluno ler o feedback visual antes de avançar
                st.session_state.indice_pergunta += 1
                st.session_state.tempo_inicial = time.time()
                st.rerun()
            
            # Auto-refresh controlado para atualizar o cronômetro na tela
            time.sleep(1)
            st.rerun()
            
    else:
        st.title("🎉 MISSÃO CONCLUÍDA 🎉")
        st.write("Todas as etapas foram descriptografadas com sucesso.")
        st.markdown(f"### Desempenho Final: `{st.session_state.pontuacao}` acertos de `{len(st.session_state.perguntas)}` módulos.")
        
        # Envio automático em background para o Google Sheets
        if conn:
            try:
                dados_novos = pd.DataFrame([{
                    "Aluno": st.session_state.usuario_nome,
                    "Pontuacao": st.session_state.pontuacao,
                    "Total": len(st.session_state.perguntas)
                }])
                conn.create(data=dados_novos)
                st.success("📊 Relatório de notas transmitido para o servidor central com sucesso!")
            except Exception:
                pass
        
        st.markdown("---")
        st.subheader("🏆 TOP 5 OPERADORES (RANKING)")
        
        df_ranking = buscar_ranking()
        if not df_ranking.empty:
            st.dataframe(df_ranking, use_container_width=True)
        else:
            st.write("*Carregando dados globais da tabela do professor...*")
        
        if st.button("REINICIALIZAR TERMINAL"):
            st.session_state.indice_pergunta = 0
            st.session_state.pontuacao = 0
            st.session_state.tempo_inicial = time.time()
            st.rerun()