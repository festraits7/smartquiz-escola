import streamlit as st
from docx import Document
import pandas as pd
import random
import os
import time
from streamlit_gsheets import GSheetsConnection

# Configuração da página para o tema Dark/Cyberpunk
st.set_page_config(page_title="SmartQuiz Escola", page_icon="⚡", layout="centered")

# --- AS 88 LINHAS EXATAS DE ESTILIZAÇÃO CSS HACKER E EFEITOS NEON ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Fira+Code:wght@400;700&display=swap');

    body {
        background-color: #050a05 !important;
    }

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
        opacity: 0.4;
        pointer-events: none;
        background-size: 100% 4px, 6px 100%;
    }

    h1, h2, h3, h4, h5, h6 {
        color: #39FF14 !important;
        text-shadow: 0 0 12px rgba(57, 255, 20, 0.8) !important;
        font-family: 'Fira Code', monospace !important;
        font-weight: 700 !important;
    }

    p, label, span, div {
        color: #39FF14 !important;
        font-family: 'Fira Code', monospace !important;
        text-shadow: 0 0 4px rgba(57, 255, 20, 0.4);
    }
    
    .stTextInput>div>div>input {
        background-color: #000000 !important;
        color: #39FF14 !important;
        border: 2px solid #39FF14 !important;
        box-shadow: 0 0 10px rgba(57, 255, 20, 0.5) !important;
        font-family: 'Fira Code', monospace !important;
        border-radius: 4px !important;
    }

    .stTextInput>div>div>input:focus {
        border-color: #ffffff !important;
        box-shadow: 0 0 15px rgba(255, 255, 255, 0.6) !important;
    }
    
    .stButton>button {
        background-color: #021102 !important;
        color: #39FF14 !important;
        border: 2px solid #39FF14 !important;
        border-radius: 6px !important;
        font-weight: bold !important;
        font-family: 'Fira Code', monospace !important;
        letter-spacing: 2px !important;
        transition: all 0.3s ease-in-out !important;
        box-shadow: 0 0 8px rgba(57, 255, 20, 0.3) !important;
    }
    
    .stButton>button:hover {
        background-color: #39FF14 !important;
        color: #000000 !important;
        box-shadow: 0 0 25px #39FF14, 0 0 40px rgba(57, 255, 20, 0.6) !important;
        transform: scale(1.02);
    }

    div[data-testid="stRadio"] div[role="radiogroup"] > label {
        background-color: rgba(0, 5, 0, 0.8) !important;
        padding: 14px 24px !important;
        border: 1px solid rgba(57, 255, 20, 0.3) !important;
        border-radius: 8px !important;
        margin-bottom: 10px !important;
        transition: cubic-bezier(0.1, 0.7, 0.1, 1) 0.2s;
    }

    div[data-testid="stRadio"] div[role="radiogroup"] > label:hover {
        border-color: #39FF14 !important;
        background-color: rgba(57, 255, 20, 0.15) !important;
        box-shadow: 0 0 12px rgba(57, 255, 20, 0.3);
    }

    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #1f771f 0%, #39FF14 100%) !important;
        box-shadow: 0 0 10px #39FF14;
    }

    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception:
    conn = None

def carregar_perguntas():
    nome_arquivo = "pergunta.docx"
    perguntas = []
    
    # Se o arquivo existir, tenta ler
    if os.path.exists(nome_arquivo):
        try:
            doc = Document(nome_arquivo)
            for paragrafo in doc.paragraphs:
                texto = paragrafo.text.strip()
                if texto and "-" in texto:
                    partes = texto.split("-")
                    if len(partes) >= 3:
                        pergunta = partes[0].strip()
                        opcoes = [o.strip() for o in partes[1:-1]]
                        correta = partes[-1].strip()
                        
                        opcoes_misturadas = opcoes.copy()
                        random.shuffle(opcoes_misturadas)
                        
                        perguntas.append({
                            "pergunta": pergunta,
                            "opcoes": opcoes_misturadas,
                            "correta": correta
                        })
        except Exception:
            pass

    # SE TUDO FALHAR: Gera perguntas padrão automáticas para o sistema nunca quebrar
    if not perguntas:
        perguntas = [
            {
                "pergunta": "Qual a sintaxe correta para exibir texto em Python?",
                "opcoes": ["print('Olá')", "echo 'Olá'", "system.out.print('Olá')", "console.log('Olá')"],
                "correta": "print('Olá')"
            },
            {
                "pergunta": "Qual tag HTML é usada para criar uma quebra de linha?",
                "opcoes": ["<br>", "<lb>", "<break>", "<p>"],
                "correta": "<br>"
            }
        ]
    return perguntas

def buscar_ranking():
    if conn:
        try:
            df = conn.read(spreadsheet=st.secrets.get("connections", {}).get("gsheets", {}).get("spreadsheet"), ttl="5s")
            if not df.empty:
                return df.sort_values(by="Pontuacao", ascending=False).head(5)
        except Exception:
            pass
    return pd.DataFrame()

if "perguntas" not in st.session_state or len(st.session_state.perguntas) == 0:
    st.session_state.perguntas = carregar_perguntas()
if "logado" not in st.session_state:
    st.session_state.logado = False
if "usuario_nome" not in st.session_state:
    st.session_state.usuario_nome = ""
if "indice_pergunta" not in st.session_state:
    st.session_state.indice_pergunta = 0
if "pontuacao" not in st.session_state:
    st.session_state.pontuacao = 0
if "tempo_inicial" not in st.session_state:
    st.session_state.tempo_inicial = time.time()

# --- INTERFACE DE LOGIN ---
if not st.session_state.logado:
    st.title("⚡ SYSTEM INTERFACE: ONLINE QUIZ")
    st.write("➔ Estabelecendo conexão segura... Insira seu ID de operador.")
    
    usuario = st.text_input("ENTER USERNAME:")
    
    if st.button("INITIALIZE CORE SYSTEM"):
        if usuario.strip():
            with st.spinner("Sincronizando módulos virtuais..."):
                time.sleep(1.2)
            st.session_state.logado = True
            st.session_state.usuario_nome = usuario.strip()
            st.session_state.perguntas = carregar_perguntas()
            st.session_state.tempo_inicial = time.time()
            st.rerun()
        else:
            st.error("ERRO: IDENTIFICAÇÃO EM BRANCO.")

# --- EXECUÇÃO DO QUIZ ---
else:
    if st.session_state.indice_pergunta < len(st.session_state.perguntas):
        st.title("📟 TERMINAL OPERACIONAL")
        st.write(f"Operador Ativo: `[ {st.session_state.usuario_nome} ]` ➔ Link Estável.")
        st.markdown("---")
        
        q_atual = st.session_state.perguntas[st.session_state.indice_pergunta]
        st.subheader(f"Módulo {st.session_state.indice_pergunta + 1} de {len(st.session_state.perguntas)}")
        st.write(f"**>> {q_atual['pergunta']}**")
        
        tempo_limite = 30
        passado = time.time() - st.session_state.tempo_inicial
        restante = max(0, int(tempo_limite - passado))
        
        st.progress(restante / tempo_limite)
        st.write(f"⏱ Tempo restante para transmissão: `{restante}s`")
        
        if restante <= 0:
            st.error("🚨 SENSOR ESTOUROU: Tempo esgotado!")
            time.sleep(1.5)
            st.session_state.indice_pergunta += 1
            st.session_state.tempo_inicial = time.time()
            st.rerun()
            
        resposta = st.radio("Selecione o vetor de dados:", q_atual["opcoes"], key=f"r_{st.session_state.indice_pergunta}")
        
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("PROCESSAR VETOR DE RESPOSTA"):
            if resposta == q_atual["correta"]:
                st.session_state.pontuacao += 1
                st.success("✔ COMPILADO: Bit correspondente validado.")
            else:
                st.error(f"❌ ANOMALIA: Incorreto. Setor esperado era: {q_atual['correta']}")
                
            time.sleep(1.5)
            st.session_state.indice_pergunta += 1
            st.session_state.tempo_inicial = time.time()
            st.rerun()
            
        time.sleep(1)
        st.rerun()
        
    else:
        st.title("🎉 MISSÃO CONCLUÍDA 🎉")
        st.write("Todas as etapas foram descriptografadas com sucesso.")
        st.markdown(f"### Desempenho Final: `{st.session_state.pontuacao}` acertos de `{len(st.session_state.perguntas)}` módulos.")
        
        if conn and st.session_state.usuario_nome:
            try:
                dados = pd.DataFrame([{"Aluno": st.session_state.usuario_nome, "Pontuacao": st.session_state.pontuacao, "Total": len(st.session_state.perguntas)}])
                conn.create(data=dados)
                st.success("📊 Registro enviado para a central do professor!")
            except:
                pass
                
        st.markdown("---")
        st.subheader("🏆 TOP 5 OPERADORES (RANKING)")
        df_r = buscar_ranking()
        if not df_r.empty:
            st.dataframe(df_r, use_container_width=True)
        else:
            st.write("*Carregando dados globais da tabela do professor...*")
            
        if st.button("REINICIALIZAR TERMINAL"):
            st.session_state.indice_pergunta = 0
            st.session_state.pontuacao = 0
            st.session_state.perguntas = carregar_perguntas()
            st.session_state.tempo_inicial = time.time()
            st.rerun()
