import streamlit as st
from docx import Document
import os
import time
from datetime import datetime
from streamlit_gsheets import GSheetsConnection

# Configuração da página do site
st.set_page_config(page_title="Quiz Escolar Online", page_icon="🏫", layout="centered")

# --- ESTILIZAÇÃO PERSONALIZADA TECH/PROGRAMAÇÃO (CSS + MATRIX EFFECT) ---
st.markdown("""
    <style>
    /* Fundo escuro estilo terminal/hacker */
    .stApp {
        background-color: #0d1117;
        font-family: 'Courier New', Courier, monospace;
        color: #c9d1d9;
    }
    
    /* Efeito de linhas de código binário subindo bem suave no fundo */
    .stApp::before {
        content: "0101010101101010101010110101010101011010101010101101010101010110101010101011010101\\A 1010101011010101010101101010101010110101010101011010101010101101010101010110101010\\A 01101010101010110101010101011010101010101101010101010110101010101011010101010101101";
        position: fixed;
        top: 0; left: 0; width: 100%; height: 100%;
        font-size: 14px;
        color: rgba(0, 255, 102, 0.03); /* Verde Matrix bem transparente para não atrapalhar a leitura */
        white-space: pre-wrap;
        word-wrap: break-word;
        pointer-events: none;
        z-index: 0;
        line-height: 20px;
    }
    
    /* Customização do Título Tech */
    h1 {
        color: #00ff66 !important;
        text-align: center;
        font-weight: bold !important;
        text-shadow: 0 0 10px rgba(0, 255, 102, 0.5);
    }
    
    /* Cards brancos viraram painéis de comando escuros */
    .css-1r6slb0, .stForm, div[data-testid="stVerticalBlock"] > div {
        background-color: #161b22 !important;
        padding: 30px !important;
        border-radius: 12px !important;
        box-shadow: 0px 4px 20px rgba(0, 255, 102, 0.05) !important;
        border: 1px solid #30363d !important;
    }
    
    /* Pergunta destacada como um comando de terminal */
    .stAlert {
        background-color: #21262d !important;
        border-left: 5px solid #00ff66 !important;
        color: #00ff66 !important;
        border-radius: 6px !important;
        font-size: 1.15rem !important;
        font-weight: bold !important;
    }
    
    /* Botões estilo botão de ativação/energia */
    .stButton>button {
        background-color: #00ff66 !important;
        color: #0d1117 !important;
        border-radius: 6px !important;
        padding: 12px 24px !important;
        font-weight: bold !important;
        border: none !important;
        transition: all 0.3s ease !important;
        box-shadow: 0px 0px 10px rgba(0, 255, 102, 0.3) !important;
    }
    
    .stButton>button:hover {
        background-color: #00cc52 !important;
        box-shadow: 0px 0px 20px rgba(0, 255, 102, 0.6) !important;
        transform: scale(1.01);
    }
    
    /* Caixa das alternativas imitando blocos de código */
    div[data-testid="stRadio"] label {
        background-color: #21262d !important;
        color: #c9d1d9 !important;
        padding: 12px 16px !important;
        border-radius: 6px !important;
        margin-bottom: 8px !important;
        border: 1px solid #30363d !important;
        transition: all 0.2s ease !important;
    }
    div[data-testid="stRadio"] label:hover {
        background-color: #30363d !important;
        border-color: #00ff66 !important;
        color: #00ff66 !important;
    }
    
    /* Deixar textos de input combinando com o terminal */
    input {
        background-color: #0d1117 !important;
        color: #00ff66 !important;
        border: 1px solid #30363d !important;
    }
    </style>
""", unsafe_allowed_html=True)

# --- CONFIGURAÇÃO DE ACESSO ---
USUARIO_CORRETO = "admin"
SENHA_CORRETA = "festraits7"

# --- FUNÇÃO PARA SALVAR OS RESULTADOS NO GOOGLE SHEETS ---
def salvar_resultado_google_sheets(nome_aluno, acertos, total_perguntas):
    porcentagem = (acertos / total_perguntas) * 100
    data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        df_existente = conn.read(ttl=0)
        
        novo_dado = {
            "Data/Hora": [data_hora],
            "Aluno": [nome_aluno],
            "Acertos": [f"{acertos}/{total_perguntas}"],
            "Aproveitamento": [f"{porcentagem:.1f}%"]
        }
        import pandas as pd
        df_novo = pd.DataFrame(novo_dado)
        df_final = pd.concat([df_existente, df_novo], ignore_index=True)
        conn.update(data=df_final)
    except Exception as e:
        with open("resultados_backup.txt", "a", encoding="utf-8") as f:
            f.write(f"[{data_hora}] {nome_aluno} | {acertos}/{total_perguntas}\n")

# --- FUNÇÃO PARA LER O ARQUIVO DO WORD ---
def carregar_perguntas_word(nome_arquivo="perguntas.docx"):
    if not os.path.exists(nome_arquivo):
        return []
    
    doc = Document(nome_arquivo)
    lista_perguntas = []
    bloco = []
    
    for paragrafo in doc.paragraphs:
        texto = paragrafo.text.strip()
        if texto:
            bloco.append(texto)
        else:
            if len(bloco) >= 6:
                lista_perguntas.append({
                    "pergunta": bloco[0],
                    "opcoes": bloco[1:5],
                    "correta": bloco[5]
                })
            bloco = []
            
    if len(bloco) >= 6:
        lista_perguntas.append({
            "pergunta": bloco[0],
            "opcoes": bloco[1:5],
            "correta": bloco[5]
        })
        
    return lista_perguntas

perguntas = carregar_perguntas_word()

if "fase" not in st.session_state:
    st.session_state.fase = "login"
    st.session_state.indice_pergunta = 0
    st.session_state.pontuacao = 0
    st.session_state.nome_aluno = ""
    st.session_state.tempo_limite = 0
    st.session_state.salvou_nota = False

# Cabeçalho estilo Hacker/Programação
st.write("<h1>⚡ QUIZ_INTERATIVO.EXE</h1>", unsafe_allowed_html=True)
st.markdown("<p style='text-align: center; color: #8b949e; font-family: monospace;'>[Status: Online] Core System v2.0</p>", unsafe_allowed_html=True)
st.markdown("<hr style='border-color: #30363d;'>", unsafe_allowed_html=True)

if not perguntas:
    st.error("⚠️ CRITICAL_ERROR: 'perguntas.docx' not found.")
else:
    # --- TELA 1: LOGIN ---
    if st.session_state.fase == "login":
        st.markdown("<h3 style='color: #00ff66;'>🔒 AUTHENTICATION_REQUIRED</h3>", unsafe_allowed_html=True)
        st.write("Insira a chave de criptografia para liberar o terminal de teste.")
        
        with st.container():
            usuario = st.text_input("USER_ID:", key="user_input")
            senha = st.text_input("ACCESS_PASSWORD:", type="password", key="pass_input")
            
            if st.button("EXECUTE_LOGIN", use_container_width=True):
                if usuario == USUARIO_CORRETO and senha == SENHA_CORRETA:
                    st.session_state.fase = "aluno"
                    st.rerun()
                else:
                    st.error("ACCESS_DENIED: Credenciais incorretas.")

    # --- TELA 2: IDENTIFICAÇÃO DO ALUNO ---
    elif st.session_state.fase == "aluno":
        st.markdown("<h3 style='color: #00ff66;'>📝 STUDENT_REGISTER</h3>", unsafe_allowed_html=True)
        
        with st.container():
            nome = st.text_input("NOME COMPLETO DO ESTUDANTE:")
            
            if st.button("INITIALIZE_QUIZ_MODULE", use_container_width=True):
                if nome.strip() != "":
                    st.session_state.nome_aluno = nome
                    st.session_state.indice_pergunta = 0
                    st.session_state.pontuacao = 0
                    st.session_state.salvou_nota = False
                    st.session_state.fase = "proxima_com_tempo"
                    st.rerun()
                else:
                    st.warning("Aviso: Identificação obrigatória para gerar relatórios.")

    # --- PONTE PARA RESETAR O TEMPO ---
    elif st.session_state.fase == "proxima_com_tempo":
        st.session_state.tempo_limite = time.time() + 60
        st.session_state.fase = "jogando"
        st.rerun()

    # --- TELA 3: O QUIZ ---
    elif st.session_state.fase == "jogando":
        idx = st.session_state.indice_pergunta
        dados = perguntas[idx]
        
        tempo_restante = int(st.session_state.tempo_limite - time.time())
        
        if tempo_restante <= 0:
            st.warning("⏱️ TIME_OUT! Próxima questão carregando...")
            time.sleep(1)
            if idx + 1 < len(perguntas):
                st.session_state.indice_pergunta += 1
                st.session_state.fase = "proxima_com_tempo"
            else:
                st.session_state.fase = "fim"
            st.rerun()
            
        col_aluno, col_temp = st.columns([3, 1])
        with col_aluno:
            st.markdown(f"👤 **USER:** {st.session_state.nome_aluno}")
            st.markdown(f"📋 **TASK:** {idx + 1} de {len(perguntas)}")
        with col_temp:
            cor_tempo = "#EF4444" if tempo_restante <= 15 else "#00ff66"
            st.markdown(f"<h3 style='text-align: right; color: {cor_tempo}; margin:0;'>⏱️ {tempo_restante}s</h3>", unsafe_allowed_html=True)
            
        st.markdown("<br>", unsafe_allowed_html=True)
        st.info(f"// QUESTÃO {idx + 1}:\n{dados['pergunta']}")
        
        with st.form(key=f"form_perg_{idx}"):
            resposta = st.radio("Selecione a resposta:", dados["opcoes"], label_visibility="collapsed")
            enviar = st.form_submit_button(label="SUBMIT_RESPONSE >>", use_container_width=True)
            
            if enviar:
                if resposta == dados["correta"]:
                    st.session_state.pontuacao += 1
                    
                if idx + 1 < len(perguntas):
                    st.session_state.indice_pergunta += 1
                    st.session_state.fase = "proxima_com_tempo"
                else:
                    st.session_state.fase = "fim"
                st.rerun()
                
        time.sleep(1)
        st.rerun()

    # --- TELA 4: FIM E PONTUAÇÃO ---
    elif st.session_state.fase == "fim":
        total = len(perguntas)
        acertos = st.session_state.pontuacao
        
        if not st.session_state.salvou_nota:
            salvar_resultado_google_sheets(st.session_state.nome_aluno, acertos, total)
            st.session_state.salvou_nota = True
            
        st.balloons()
        
        with st.container():
            st.markdown("<h3 style='color: #00ff66; text-align: center;'>🎉 EVALUATION_SUCCESSFUL</h3>", unsafe_allowed_html=True)
            st.markdown(f"<p style='text-align: center;'>Dados criptografados e enviados para a central do professor.</p>", unsafe_allowed_html=True)
            
            porcentagem = (acertos / total) * 100
            
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                st.metric(label="ACERTOS", value=f"{acertos} / {total}")
            with col_m2:
                st.metric(label="A PROVEITAMENTO", value=f"{porcentagem:.1f}%")
                
            st.write("Conexão encerrada com segurança. Você já pode fechar a janela.")
        
        st.markdown("<br><br>", unsafe_allowed_html=True)
        if st.button("RESET_TERMINAL_FOR_NEXT_USER", use_container_width=True):
            for chave in list(st.session_state.keys()):
                del st.session_state[chave]
            st.rerun()