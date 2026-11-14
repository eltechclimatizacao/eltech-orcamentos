import os
import re
import json
import sqlite3
from urllib.parse import quote
from datetime import datetime
import streamlit as st
from fpdf import FPDF
import base64

# 1. CONFIGURAÇÃO DA PÁGINA (Sempre o primeiro comando Streamlit)
st.set_page_config(
    page_title="ELTECH Orçamentos",
    page_icon="❄️",
    layout="wide"
)

DB_PATH = "orcamentos.db"


def aplicar_estilo_visual():
    """Aplica CSS moderno e identidade visual corporativa ELTECH."""
    st.markdown("""
        <style>
        /* Ajusta largura máxima para ficar elegante em telas grandes */
        .block-container {
            max-width: 1120px;
            padding-top: 1.8rem;
            padding-bottom: 3rem;
        }

        /* Banner Superior Corporativo */
        .eltech-header {
            background: linear-gradient(135deg, #0B1F52 0%, #102C74 60%, #1D4ED8 100%);
            padding: 1.6rem 2rem;
            border-radius: 14px;
            color: #FFFFFF;
            margin-bottom: 1.5rem;
            box-shadow: 0 8px 20px rgba(16, 44, 116, 0.18);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        .eltech-header h1 {
            color: #FFFFFF !important;
            font-size: 1.75rem;
            font-weight: 700;
            margin: 0;
            letter-spacing: -0.02em;
        }
        .eltech-header p {
            color: #DBEAFE;
            font-size: 0.95rem;
            margin: 0.25rem 0 0 0;
        }
        .eltech-badge {
            background: rgba(255, 255, 255, 0.15);
            border: 1px solid rgba(255, 255, 255, 0.25);
            padding: 0.4rem 0.9rem;
            border-radius: 999px;
            font-size: 0.8rem;
            font-weight: 600;
            color: #FFFFFF;
        }

        /* Cartões de Métricas (Dashboard) */
        div[data-testid="stMetric"] {
            background-color: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-left: 5px solid #102C74;
            padding: 1rem 1.2rem;
            border-radius: 10px;
            box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
        }

        /* Containers com borda mais suaves */
        div[data-testid="stVerticalBlockBorderWrapper"] {
            background-color: #FFFFFF;
            border-radius: 12px !important;
            border: 1px solid #E2E8F0 !important;
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.03);
        }

        /* Abas de Navegação */
        button[data-baseweb="tab"] {
            font-weight: 600;
            font-size: 0.95rem;
            padding: 0.6rem 1.2rem;
        }

        /* Botões mais modernos */
        .stButton > button, .stDownloadButton > button, .stLinkButton > a {
            border-radius: 8px !important;
            font-weight: 600 !important;
            transition: all 0.2s ease;
        }
        /* Cartão da Tela de Login */
        .login-banner {
            background: linear-gradient(135deg, #0B1F52 0%, #102C74 60%, #1D4ED8 100%);
            padding: 2rem 1.5rem;
            border-radius: 14px;
            text-align: center;
            color: #FFFFFF;
            margin-bottom: 1.2rem;
            box-shadow: 0 10px 25px rgba(16, 44, 116, 0.2);
        }
        .login-banner h2 {
            color: #FFFFFF !important;
            margin: 0.5rem 0 0.2rem 0;
            font-size: 1.5rem;
            font-weight: 700;
        }
        .login-banner p {
            color: #DBEAFE;
            margin: 0;
            font-size: 0.9rem;
        }

        </style>
    """, unsafe_allow_html=True)


def exibir_cabecalho_corporativo():
    """Renderiza o banner superior da ELTECH com logo (se existir) ou título estilizado."""
    logo_html = ""
    if os.path.exists("logo.png"):
        with open("logo.png", "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        logo_html = f'<img src="data:image/png;base64,{b64}" style="height:48px; background:white; padding:6px 12px; border-radius:8px; margin-right:16px;">'

    st.markdown(f"""
        <div class="eltech-header">
            <div style="display:flex; align-items:center;">
                {logo_html}
                <div>
                    <h1>ELTECH CLIMATIZAÇÃO</h1>
                    <p>Sistema Inteligente de Orçamentos e Gestão Comercial</p>
                </div>
            </div>
            <div class="eltech-badge">❄️ Painel Administrativo</div>
        </div>
    """, unsafe_allow_html=True)


# Ativa o estilo visual logo após as definições acima
aplicar_estilo_visual()


# 2. FUNÇÕES AUXILIARES
def formatar_moeda(valor: float) -> str:
    """Formata float para o padrão monetário brasileiro (R$ 1.234,56)."""
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def limpar_texto_pdf(texto: str) -> str:
    """Garante compatibilidade com FPDF clássico (latin-1)."""
    if not texto:
        return ""
    return texto.encode("latin-1", "replace").decode("latin-1")


def gerar_link_whatsapp(telefone: str, cliente: str, num_orcamento: int, total: float) -> str:
    """Gera link wa.me formatado para o Brasil (+55) com mensagem pronta."""
    digitos = re.sub(r"\D", "", telefone or "")
    if not digitos:
        return ""
    if len(digitos) in (10, 11):
        digitos = "55" + digitos
    mensagem = (
        f"Olá {cliente}, tudo bem? Segue o *Orçamento #{num_orcamento:04d}* "
        f"da *ELTECH Climatização* no valor total de *{formatar_moeda(total)}*.\n\n"
        f"Estou enviando o documento PDF detalhado logo abaixo. Qualquer dúvida, estou à disposição!"
    )
    return f"https://wa.me/{digitos}?text={quote(mensagem)}"


def gerar_csv_orcamentos(orcamentos: list) -> bytes:
    """Gera arquivo CSV compatível com Excel Brasil (separado por ';') sem dependências extras."""
    linhas = [
        "Numero;Data Emissao;Cliente;CPF/CNPJ;Contato;Endereco;Subtotal;Desconto;Total Final;Pagamento;Status;Itens"
    ]
    for orc in orcamentos:
        itens_lista = json.loads(orc["itens_json"])
        resumo_itens = " | ".join([f"{it['produto']} (x{it['qtd']})" for it in itens_lista])
        linha = ";".join([
            f"{orc['id']:04d}",
            orc["data_emissao"],
            (orc["cliente"] or "").replace(";", ","),
            (orc["documento"] or "").replace(";", ","),
            (orc["contato"] or "").replace(";", ","),
            (orc["endereco"] or "").replace(";", ","),
            f"{orc['subtotal']:.2f}".replace(".", ","),
            f"{orc['desconto']:.2f}".replace(".", ","),
            f"{orc['total']:.2f}".replace(".", ","),
            (orc["condicao_pagamento"] or "").replace(";", ","),
            orc["status"],
            resumo_itens.replace(";", ",")
        ])
        linhas.append(linha)
    conteudo = "\n".join(linhas)
    return conteudo.encode("utf-8-sig")



# 3. BANCO DE DADOS (PRODUTOS E HISTÓRICO DE ORÇAMENTOS)
def inicializar_banco():
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()

        # Tabela de Produtos/Serviços
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS produtos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT UNIQUE NOT NULL,
                preco REAL NOT NULL
            )
        """)

        # Tabela de Orçamentos Salvos (Histórico + Numeração)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS orcamentos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                data_emissao TEXT NOT NULL,
                cliente TEXT NOT NULL,
                contato TEXT,
                endereco TEXT,
                documento TEXT,
                itens_json TEXT NOT NULL,
                subtotal REAL NOT NULL,
                desconto REAL NOT NULL,
                total REAL NOT NULL,
                condicao_pagamento TEXT,
                observacoes TEXT,
                status TEXT DEFAULT 'Pendente'
            )
        """)

        cursor.execute("SELECT COUNT(*) FROM produtos")
        if cursor.fetchone() == 0:
            produtos_iniciais = [
                ("Instalação Ar-Condicionado 9000 BTUs", 350.0),
                ("Higienização e Limpeza Completa", 150.0),
                ("Carga de Gás R410A", 220.0),
                ("Manutenção Corretiva Sistema de Climatização", 250.0)
            ]
            cursor.executemany(
                "INSERT INTO produtos (nome, preco) VALUES (?, ?)",
                produtos_iniciais
            )
        conn.commit()


def listar_produtos():
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, nome, preco FROM produtos ORDER BY nome ASC")
        return cursor.fetchall()


def adicionar_produto_db(nome: str, preco: float) -> bool:
    try:
        with sqlite3.connect(DB_PATH) as conn:
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO produtos (nome, preco) VALUES (?, ?)",
                (nome.strip(), preco)
            )
            conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False


def atualizar_preco_db(produto_id: int, novo_preco: float):
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE produtos SET preco = ? WHERE id = ?",
            (novo_preco, produto_id)
        )
        conn.commit()


def excluir_produto_db(produto_id: int):
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM produtos WHERE id = ?", (produto_id,))
        conn.commit()


def salvar_orcamento_db(dados: dict) -> int:
    """Salva o orçamento no banco e retorna o número sequencial (ID) gerado."""
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO orcamentos (
                data_emissao, cliente, contato, endereco, documento,
                itens_json, subtotal, desconto, total,
                condicao_pagamento, observacoes, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pendente')
        """, (
            dados["data_emissao"],
            dados["cliente"],
            dados["contato"],
            dados["endereco"],
            dados["documento"],
            json.dumps(dados["itens"], ensure_ascii=False),
            dados["subtotal"],
            dados["desconto"],
            dados["total"],
            dados["condicao_pagamento"],
            dados["observacoes"]
        ))
        conn.commit()
        return cursor.lastrowid


def listar_orcamentos_db(filtro_busca: str = "", filtro_status: str = "Todos"):
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        query = "SELECT * FROM orcamentos WHERE 1=1"
        params = []

        if filtro_busca.strip():
            termo = f"%{filtro_busca.strip()}%"
            query += " AND (cliente LIKE ? OR CAST(id AS TEXT) LIKE ?)"
            params.extend([termo, termo])

        if filtro_status != "Todos":
            query += " AND status = ?"
            params.append(filtro_status)

        query += " ORDER BY id DESC"
        cursor.execute(query, params)
        return [dict(row) for row in cursor.fetchall()]


def atualizar_status_orcamento_db(orcamento_id: int, novo_status: str):
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE orcamentos SET status = ? WHERE id = ?",
            (novo_status, orcamento_id)
        )
        conn.commit()


# 4. CLASSE PARA GERAR O PDF DA ELTECH
class PDF_Eltech(FPDF):
    def header(self):
        if os.path.exists("logo.png"):
            self.image("logo.png", 10, 8, 65)
        else:
            self.set_font("Arial", "B", 16)
            self.set_text_color(16, 44, 116)
            self.cell(0, 10, limpar_texto_pdf("ELTECH CLIMATIZAÇÃO"), ln=False)

        self.set_font("Arial", "", 9)
        self.set_text_color(50, 50, 50)
        self.set_x(-100)
        cabecalho_dir = (
            "ELTECH CLIMATIZAÇÃO\n"
            "E-mail: contato@eltechclimatizacao.com.br\n"
            "Telefone/WhatsApp: (51) 99232-9339\n"
            "Soluções em Conforto Térmico"
        )
        self.multi_cell(90, 4, limpar_texto_pdf(cabecalho_dir), align="R")

        self.ln(12)
        self.set_draw_color(16, 44, 116)
        self.set_line_width(0.5)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.set_text_color(120, 120, 120)
        self.cell(0, 10, limpar_texto_pdf(f"ELTECH CLIMATIZAÇÃO - Página {self.page_no()}"), align="C")


def gerar_pdf_documento(
    num_orcamento,
    data_emissao,
    cliente,
    contato,
    endereco,
    documento,
    itens,
    subtotal,
    desconto,
    total,
    condicao_pagamento,
    obs_adicional
):
    pdf = PDF_Eltech()
    pdf.add_page()

    pdf.ln(2)
    pdf.set_font("Arial", "B", 14)
    pdf.set_text_color(16, 44, 116)
    titulo_doc = f"ORÇAMENTO DE SERVIÇOS / PRODUTOS  -  Nº {num_orcamento:04d}"
    pdf.cell(0, 8, limpar_texto_pdf(titulo_doc), ln=True, align="L")

    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(0, 5, limpar_texto_pdf(f"Data de Emissão: {data_emissao}"), ln=True, align="L")
    pdf.ln(4)

    # Bloco Dados do Cliente
    pdf.set_font("Arial", "B", 11)
    pdf.set_text_color(255, 255, 255)
    pdf.set_fill_color(16, 44, 116)
    pdf.cell(0, 7, "  DADOS DO CLIENTE", ln=True, fill=True)

    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", "", 10)
    pdf.ln(2)
    pdf.cell(0, 6, limpar_texto_pdf(f"Nome / Razão Social: {cliente}"), ln=True)
    if documento:
        pdf.cell(0, 6, limpar_texto_pdf(f"CPF / CNPJ: {documento}"), ln=True)
    if contato:
        pdf.cell(0, 6, limpar_texto_pdf(f"Celular / WhatsApp: {contato}"), ln=True)
    if endereco:
        pdf.cell(0, 6, limpar_texto_pdf(f"Endereço: {endereco}"), ln=True)
    pdf.ln(5)

    # Cabeçalho da Tabela
    pdf.set_font("Arial", "B", 10)
    pdf.set_text_color(255, 255, 255)
    pdf.set_fill_color(40, 40, 40)

    pdf.cell(95, 8, limpar_texto_pdf(" Descrição Técnica do Serviço / Equipamento"), border=1, fill=True)
    pdf.cell(20, 8, "Qtd", border=1, align="C", fill=True)
    pdf.cell(35, 8, limpar_texto_pdf("Preço Unitário"), border=1, align="R", fill=True)
    pdf.cell(40, 8, "Total Item", border=1, align="R", fill=True, ln=True)

    # Listagem dos itens com quebra de linha automática
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", "", 9.5)
    for item in itens:
        desc = limpar_texto_pdf(f" {item['produto']}")
        qtd_str = str(item["qtd"])
        unit_str = formatar_moeda(item["preco_unit"])
        tot_str = formatar_moeda(item["total_item"])

        linhas = pdf.multi_cell(95, 6, desc, split_only=True)
        altura_linha = max(8, len(linhas) * 6)

        if pdf.get_y() + altura_linha > 255:
            pdf.add_page()

        x_inicio = pdf.get_x()
        y_inicio = pdf.get_y()

        pdf.multi_cell(95, altura_linha / len(linhas), desc, border=1, align="L")
        pdf.set_xy(x_inicio + 95, y_inicio)
        pdf.cell(20, altura_linha, qtd_str, border=1, align="C")
        pdf.cell(35, altura_linha, unit_str, border=1, align="R")
        pdf.cell(40, altura_linha, tot_str, border=1, align="R", ln=True)

    if obs_adicional:
        pdf.ln(4)
        pdf.set_font("Arial", "B", 10)
        pdf.cell(0, 5, limpar_texto_pdf("Observações Técnicas / Escopo do Serviço:"), ln=True)
        pdf.set_font("Arial", "", 10)
        pdf.multi_cell(0, 5, limpar_texto_pdf(obs_adicional))

    pdf.ln(6)

    # Quadro de Totais (Exibe Subtotal e Desconto se houver)
    if desconto > 0:
        pdf.set_font("Arial", "", 10)
        pdf.cell(130, 7, " SUBTOTAL:", border=1)
        pdf.cell(60, 7, f"{formatar_moeda(subtotal)} ", border=1, align="R", ln=True)

        pdf.set_text_color(180, 0, 0)
        pdf.cell(130, 7, " DESCONTO CONCEDIDO:", border=1)
        pdf.cell(60, 7, f"- {formatar_moeda(desconto)} ", border=1, align="R", ln=True)
        pdf.set_text_color(0, 0, 0)

    pdf.set_font("Arial", "B", 12)
    pdf.set_fill_color(230, 240, 255)
    pdf.cell(130, 10, " VALOR TOTAL DO INVESTIMENTO:", border=1, fill=True)
    pdf.cell(60, 10, f"{formatar_moeda(total)} ", border=1, align="R", fill=True, ln=True)

    # Condições comerciais
    pdf.ln(8)
    pdf.set_font("Arial", "B", 9)
    pdf.set_text_color(60, 60, 60)
    pdf.cell(0, 5, limpar_texto_pdf(f"* Condições de pagamento: {condicao_pagamento}"), ln=True)

    pdf.set_font("Arial", "I", 9)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 5, limpar_texto_pdf(f"* Validade deste orçamento: 10 dias a contar de {data_emissao}."), ln=True)
    pdf.cell(0, 5, limpar_texto_pdf("* Garantia de instalação de equipamentos novos e usados: 3 meses."), ln=True)

    # Bloco de Assinaturas
    if pdf.get_y() > 240:
        pdf.add_page()
    pdf.ln(18)
    pdf.set_draw_color(120, 120, 120)
    y_assinatura = pdf.get_y()
    pdf.line(15, y_assinatura, 95, y_assinatura)
    pdf.line(115, y_assinatura, 195, y_assinatura)

    pdf.set_font("Arial", "", 8.5)
    pdf.set_text_color(60, 60, 60)
    pdf.set_xy(15, y_assinatura + 2)
    pdf.cell(80, 5, limpar_texto_pdf("ELTECH CLIMATIZAÇÃO"), align="C")
    pdf.set_xy(115, y_assinatura + 2)
    pdf.cell(80, 5, limpar_texto_pdf(f"De acordo: {cliente[:35]}"), align="C")

    return bytes(pdf.output())


# 5. INICIALIZAÇÃO DE ESTADO E AUTENTICAÇÃO
inicializar_banco()

if "itens_orcamento" not in st.session_state:
    st.session_state.itens_orcamento = []

if "ultimo_orcamento_gerado" not in st.session_state:
    st.session_state.ultimo_orcamento_gerado = None

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if "usuario_logado" not in st.session_state:
    st.session_state.usuario_logado = ""

# Usuários autorizados: "login" (sempre em minúsculo): ("senha", "Nome de Exibição")
USUARIOS_SISTEMA = {
    "george": ("21122018", "George - Administrador"),
    "eltech": ("21122018", "Equipe Comercial"),
}

if not st.session_state.autenticado:
    _, col_login, _ = st.columns([1, 1.3, 1])

    with col_login:
        st.write("")
        logo_login_html = ""
        if os.path.exists("logo.png"):
            with open("logo.png", "rb") as f:
                b64_logo = base64.b64encode(f.read()).decode()
            logo_login_html = (
                f'<img src="data:image/png;base64,{b64_logo}" '
                f'style="height:54px; background:white; padding:8px 14px; border-radius:10px; margin-bottom:8px;">'
            )
        else:
            logo_login_html = '<div style="font-size:2.2rem;">❄️</div>'

        st.markdown(f"""
            <div class="login-banner">
                {logo_login_html}
                <h2>ELTECH CLIMATIZAÇÃO</h2>
                <p>Portal Administrativo de Orçamentos</p>
            </div>
        """, unsafe_allow_html=True)

        with st.container(border=True):
            st.markdown("#### 🔐 Acesso ao Sistema")
            with st.form("form_login"):
                usuario_digitado = st.text_input(
                    "Usuário",
                    placeholder="login"
                ).strip().lower()
                senha_digitada = st.text_input(
                    "Senha",
                    type="password",
                    placeholder="senha"
                )
                st.write("")
                entrar = st.form_submit_button("Entrar no Painel ➡️", type="primary", use_container_width=True)

                if entrar:
                    dados_usuario = USUARIOS_SISTEMA.get(usuario_digitado)
                    if dados_usuario:
                        senha_correta, nome_exibicao = dados_usuario
                        if senha_digitada == senha_correta:
                            st.session_state.autenticado = True
                            st.session_state.usuario_logado = nome_exibicao
                            st.rerun()
                        else:
                            st.error("❌ Senha incorreta!")
                    else:
                        st.error("❌ Usuário não encontrado!")

        st.caption("© ELTECH Climatização — Acesso restrito a colaboradores autorizados.")
    st.stop()


# 6. INTERFACE PRINCIPAL
exibir_cabecalho_corporativo()

col_bem_vindo, col_sair = st.columns([5.5, 1.2])
with col_bem_vindo:
    st.caption(f"👤 Conectado como: **{st.session_state.usuario_logado}**")
with col_sair:
    if st.button("🔒 Sair", use_container_width=True):
        st.session_state.autenticado = False
        st.session_state.usuario_logado = ""
        st.rerun()

aba1, aba2, aba3 = st.tabs([
    "📄 Novo Orçamento",
    "🗂️ Histórico & Dashboard",
    "⚙️ Produtos e Preços"
])

# ==============================================================================
# ABA 1: CONSTRUIR NOVO ORÇAMENTO
# ==============================================================================
with aba1:


    st.header("Construir Novo Orçamento")
    produtos_db = listar_produtos()
    dicionario_produtos = {nome: preco for _, nome, preco in produtos_db}

    # Campos com suporte a preenchimento automático ao duplicar orçamento do histórico
    col1, col2 = st.columns(2)
    with col1:
        cliente = st.text_input("Nome do Cliente / Empresa *", key="input_cliente")
    with col2:
        contato = st.text_input("Celular / WhatsApp (Ex: 51999998888)", key="input_contato")

    col3, col4 = st.columns(2)
    with col3:
        endereco = st.text_input("Endereço", key="input_endereco")
    with col4:
        documento = st.text_input("CPF / CNPJ", key="input_documento")

    st.markdown("### Adicionar Itens")
    tipo_item = st.radio(
        "Origem do item:",
        ["Selecionar do Catálogo", "Item Avulso (Digitar na hora)"],
        horizontal=True
    )

    if tipo_item == "Selecionar do Catálogo":
        if produtos_db:
            col_prod, col_qtd, col_preco = st.columns([3, 1, 1.5])
            with col_prod:
                produto_selecionado = st.selectbox(
                    "Serviço / Produto",
                    list(dicionario_produtos.keys()),
                    key="produto_atual"
                )
            with col_qtd:
                quantidade = st.number_input("Qtd", min_value=1, value=1, step=1, key="qtd_cat")
            with col_preco:
                preco_unitario = st.number_input(
                    "Valor Unit. (R$)",
                    min_value=0.0,
                    value=float(dicionario_produtos[produto_selecionado]),
                    step=10.0,
                    key="preco_cat"
                )

            if st.button("➕ Adicionar Item ao Orçamento", use_container_width=True):
                st.session_state.itens_orcamento.append({
                    "produto": produto_selecionado,
                    "qtd": quantidade,
                    "preco_unit": preco_unitario,
                    "total_item": preco_unitario * quantidade
                })
                st.session_state.ultimo_orcamento_gerado = None
                st.toast(f"'{produto_selecionado}' adicionado!", icon="✅")
                st.rerun()
        else:
            st.warning("Nenhum produto cadastrado no catálogo. Use 'Item Avulso' ou cadastre na aba ao lado.")
    else:
        col_av_desc, col_av_qtd, col_av_preco = st.columns([3, 1, 1.5])
        with col_av_desc:
            avulso_nome = st.text_input("Descrição do serviço/peça avulsa", placeholder="Ex: Suporte reforçado 50cm")
        with col_av_qtd:
            avulso_qtd = st.number_input("Qtd", min_value=1, value=1, step=1, key="qtd_av")
        with col_av_preco:
            avulso_preco = st.number_input("Valor Unit. (R$)", min_value=0.0, value=0.0, step=10.0, key="preco_av")

        if st.button("➕ Adicionar Item Avulso", use_container_width=True):
            if avulso_nome.strip():
                st.session_state.itens_orcamento.append({
                    "produto": avulso_nome.strip(),
                    "qtd": avulso_qtd,
                    "preco_unit": avulso_preco,
                    "total_item": avulso_preco * avulso_qtd
                })
                st.session_state.ultimo_orcamento_gerado = None
                st.toast(f"'{avulso_nome.strip()}' adicionado!", icon="✅")
                st.rerun()
            else:
                st.warning("Digite a descrição do item avulso.")

    # Carrinho e Fechamento
    if st.session_state.itens_orcamento:
        st.markdown("---")
        st.markdown("### Itens no Orçamento")
        subtotal_calculado = 0.0

        for idx, item in enumerate(st.session_state.itens_orcamento):
            c_desc, c_val, c_del = st.columns([4, 2, 0.7])
            with c_desc:
                st.write(f"🔹 **{item['produto']}** (x{item['qtd']} de {formatar_moeda(item['preco_unit'])})")
            with c_val:
                st.write(formatar_moeda(item["total_item"]))
            with c_del:
                if st.button("🗑️", key=f"del_item_{idx}", help="Remover este item"):
                    st.session_state.itens_orcamento.pop(idx)
                    st.session_state.ultimo_orcamento_gerado = None
                    st.rerun()
            subtotal_calculado += item["total_item"]

        if st.button("❌ Limpar Carrinho"):
            st.session_state.itens_orcamento = []
            st.session_state.ultimo_orcamento_gerado = None
            st.rerun()

        st.markdown("### Condições Comerciais e Desconto")
        col_desc_tipo, col_desc_val = st.columns([1.5, 2])
        with col_desc_tipo:
            tipo_desconto = st.selectbox("Tipo de Desconto", ["Sem desconto", "Valor Fixo (R$)", "Porcentagem (%)"])
        with col_desc_val:
            if tipo_desconto == "Valor Fixo (R$)":
                valor_desconto_input = st.number_input(
                    "Valor do Desconto (R$)",
                    min_value=0.0,
                    max_value=float(subtotal_calculado),
                    value=0.0,
                    step=10.0
                )
                desconto_calculado = valor_desconto_input
            elif tipo_desconto == "Porcentagem (%)":
                perc_desconto = st.number_input(
                    "Porcentagem de Desconto (%)",
                    min_value=0.0,
                    max_value=100.0,
                    value=0.0,
                    step=5.0
                )
                desconto_calculado = subtotal_calculado * (perc_desconto / 100.0)
            else:
                desconto_calculado = 0.0

        total_final = max(0.0, subtotal_calculado - desconto_calculado)

        opcoes_pagamento = [
            "À vista no PIX ou Dinheiro",
            "50% de entrada + 50% na conclusão do serviço",
            "Cartão de Crédito em até 3x sem juros",
            "A combinar conforme execução",
            "Outra (Digitar personalizada)"
        ]
        condicao_escolhida = st.selectbox("Forma de Pagamento", opcoes_pagamento)
        if condicao_escolhida == "Outra (Digitar personalizada)":
            condicao_pagamento_final = st.text_input("Descreva a condição de pagamento", value="A combinar")
        else:
            condicao_pagamento_final = condicao_escolhida

        obs_tecnica = st.text_area(
            "Observações Técnicas / Escopo (Opcional)",
            placeholder="Ex: Inclui infraestrutura de cobre de até 3 metros e suporte externo."
        )

        st.markdown("---")
        if desconto_calculado > 0:
            st.write(f"**Subtotal:** {formatar_moeda(subtotal_calculado)} | **Desconto:** -{formatar_moeda(desconto_calculado)}")
        st.subheader(f"💰 Valor Total Final: {formatar_moeda(total_final)}")

        if cliente.strip():
            if st.button("✅ Salvar Orçamento e Gerar PDF", type="primary", use_container_width=True):
                data_hoje = datetime.now().strftime("%d/%m/%Y")
                num_gerado = salvar_orcamento_db({
                    "data_emissao": data_hoje,
                    "cliente": cliente.strip(),
                    "contato": contato.strip(),
                    "endereco": endereco.strip(),
                    "documento": documento.strip(),
                    "itens": st.session_state.itens_orcamento,
                    "subtotal": subtotal_calculado,
                    "desconto": desconto_calculado,
                    "total": total_final,
                    "condicao_pagamento": condicao_pagamento_final.strip(),
                    "observacoes": obs_tecnica.strip()
                })

                pdf_bytes = gerar_pdf_documento(
                    num_orcamento=num_gerado,
                    data_emissao=data_hoje,
                    cliente=cliente.strip(),
                    contato=contato.strip(),
                    endereco=endereco.strip(),
                    documento=documento.strip(),
                    itens=st.session_state.itens_orcamento,
                    subtotal=subtotal_calculado,
                    desconto=desconto_calculado,
                    total=total_final,
                    condicao_pagamento=condicao_pagamento_final.strip(),
                    obs_adicional=obs_tecnica.strip()
                )

                st.session_state.ultimo_orcamento_gerado = {
                    "numero": num_gerado,
                    "cliente": cliente.strip(),
                    "contato": contato.strip(),
                    "total": total_final,
                    "pdf": pdf_bytes
                }
                st.success(f"Orçamento **#{num_gerado:04d}** salvo no histórico com sucesso!")

                        # Exibe os botões de Download e WhatsApp após gerar
            if st.session_state.ultimo_orcamento_gerado:
                dados_gerados = st.session_state.ultimo_orcamento_gerado
                nome_arq = f"Orcamento_{dados_gerados['numero']:04d}_ELTECH_{dados_gerados['cliente'].replace(' ', '_')}.pdf"

                col_dl, col_wa = st.columns(2)
                with col_dl:
                    st.download_button(
                        label=f"🖨️ Baixar PDF (#{dados_gerados['numero']:04d})",
                        data=dados_gerados["pdf"],
                        file_name=nome_arq,
                        mime="application/pdf",
                        use_container_width=True
                    )
                with col_wa:
                    link_wa = gerar_link_whatsapp(
                        dados_gerados["contato"],
                        dados_gerados["cliente"],
                        dados_gerados["numero"],
                        dados_gerados["total"]
                                        )
                    if link_wa:
                        st.link_button(
                            label="📲 Enviar Mensagem no WhatsApp",
                            url=link_wa,
                            use_container_width=True
                        )
                    else:
                        st.info("Informe o celular do cliente para habilitar o link do WhatsApp.")
        else:
            st.warning("⚠️ Preencha o **Nome do Cliente** no topo para salvar e gerar o PDF.")


# ==============================================================================
# ABA 2: DASHBOARD FINANCEIRO E HISTÓRICO DE ORÇAMENTOS
# ==============================================================================
with aba2:
    st.header("Dashboard Financeiro & Histórico")

    # Filtros de busca e status
    col_busca, col_status_filtro = st.columns([2.5, 1.5])
    with col_busca:
        busca_hist = st.text_input("Buscar por cliente ou número", placeholder="Ex: João ou 12")
    with col_status_filtro:
        filtro_status = st.selectbox("Filtrar por Status", ["Todos", "Pendente", "Aprovado", "Recusado"])

    orcamentos_salvos = listar_orcamentos_db(busca_hist, filtro_status)

    if orcamentos_salvos:
        # Cálculos do Dashboard Financeiro
        total_geral_orcado = sum(o["total"] for o in orcamentos_salvos)
        total_aprovado = sum(o["total"] for o in orcamentos_salvos if o["status"] == "Aprovado")
        total_pendente = sum(o["total"] for o in orcamentos_salvos if o["status"] == "Pendente")

        qtd_total = len(orcamentos_salvos)
        qtd_aprovados = sum(1 for o in orcamentos_salvos if o["status"] == "Aprovado")
        taxa_conversao = (qtd_aprovados / qtd_total * 100) if qtd_total > 0 else 0.0

        # Cartões de Indicadores (KPIs)
        with st.container(border=True):
            st.markdown("#### 📊 Resumo Financeiro (Baseado no filtro atual)")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Orçado", formatar_moeda(total_geral_orcado))
            m2.metric("Faturamento Fechado", formatar_moeda(total_aprovado), delta=f"{qtd_aprovados} aprovado(s)")
            m3.metric("Em Negociação", formatar_moeda(total_pendente))
            m4.metric("Taxa de Aprovação", f"{taxa_conversao:.1f}%")

        # Botão de Exportação para Excel
        col_caption, col_excel = st.columns([2.2, 1.8])
        with col_caption:
            st.caption(f"Exibindo {qtd_total} orçamento(s) encontrado(s).")
        with col_excel:
            csv_bytes = gerar_csv_orcamentos(orcamentos_salvos)
            data_arq = datetime.now().strftime("%Y-%m-%d")
            st.download_button(
                label="📥 Exportar Planilha (Excel)",
                data=csv_bytes,
                file_name=f"Relatorio_Orcamentos_ELTECH_{data_arq}.csv",
                mime="text/csv",
                use_container_width=True
            )

        st.markdown("---")

        # Listagem detalhada dos orçamentos
        for orc in orcamentos_salvos:
            icone_status = {"Pendente": "🟡", "Aprovado": "🟢", "Recusado": "🔴"}.get(orc["status"], "⚪")
            titulo_expander = (
                f"{icone_status} **#{orc['id']:04d}** | {orc['cliente']} "
                f"— {formatar_moeda(orc['total'])} ({orc['data_emissao']})"
            )

            with st.expander(titulo_expander):
                c_info1, c_info2 = st.columns(2)
                with c_info1:
                    st.write(f"**Contato:** {orc['contato'] or 'Não informado'}")
                    st.write(f"**CPF/CNPJ:** {orc['documento'] or 'Não informado'}")
                with c_info2:
                    st.write(f"**Endereço:** {orc['endereco'] or 'Não informado'}")
                    st.write(f"**Pagamento:** {orc['condicao_pagamento'] or 'A combinar'}")

                itens_lista = json.loads(orc["itens_json"])
                st.markdown("**Itens cotados:**")
                for it in itens_lista:
                    st.write(f"- {it['produto']} (x{it['qtd']}) — {formatar_moeda(it['total_item'])}")

                if orc["desconto"] > 0:
                    st.write(f"**Desconto aplicado:** {formatar_moeda(orc['desconto'])}")
                if orc["observacoes"]:
                    st.write(f"**Observações:** {orc['observacoes']}")

                st.markdown("---")
                col_st, col_reimp, col_dup = st.columns([1.5, 1.5, 1.5])

                with col_st:
                    opcoes_st = ["Pendente", "Aprovado", "Recusado"]
                    idx_atual = opcoes_st.index(orc["status"]) if orc["status"] in opcoes_st else 0
                    novo_st = st.selectbox(
                        "Status",
                        opcoes_st,
                        index=idx_atual,
                        key=f"st_orc_{orc['id']}",
                        label_visibility="collapsed"
                    )
                    if novo_st != orc["status"]:
                        atualizar_status_orcamento_db(orc["id"], novo_st)
                        st.toast(f"Orçamento #{orc['id']:04d} atualizado para {novo_st}!")
                        st.rerun()

                with col_reimp:
                    pdf_reimpressao = gerar_pdf_documento(
                        num_orcamento=orc["id"],
                        data_emissao=orc["data_emissao"],
                        cliente=orc["cliente"],
                        contato=orc["contato"] or "",
                        endereco=orc["endereco"] or "",
                        documento=orc["documento"] or "",
                        itens=itens_lista,
                        subtotal=orc["subtotal"],
                        desconto=orc["desconto"],
                        total=orc["total"],
                        condicao_pagamento=orc["condicao_pagamento"] or "A combinar",
                        obs_adicional=orc["observacoes"] or ""
                    )
                    st.download_button(
                        label="🖨️ Baixar PDF",
                        data=pdf_reimpressao,
                        file_name=f"Orcamento_{orc['id']:04d}_ELTECH_{orc['cliente'].replace(' ', '_')}.pdf",
                        mime="application/pdf",
                        key=f"dl_hist_{orc['id']}",
                        use_container_width=True
                    )

                with col_dup:
                    if st.button("📋 Carregar p/ Editar", key=f"dup_hist_{orc['id']}", use_container_width=True):
                        st.session_state.input_cliente = orc["cliente"]
                        st.session_state.input_contato = orc["contato"] or ""
                        st.session_state.input_endereco = orc["endereco"] or ""
                        st.session_state.input_documento = orc["documento"] or ""
                        st.session_state.itens_orcamento = itens_lista
                        st.session_state.ultimo_orcamento_gerado = None
                        st.toast("Dados carregados na aba 'Novo Orçamento'!", icon="📋")
    else:
        st.info("Nenhum orçamento encontrado no histórico.")
# ==============================================================================
# ABA 3: GERENCIAMENTO DE PRODUTOS E PREÇOS
# ==============================================================================
with aba3:
    st.header("Catálogo de Produtos e Serviços")
    sub_acao = st.radio(
        "Selecione a ação:",
        ["Incluir Novo Produto", "Visualizar, Alterar ou Excluir"],
        horizontal=True
    )

    if sub_acao == "Incluir Novo Produto":
        st.subheader("Cadastrar Novo Serviço ou Item")
        with st.form("form_novo_produto", clear_on_submit=True):
            novo_nome = st.text_input("Nome descritivo do produto/serviço")
            novo_preco = st.number_input("Preço de tabela (R$)", min_value=0.0, value=0.0, step=50.0)
            salvar_btn = st.form_submit_button("💾 Salvar Produto no Banco")

            if salvar_btn:
                if novo_nome.strip():
                    if adicionar_produto_db(novo_nome, novo_preco):
                        st.success(f"'{novo_nome}' cadastrado com sucesso!")
                        st.rerun()
                    else:
                        st.error("Erro: Já existe um produto com esse nome!")
                else:
                    st.warning("Preencha o nome do produto.")

    elif sub_acao == "Visualizar, Alterar ou Excluir":
        st.subheader("Tabela de Preços Atual")
        itens_armazenados = listar_produtos()

        if itens_armazenados:
            tabela_exibicao = [
                {"ID": i, "Serviço / Produto": n, "Preço Tabela": formatar_moeda(p)}
                for i, n, p in itens_armazenados
            ]
            st.dataframe(tabela_exibicao, use_container_width=True, hide_index=True)

            st.markdown("---")
            st.subheader("Editar ou Remover Item")
            opcoes_edicao = {f"{n} ({formatar_moeda(p)})": (i, n, p) for i, n, p in itens_armazenados}
            selecao_edicao = st.selectbox("Selecione o item:", list(opcoes_edicao.keys()))
            id_sel, nome_sel, preco_atual_sel = opcoes_edicao[selecao_edicao]

            novo_valor_editado = st.number_input(
                "Novo valor comercial (R$)",
                min_value=0.0,
                value=float(preco_atual_sel),
                step=10.0,
                key=f"edit_val_{id_sel}"
            )

            col_upd, col_rem = st.columns(2)
            with col_upd:
                if st.button("🔄 Gravar Alteração de Preço", use_container_width=True):
                    atualizar_preco_db(id_sel, novo_valor_editado)
                    st.success(f"Preço de '{nome_sel}' atualizado!")
                    st.rerun()
            with col_rem:
                if st.button("🗑️ Excluir Produto do Banco", type="secondary", use_container_width=True):
                    excluir_produto_db(id_sel)
                    st.warning(f"Produto '{nome_sel}' excluído.")
                    st.rerun()
        else:
            st.info("Nenhum produto cadastrado.")
