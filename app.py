import os
import re
import json
import base64
from urllib.parse import quote
from datetime import datetime
import streamlit as st
from fpdf import FPDF
import gspread
from google.oauth2.service_account import Credentials

LOGO_PATH = "logo.png"

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
    page_title="ELTECH Climatização - Gestão & Orçamentos",
    page_icon=LOGO_PATH if os.path.exists(LOGO_PATH) else "❄️",
    layout="wide"
)


# 2. ESTILO VISUAL CORPORATIVO ELTECH
def aplicar_estilo_visual():
    st.markdown("""
        <style>
        .block-container {
            max-width: 1180px;
            padding-top: 1.8rem;
            padding-bottom: 3rem;
        }
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
        div[data-testid="stMetric"] {
            background-color: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-left: 5px solid #102C74;
            padding: 1rem 1.2rem;
            border-radius: 10px;
            box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
        }
        div[data-testid="stVerticalBlockBorderWrapper"] {
            background-color: #FFFFFF;
            border-radius: 12px !important;
            border: 1px solid #E2E8F0 !important;
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.03);
        }
        button[data-baseweb="tab"] {
            font-weight: 600;
            font-size: 0.95rem;
            padding: 0.6rem 1.1rem;
        }
        .stButton > button, .stDownloadButton > button, .stLinkButton > a {
            border-radius: 8px !important;
            font-weight: 600 !important;
        }
        </style>
    """, unsafe_allow_html=True)


def exibir_cabecalho_corporativo():
    logo_html = ""
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        logo_html = f'<img src="data:image/png;base64,{b64}" style="height:48px; background:white; padding:6px 12px; border-radius:8px; margin-right:16px;">'

    st.markdown(f"""
        <div class="eltech-header">
            <div style="display:flex; align-items:center;">
                {logo_html}
                <div>
                    <h1>ELTECH CLIMATIZAÇÃO</h1>
                    <p>Gestão Comercial, Financeira e Controle de Estoque</p>
                </div>
            </div>
            <div class="eltech-badge">☁️ Sincronizado Google Sheets</div>
        </div>
    """, unsafe_allow_html=True)


aplicar_estilo_visual()


# 3. FUNÇÕES AUXILIARES
def formatar_moeda(valor: float) -> str:
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def parse_float(valor) -> float:
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor or "0").replace("R$", "").strip()
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return 0.0


def limpar_texto_pdf(texto: str) -> str:
    if not texto:
        return ""
    return texto.encode("latin-1", "replace").decode("latin-1")


def gerar_link_whatsapp(telefone: str, cliente: str, num_orcamento: int, total: float) -> str:
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
    return f"https://wa.me{digitos}?text={quote(mensagem)}"


# 4. CONEXÃO E OPERAÇÕES COM GOOGLE SHEETS
@st.cache_resource
def conectar_planilha():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    # Carrega as credenciais convertendo para dicionário do Python
    creds_dict = dict(st.secrets["gcp_service_account"])
    
    # Garante a formatação correta das quebras de linha na chave privada
    if "private_key" in creds_dict:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_url(st.secrets["sheet_url"])


def obter_ou_criar_aba(planilha, nome_aba: str, cabecalhos: list):
    try:
        ws = planilha.worksheet(nome_aba)
    except gspread.WorksheetNotFound:
        ws = planilha.add_worksheet(title=nome_aba, rows=1000, cols=len(cabecalhos) + 2)
        ws.append_row(cabecalhos)
    if not ws.row_values(1):
        ws.append_row(cabecalhos)
    return ws


def inicializar_estrutura_sheets():
    if "sheets_inicializado" in st.session_state:
        return
    planilha = conectar_planilha()

    ws_prod = obter_ou_criar_aba(
        planilha,
        "Produtos_Estoque",
        ["id", "nome", "tipo", "preco", "estoque_atual", "estoque_minimo"]
    )
    if len(ws_prod.get_all_values()) <= 1:
        ws_prod.append_rows([
            [1, "Instalação Ar-Condicionado 9000 BTUs", "Serviço", 350.0, 0, 0],
            [2, "Higienização e Limpeza Completa", "Serviço", 150.0, 0, 0],
            [3, "Carga de Gás R410A (Kg)", "Produto", 220.0, 15, 5],
            [4, "Suporte Reforçado 50cm (Par)", "Produto", 65.0, 20, 5],
            [5, "Tubulação de Cobre 1/4 (Metro)", "Produto", 38.0, 50, 15],
        ])

    obter_ou_criar_aba(
        planilha,
        "Orcamentos",
        [
            "id", "data_emissao", "cliente", "contato", "endereco", "documento",
            "itens_json", "subtotal", "desconto", "total", "condicao_pagamento",
            "observacoes", "status_servico", "estoque_baixado"
        ]
    )

    obter_ou_criar_aba(
        planilha,
        "Financeiro",
        [
            "id_lancamento", "id_orcamento", "data_aprovacao", "cliente",
            "valor_total", "forma_pagamento", "status_pagamento", "data_pagamento"
        ]
    )

    obter_ou_criar_aba(
        planilha,
        "Clientes",
        ["nome", "contato", "endereco", "documento"]
    )
    st.session_state.sheets_inicializado = True


# --- FUNÇÕES DE PRODUTOS E ESTOQUE ---

def listar_produtos_sheets() -> list:
    ws = conectar_planilha().worksheet("Produtos_Estoque")
    registros = ws.get_all_records()
    lista = []
    for idx, r in enumerate(registros, start=2):
        lista.append({
            "linha": idx,
            "id": int(r.get("id", idx - 1)),
            "nome": str(r.get("nome", "")),
            "tipo": str(r.get("tipo", "Serviço")),
            "preco": parse_float(r.get("preco", 0)),
            "estoque_atual": int(r.get("estoque_atual", 0)),
            "estoque_minimo": int(r.get("estoque_minimo", 0))
        })
    return lista


def adicionar_produto_sheets(nome: str, tipo: str, preco: float, estoque_inicial: int, estoque_min: int) -> bool:
    ws = conectar_planilha().worksheet("Produtos_Estoque")
    registros = ws.get_all_records()
    
    if any(str(r.get("nome", "")).strip().lower() == nome.strip().lower() for r in registros):
        return False
        
    novo_id = max([int(r.get("id", 0)) for r in registros], default=0) + 1
    ws.append_row([novo_id, nome.strip(), tipo, preco, estoque_inicial, estoque_min])
    return True


def atualizar_produto_sheets(linha: int, preco: float, estoque: int, minimo: int):
    ws = conectar_planilha().worksheet("Produtos_Estoque")
    ws.update(range_name=f"D{linha}:F{linha}", values=[[preco, estoque, minimo]])


def excluir_produto_sheets(linha: int):
    ws = conectar_planilha().worksheet("Produtos_Estoque")
    ws.delete_rows(linha)


def movimentar_estoque_por_orcamento(itens: list, operacao: str = "baixar"):
    ws = conectar_planilha().worksheet("Produtos_Estoque")
    registros = ws.get_all_records()
    
    for item in itens:
        for idx, r in enumerate(registros, start=2):
            if str(r.get("nome", "")).strip().lower() == str(item.get("produto", "")).strip().lower():
                if r.get("tipo") == "Produto":
                    estoque_atual = int(r.get("estoque_atual", 0))
                    qtd_movimentar = int(item.get("qtd", 0))
                    
                    if operacao == "baixar":
                        novo_saldo = estoque_atual - qtd_movimentar
                    else:
                        novo_saldo = estoque_atual + qtd_movimentar
                        
                    ws.update_cell(idx, 5, novo_saldo)
                break

def listar_produtos_sheets() -> list:
    ws = conectar_planilha().worksheet("Produtos_Estoque")
    registros = ws.get_all_records()
    resultado = []
    for idx, r in enumerate(registros, start=2):
        resultado.append({
            "linha": idx,
            "id": int(r.get("id", idx - 1)),
            "nome": str(r.get("nome", "")),
            "tipo": str(r.get("tipo", "Serviço")),
            "preco": parse_float(r.get("preco", 0)),
            "estoque_atual": int(parse_float(r.get("estoque_atual", 0))),
            "estoque_minimo": int(parse_float(r.get("estoque_minimo", 0))),
        })
    return resultado


def adicionar_produto_sheets(nome: str, tipo: str, preco: float, estoque: int, minimo: int) -> bool:
    produtos = listar_produtos_sheets()
    if any(p["nome"].lower() == nome.strip().lower() for p in produtos):
        return False
    novo_id = max([p["id"] for p in produtos], default=0) + 1
    ws = conectar_planilha().worksheet("Produtos_Estoque")
    ws.append_row([novo_id, nome.strip(), tipo, preco, estoque if tipo == "Produto" else 0, minimo if tipo == "Produto" else 0])
    return True


def atualizar_produto_sheets(linha: int, preco: float, estoque: int, minimo: int):
    ws = conectar_planilha().worksheet("Produtos_Estoque")
    ws.update(f"D{linha}:F{linha}", [[preco, estoque, minimo]])


def excluir_produto_sheets(linha: int):
    ws = conectar_planilha().worksheet("Produtos_Estoque")
    ws.delete_rows(linha)


def movimentar_estoque_por_orcamento(itens_lista: list, operacao: str):
    """operacao = 'baixar' (subtrai) ou 'estornar' (devolve ao estoque)."""
    produtos = listar_produtos_sheets()
    mapa_produtos = {p["nome"]: p for p in produtos}
    ws = conectar_planilha().worksheet("Produtos_Estoque")

    for item in itens_lista:
        nome_item = item["produto"]
        qtd_usada = int(item["qtd"])
        prod_db = mapa_produtos.get(nome_item)
        if prod_db and prod_db["tipo"] == "Produto":
            if operacao == "baixar":
                novo_saldo = max(0, prod_db["estoque_atual"] - qtd_usada)
            else:
                novo_saldo = prod_db["estoque_atual"] + qtd_usada
            ws.update_cell(prod_db["linha"], 5, novo_saldo)


    # --- FUNÇÕES DE CLIENTES --- 
def listar_clientes_sheets() -> list:
    ws = conectar_planilha().worksheet("Clientes")
    return ws.get_all_records()

def salvar_cliente_sheets(nome: str, contato: str, endereco: str, documento: str):
    if not nome.strip():
        return
    ws = conectar_planilha().worksheet("Clientes")
    registros = ws.get_all_records()
    for idx, c in enumerate(registros, start=2):
        if str(c.get("nome", "")).strip().lower() == nome.strip().lower():
            ws.update(f"A{idx}:D{idx}", [[nome.strip(), contato.strip(), endereco.strip(), documento.strip()]])
            return
    ws.append_row([nome.strip(), contato.strip(), endereco.strip(), documento.strip()])

# --- FUNÇÕES DE ORÇAMENTOS E FINANCEIRO ---

def salvar_orcamento_sheets(dados: dict) -> int:
    salvar_cliente_sheets(dados["cliente"], dados["contato"], dados["endereco"], dados["documento"])
    ws = conectar_planilha().worksheet("Orcamentos")
    registros = ws.get_all_records()
    novo_id = max([int(r.get("id", 0)) for r in registros], default=0) + 1
    ws.append_row([
        novo_id,
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
        dados["observacoes"],
        "Pendente",
        "Não"
    ])
    return novo_id

def listar_orcamentos_sheets() -> list:
    ws = conectar_planilha().worksheet("Orcamentos")
    registros = ws.get_all_records()
    lista = []
    for idx, r in enumerate(registros, start=2):
        lista.append({
            "linha": idx,
            "id": int(r.get("id", idx - 1)),
            "data_emissao": str(r.get("data_emissao", "")),
            "cliente": str(r.get("cliente", "")),
            "contato": str(r.get("contato", "")),
            "endereco": str(r.get("endereco", "")),
            "documento": str(r.get("documento", "")),
            "itens_json": str(r.get("itens_json", "[]")),
            "subtotal": parse_float(r.get("subtotal", 0)),
            "desconto": parse_float(r.get("desconto", 0)),
            "total": parse_float(r.get("total", 0)),
            "condicao_pagamento": str(r.get("condicao_pagamento", "")),
            "observacoes": str(r.get("observacoes", "")),
            "status_servico": str(r.get("status_servico", "Pendente")),
            "estoque_baixado": str(r.get("estoque_baixado", "Não")),
        })
    return list(reversed(lista))

def listar_financeiro_sheets() -> list:
    ws = conectar_planilha().worksheet("Financeiro")
    registros = ws.get_all_records()
    lista = []
    for idx, r in enumerate(registros, start=2):
        lista.append({
            "linha": idx,
            "id_lancamento": int(r.get("id_lancamento", idx - 1)),
            "id_orcamento": int(r.get("id_orcamento", 0)),
            "data_aprovacao": str(r.get("data_aprovacao", "")),
            "cliente": str(r.get("cliente", "")),
            "valor_total": parse_float(r.get("valor_total", 0)),
            "forma_pagamento": str(r.get("forma_pagamento", "")),
            "status_pagamento": str(r.get("status_pagamento", "A Receber")),
            "data_pagamento": str(r.get("data_pagamento", "")),
        })
    return list(reversed(lista))

def alterar_status_orcamento_e_integrar(orc: dict, novo_status: str):
    """Atualiza status do serviço, movimenta estoque e cria/remove lançamento financeiro."""
    planilha = conectar_planilha()
    ws_orc = planilha.worksheet("Orcamentos")
    ws_fin = planilha.worksheet("Financeiro")
    itens_lista = json.loads(orc["itens_json"])
    estoque_ja_baixado = (orc["estoque_baixado"] == "Sim")
    
    if novo_status == "Aprovado":
        # 1. Baixa automática no estoque (se ainda não foi baixado)
        if not estoque_ja_baixado:
            movimentar_estoque_por_orcamento(itens_lista, operacao="baixar")
            ws_orc.update_cell(orc["linha"], 14, "Sim")
            
        # 2. Cria lançamento na aba Financeiro como 'A Receber' (se ainda não existir)
        fin_registros = ws_fin.get_all_records()
        ja_no_financeiro = any(int(f.get("id_orcamento", 0)) == orc["id"] for f in fin_registros)
        if not ja_no_financeiro:
            novo_id_fin = max([int(f.get("id_lancamento", 0)) for f in fin_registros], default=0) + 1
            ws_fin.append_row([
                novo_id_fin,
                orc["id"],
                datetime.now().strftime("%d/%m/%Y"),
                orc["cliente"],
                orc["total"],
                orc["condicao_pagamento"],
                "A Receber",
                ""
            ])
    else:
        # Se deixou de ser Aprovado (voltou para Pendente ou Recusado), estorna o estoque
        if estoque_ja_baixado:
            movimentar_estoque_por_orcamento(itens_lista, operacao="estornar")
            ws_orc.update_cell(orc["linha"], 14, "Não")
            
        # Remove do financeiro se ainda estiver como 'A Receber'
        fin_registros = ws_fin.get_all_records()
        for idx, f in enumerate(fin_registros, start=2):
            if int(f.get("id_orcamento", 0)) == orc["id"] and f.get("status_pagamento") == "A Receber":
                ws_fin.delete_rows(idx)
                break
                
    ws_orc.update_cell(orc["linha"], 13, novo_status)

def atualizar_pagamento_financeiro(linha_fin: int, novo_status_pgto: str):
    ws_fin = conectar_planilha().worksheet("Financeiro")
    data_pgto = datetime.now().strftime("%d/%m/%Y") if novo_status_pgto == "Pago" else ""
    ws_fin.update(range_name=f"G{linha_fin}:H{linha_fin}", values=[[novo_status_pgto, data_pgto]])

# 5. CLASSE PARA GERAR O PDF DA ELTECH

class PDF_Eltech(FPDF):
    def header(self):
        if os.path.exists(LOGO_PATH):
            self.image(LOGO_PATH, 10, 8, 62)
        else:
            self.set_font("Arial", "B", 16)
            self.set_text_color(16, 44, 116)
            self.cell(0, 10, limpar_texto_pdf("ELTECH CLIMATIZAÇÃO"), ln=False)
            
        self.set_font("Arial", "", 9)
        self.set_text_color(50, 50, 50)
        self.set_x(-105)
        
        cabecalho_dir = (
            "ELTECH CLIMATIZAÇÃO\n"
            "E-mail: contato@eltechclimatizacao.com.br\n"
            "Telefone/WhatsApp: (51) 99232-9339\n"
            "Soluções em Conforto Térmico"
        )
        self.multi_cell(95, 4, limpar_texto_pdf(cabecalho_dir), align="R")
        self.ln(11)
        self.set_draw_color(16, 44, 116)
        self.set_line_width(0.6)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(5)
        
    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.set_text_color(120, 120, 120)
        self.cell(0, 10, limpar_texto_pdf(f"ELTECH CLIMATIZAÇÃO - Página {self.page_no()}"), align="C")

def gerar_pdf_documento(
    num_orcamento, data_emissao, cliente, contato, endereco, documento,
    itens, subtotal, desconto, total, condicao_pagamento, obs_adicional,
    modo_os: bool = False
):
    pdf = PDF_Eltech()
    pdf.add_page()
    pdf.ln(2)
    pdf.set_font("Arial", "B", 14)
    pdf.set_text_color(16, 44, 116)
    
    titulo_doc = (
        f"ORDEM DE SERVIÇO / COMPROVANTE  -  Nº {num_orcamento:04d}"
        if modo_os else
        f"ORÇAMENTO DE SERVIÇOS / PRODUTOS  -  Nº {num_orcamento:04d}"
    )
    pdf.cell(0, 8, limpar_texto_pdf(titulo_doc), ln=True, align="L")
    
    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(0, 5, limpar_texto_pdf(f"Data de Emissão: {data_emissao}"), ln=True, align="L")
    pdf.ln(4)
    
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
    pdf.set_font("Arial", "B", 10)
    pdf.set_text_color(255, 255, 255)
    pdf.set_fill_color(40, 40, 40)
    pdf.cell(95, 8, limpar_texto_pdf(" Descrição Técnica do Serviço / Equipamento"), border=1, fill=True)
    pdf.cell(20, 8, "Qtd", border=1, align="C", fill=True)
    pdf.cell(35, 8, limpar_texto_pdf("Preço Unitário"), border=1, align="R", fill=True)
    pdf.cell(40, 8, "Total Item", border=1, align="R", fill=True, ln=True)
    
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
    label_total = " VALOR TOTAL DO SERVIÇO:" if modo_os else " VALOR TOTAL DO INVESTIMENTO:"
    pdf.cell(130, 10, label_total, border=1, fill=True)
    pdf.cell(60, 10, f"{formatar_moeda(total)} ", border=1, align="R", fill=True, ln=True)
    pdf.ln(8)
    
    pdf.set_font("Arial", "B", 9)
    pdf.set_text_color(60, 60, 60)
    pdf.cell(0, 5, limpar_texto_pdf(f"* Condições de pagamento: {condicao_pagamento}"), ln=True)
    
    pdf.set_font("Arial", "I", 9)
    pdf.set_text_color(100, 100, 100)
    if modo_os:
        pdf.cell(0, 5, limpar_texto_pdf("* Declaramos que os serviços/equipamentos acima foram entregues e testados."), ln=True)
    else:
        pdf.cell(0, 5, limpar_texto_pdf(f"* Validade deste orçamento: 10 dias a contar de {data_emissao}."), ln=True)
        
    pdf.cell(0, 5, limpar_texto_pdf("* Garantia de instalação de equipamentos novos e usados: 3 meses."), ln=True)
    
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

# 6. AUTENTICAÇÃO

if "itens_orcamento" not in st.session_state:
    st.session_state.itens_orcamento = []
if "ultimo_orcamento_gerado" not in st.session_state:
    st.session_state.ultimo_orcamento_gerado = None
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if "usuario_logado" not in st.session_state:
    st.session_state.usuario_logado = ""

USUARIOS_SISTEMA = {
    "george": ("21122018", "George - Administrador"),
    "eltech": ("21122018", "Equipe ELTECH"),
}

if not st.session_state.autenticado:
    _, col_login, _ = st.columns([1, 1.3, 1])
    with col_login:
        st.write("")
        if os.path.exists(LOGO_PATH):
            with open(LOGO_PATH, "rb") as f:
                b64_logo = base64.b64encode(f.read()).decode()
            logo_login_html = (
                f'<img src="data:image/png;base64,{b64_logo}" '
                f'style="height:54px; background:white; padding:8px 14px; border-radius:10px; margin-bottom:8px;">'
            )
        else:
            logo_login_html = '❄️'
            
        st.markdown(f"""
        {logo_login_html}
        ELTECH CLIMATIZAÇÃO
        Portal Administrativo e Comercial
        """, unsafe_allow_html=True)
        
        with st.container(border=True):
            st.markdown("#### 🔐 Acesso ao Sistema")
            with st.form("form_login"):
                usuario_digitado = st.text_input("Usuário", placeholder="Digite seu login (ex: george)").strip().lower()
                senha_digitada = st.text_input("Senha", type="password", placeholder="Digite sua senha")
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
# 7. INICIALIZA CONEXÃO GOOGLE SHEETS APÓS LOGIN

try:
    inicializar_estrutura_sheets()
except Exception as e:
    st.error("⚠️ Não foi possível conectar à planilha do Google Sheets. Verifique o arquivo .streamlit/secrets.toml e o compartilhamento da planilha.")
    st.code(str(e))
    st.stop()

# 8. INTERFACE PRINCIPAL

exibir_cabecalho_corporativo()
col_bem_vindo, col_sync, col_sair = st.columns([4.5, 1.5, 1.1])

with col_bem_vindo:
    st.caption(f"👤 Conectado como: {st.session_state.usuario_logado}")
    
with col_sync:
    if st.button("🔄 Atualizar Planilha", use_container_width=True):
        st.rerun()
        
with col_sair:
    if st.button("🔒 Sair", use_container_width=True):
        st.session_state.autenticado = False
        st.session_state.usuario_logado = ""
        st.rerun()

aba1, aba2, aba3, aba4 = st.tabs([
    "📄 Novo Orçamento",
    "🗂️ Histórico de Orçamentos",
    "💰 Controle Financeiro",
    "📦 Produtos & Estoque"
])

# ==============================================================================
# ABA 1: NOVO ORÇAMENTO
# ==============================================================================

with aba1:
    st.header("Construir Novo Orçamento")
    clientes_cadastrados = listar_clientes_sheets()
    
    if clientes_cadastrados:
        nomes_clientes = ["➕ Novo Cliente (Digitar manualmente)"] + [c["nome"] for c in clientes_cadastrados]
        cliente_escolhido_lista = st.selectbox(
            "Preenchimento Rápido de Cliente:",
            nomes_clientes,
            key="seletor_cliente_rapido"
        )
        
        if cliente_escolhido_lista != "➕ Novo Cliente (Digitar manualmente)":
            if st.session_state.get("ultimo_cliente_selecionado") != cliente_escolhido_lista:
                dados_cli = next((c for c in clientes_cadastrados if c["nome"] == cliente_escolhido_lista), None)
                if dados_cli:
                    st.session_state.input_cliente = str(dados_cli.get("nome", ""))
                    st.session_state.input_contato = str(dados_cli.get("contato", ""))
                    st.session_state.input_endereco = str(dados_cli.get("endereco", ""))
                    st.session_state.input_documento = str(dados_cli.get("documento", ""))
                    st.session_state.ultimo_cliente_selecionado = cliente_escolhido_lista
                    st.rerun()
                    
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
        
    produtos_lista = listar_produtos_sheets()
    mapa_produtos = {p["nome"]: p for p in produtos_lista}
    
    st.markdown("### Adicionar Itens")
    tipo_item = st.radio(
        "Origem do item:",
        ["Selecionar do Catálogo", "Item Avulso (Digitar na hora)"],
        horizontal=True
    )
    
    if tipo_item == "Selecionar do Catálogo":
        if produtos_lista:
            col_prod, col_qtd, col_preco = st.columns([3, 1, 1.5])
            with col_prod:
                produto_selecionado = st.selectbox(
                    "Serviço / Produto",
                    list(mapa_produtos.keys()),
                    key="produto_atual"
                )
                prod_info = mapa_produtos[produto_selecionado]
                if prod_info["tipo"] == "Produto":
                    st.caption(f"📦 Item físico controlado no estoque — Saldo disponível: {prod_info['estoque_atual']} un.")
            with col_qtd:
                quantidade = st.number_input("Qtd", min_value=1, value=1, step=1, key="qtd_cat")
            with col_preco:
                preco_unitario = st.number_input(
                    "Valor Unit. (R$)",
                    min_value=0.0,
                    value=float(prod_info["preco"]),
                    step=10.0,
                    key="preco_cat"
                )
                
            if st.button("➕ Adicionar Item ao Orçamento", use_container_width=True):
                st.session_state.itens_orcamento.append({
                    "produto": produto_selecionado,
                    "tipo": prod_info["tipo"],
                    "qtd": quantidade,
                    "preco_unit": preco_unitario,
                    "total_item": preco_unitario * quantidade
                })
                st.session_state.ultimo_orcamento_gerado = None
                st.toast(f"'{produto_selecionado}' adicionado!", icon="✅")
                st.rerun()
        else:
            st.warning("Nenhum produto cadastrado no catálogo.")
    else:
        col_av_desc, col_av_qtd, col_av_preco = st.columns([3, 1, 1.5])
        with col_av_desc:
            avulso_nome = st.text_input("Descrição do serviço/peça avulsa", placeholder="Ex: Mão de obra adicional")
        with col_av_qtd:
            avulso_qtd = st.number_input("Qtd", min_value=1, value=1, step=1, key="qtd_av")
        with col_av_preco:
            avulso_preco = st.number_input("Valor Unit. (R$)", min_value=0.0, value=0.0, step=10.0, key="preco_av")
            
        if st.button("➕ Adicionar Item Avulso", use_container_width=True):
            if avulso_nome.strip():
                st.session_state.itens_orcamento.append({
                    "produto": avulso_nome.strip(),
                    "tipo": "Avulso",
                    "qtd": avulso_qtd,
                    "preco_unit": avulso_preco,
                    "total_item": avulso_preco * avulso_qtd
                })
                st.session_state.ultimo_orcamento_gerado = None
                st.toast(f"'{avulso_nome.strip()}' adicionado!", icon="✅")
                st.rerun()
            else:
                st.warning("Digite a descrição do item avulso.")
                
    if st.session_state.itens_orcamento:
        st.markdown("---")
        st.markdown("### Itens no Orçamento")
        subtotal_calculado = 0.0
        
        for idx, item in enumerate(st.session_state.itens_orcamento):
            c_desc, c_val, c_del = st.columns([4, 2, 0.7])
            with c_desc:
                tag_tipo = "📦" if item.get("tipo") == "Produto" else "🔧"
                st.write(f"{tag_tipo} {item['produto']} (x{item['qtd']} de {formatar_moeda(item['preco_unit'])})")
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
                desconto_calculado = st.number_input(
                    "Valor do Desconto (R$)",
                    min_value=0.0,
                    max_value=float(subtotal_calculado),
                    value=0.0,
                    step=10.0
                )
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
            st.write(f"Subtotal: {formatar_moeda(subtotal_calculado)} | Desconto: -{formatar_moeda(desconto_calculado)}")
        st.subheader(f"💰 Valor Total Final: {formatar_moeda(total_final)}")
        
        if cliente.strip():
            if st.button("✅ Salvar Orçamento na Planilha e Gerar PDF", type="primary", use_container_width=True):
                with st.spinner("Gravando no Google Sheets e gerando PDF..."):
                    data_hoje = datetime.now().strftime("%d/%m/%Y")
                    num_gerado = salvar_orcamento_sheets({
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
                    obs_adicional=obs_tecnica.strip(),
                    modo_os=False
                )
                st.session_state.ultimo_orcamento_gerado = {
                    "numero": num_gerado,
                    "cliente": cliente.strip(),
                    "contato": contato.strip(),
                    "total": total_final,
                    "pdf": pdf_bytes
                }
                st.success(f"Orçamento #{num_gerado:04d} salvo no Google Sheets!")
                
        if st.session_state.ultimo_orcamento_gerado:
            dados_gerados = st.session_state.ultimo_orcamento_gerado
            nome_arq = f"Orcamento_{dados_gerados['numero']:04d}ELTECH{dados_gerados['cliente'].replace(' ', '_')}.pdf"
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
            st.warning("⚠️ Preencha o Nome do Cliente no topo para salvar e gerar o PDF.")
# ==============================================================================
# ABA 2: HISTÓRICO DE ORÇAMENTOS (COMERCIAL + BAIXA AUTOMÁTICA DE ESTOQUE)
# ==============================================================================

with aba2:
    st.header("Histórico Comercial de Orçamentos")
    st.caption("Ao mudar um orçamento para Aprovado, o sistema dá baixa automática nos produtos do estoque e envia o valor para o Controle Financeiro.")
    
    col_busca, col_status_filtro = st.columns([2.5, 1.5])
    with col_busca:
        busca_hist = st.text_input("Buscar por cliente ou número", placeholder="Ex: João ou 12")
    with col_status_filtro:
        filtro_status = st.selectbox("Status do Orçamento", ["Todos", "Pendente", "Aprovado", "Recusado"])
        
    todos_orcamentos = listar_orcamentos_sheets()
    orcamentos_filtrados = [
        o for o in todos_orcamentos
        if (filtro_status == "Todos" or o["status_servico"] == filtro_status)
        and (not busca_hist.strip() or busca_hist.strip().lower() in o["cliente"].lower() or busca_hist.strip() in str(o["id"]))
    ]
    
    # Indicadores Comerciais (Propostas)
    with st.container(border=True):
        c_com1, c_com2, c_com3, c_com4 = st.columns(4)
        qtd_tot = len(todos_orcamentos)
        qtd_apr = sum(1 for o in todos_orcamentos if o["status_servico"] == "Aprovado")
        qtd_pen = sum(1 for o in todos_orcamentos if o["status_servico"] == "Pendente")
        taxa_apr = (qtd_apr / qtd_tot * 100) if qtd_tot > 0 else 0.0
        
        c_com1.metric("Propostas Emitidas", f"{qtd_tot}")
        c_com2.metric("Serviços Aprovados", f"{qtd_apr}")
        c_com3.metric("Aguardando Resposta", f"{qtd_pen}")
        c_com4.metric("Taxa de Aprovação", f"{taxa_apr:.1f}%")
        
    st.markdown("---")
    
    if orcamentos_filtrados:
        for orc in orcamentos_filtrados:
            icone_st = {"Pendente": "🟡", "Aprovado": "🟢", "Recusado": "🔴"}.get(orc["status_servico"], "⚪")
            badge_est = " | 📦 Estoque Baixado" if orc["estoque_baixado"] == "Sim" else ""
            titulo_exp = (
                f"{icone_st} #{orc['id']:04d} | {orc['cliente']} "
                f"— {formatar_moeda(orc['total'])} ({orc['data_emissao']}){badge_est}"
            )
            
            with st.expander(titulo_exp):
                c_info1, c_info2 = st.columns(2)
                with c_info1:
                    st.write(f"Contato: {orc['contato'] or 'Não informado'}")
                    st.write(f"CPF/CNPJ: {orc['documento'] or 'Não informado'}")
                with c_info2:
                    st.write(f"Endereço: {orc['endereco'] or 'Não informado'}")
                    st.write(f"Pagamento Previsto: {orc['condicao_pagamento'] or 'A combinar'}")
                    
                itens_lista = json.loads(orc["itens_json"])
                st.markdown("Itens do orçamento:")
                for it in itens_lista:
                    st.write(f"- {it['produto']} (x{it['qtd']}) — {formatar_moeda(it['total_item'])}")
                    
                if orc["desconto"] > 0:
                    st.write(f"Desconto concedido: {formatar_moeda(orc['desconto'])}")
                if orc["observacoes"]:
                    st.write(f"Observações: {orc['observacoes']}")
                    
                st.markdown("---")
                col_st, col_reimp, col_dup = st.columns([1.4, 1.5, 1.5])
                
                with col_st:
                    opcoes_st = ["Pendente", "Aprovado", "Recusado"]
                    idx_atual = opcoes_st.index(orc["status_servico"]) if orc["status_servico"] in opcoes_st else 0
                    novo_st = st.selectbox(
                        "Status",
                        opcoes_st,
                        index=idx_atual,
                        key=f"st_orc_{orc['id']}",
                        label_visibility="collapsed"
                    )
                    if novo_st != orc["status_servico"]:
                        with st.spinner("Atualizando status, estoque e financeiro..."):
                            alterar_status_orcamento_e_integrar(orc, novo_st)
                        st.toast(f"Orçamento #{orc['id']:04d} alterado para {novo_st}!")
                        st.rerun()
                        
                with col_reimp:
                    pdf_reimp = gerar_pdf_documento(
                        num_orcamento=orc["id"],
                        data_emissao=orc["data_emissao"],
                        cliente=orc["cliente"],
                        contato=orc["contato"],
                        endereco=orc["endereco"],
                        documento=orc["documento"],
                        itens=itens_lista,
                        subtotal=orc["subtotal"],
                        desconto=orc["desconto"],
                        total=orc["total"],
                        condicao_pagamento=orc["condicao_pagamento"] or "A combinar",
                        obs_adicional=orc["observacoes"],
                        modo_os=False
                    )
                    st.download_button(
                        label="🖨️ Baixar Orçamento",
                        data=pdf_reimp,
                        file_name=f"Orcamento_{orc['id']:04d}ELTECH{orc['cliente'].replace(' ', '')}.pdf",
                        mime="application/pdf",
                        key=f"dl_hist{orc['id']}",
                        use_container_width=True
                    )
                    
                with col_dup:
                    if st.button("📋 Carregar p/ Editar", key=f"dup_hist_{orc['id']}", use_container_width=True):
                        st.session_state.input_cliente = orc["cliente"]
                        st.session_state.input_contato = orc["contato"]
                        st.session_state.input_endereco = orc["endereco"]
                        st.session_state.input_documento = orc["documento"]
                        st.session_state.itens_orcamento = itens_lista
                        st.session_state.ultimo_orcamento_gerado = None
                        st.toast("Dados carregados na aba 'Novo Orçamento'!", icon="📋")
                        
                if orc["status_servico"] == "Aprovado":
                    pdf_os = gerar_pdf_documento(
                        num_orcamento=orc["id"],
                        data_emissao=datetime.now().strftime("%d/%m/%Y"),
                        cliente=orc["cliente"],
                        contato=orc["contato"],
                        endereco=orc["endereco"],
                        documento=orc["documento"],
                        itens=itens_lista,
                        subtotal=orc["subtotal"],
                        desconto=orc["desconto"],
                        total=orc["total"],
                        condicao_pagamento=orc["condicao_pagamento"] or "A combinar",
                        obs_adicional=orc["observacoes"],
                        modo_os=True
                    )
                    st.download_button(
                        label="📋 Gerar Ordem de Serviço / Comprovante (PDF)",
                        data=pdf_os,
                        file_name=f"OS_{orc['id']:04d}ELTECH{orc['cliente'].replace(' ', '')}.pdf",
                        mime="application/pdf",
                        key=f"dl_os{orc['id']}",
                        type="primary",
                        use_container_width=True
                    )
    else:
        st.info("Nenhum orçamento encontrado com os filtros atuais.")
# ==============================================================================
# ABA 3: CONTROLE FINANCEIRO (A RECEBER vs. PAGO)
# ==============================================================================

with aba3:
    st.header("Controle Financeiro (Contas a Receber & Caixa)")
    st.caption("Apenas serviços Aprovados aparecem aqui. Controle separadamente o que já entrou no caixa e o que ainda falta receber.")
    
    lancamentos_fin = listar_financeiro_sheets()
    if lancamentos_fin:
        total_aprovado_geral = sum(f["valor_total"] for f in lancamentos_fin)
        total_ja_pago = sum(f["valor_total"] for f in lancamentos_fin if f["status_pagamento"] == "Pago")
        total_a_receber = sum(f["valor_total"] for f in lancamentos_fin if f["status_pagamento"] != "Pago")
        
        with st.container(border=True):
            f1, f2, f3 = st.columns(3)
            f1.metric("Total Serviços Aprovados", formatar_moeda(total_aprovado_geral))
            f2.metric("✅ Já Recebido (Caixa)", formatar_moeda(total_ja_pago))
            f3.metric("⏳ A Receber (Pendente)", formatar_moeda(total_a_receber))
            
        st.markdown("---")
        filtro_fin = st.radio(
            "Filtrar lançamentos:",
            ["Todos", "A Receber", "Parcial", "Pago"],
            horizontal=True
        )
        
        lancamentos_filtrados = [
            f for f in lancamentos_fin
            if filtro_fin == "Todos" or f["status_pagamento"] == filtro_fin
        ]
        
        for fin in lancamentos_filtrados:
            icone_pg = {"Pago": "✅", "Parcial": "🌗", "A Receber": "⏳"}.get(fin["status_pagamento"], "⏳")
            with st.container(border=True):
                c_f1, c_f2, c_f3 = st.columns([3, 1.8, 1.5])
                with c_f1:
                    st.markdown(f"{icone_pg} Orçamento #{fin['id_orcamento']:04d} — {fin['cliente']}")
                    info_data = f"Aprovado em: {fin['data_aprovacao']} | Forma: {fin['forma_pagamento']}"
                    if fin["data_pagamento"]:
                        info_data += f" | Pago em: {fin['data_pagamento']}"
                    st.caption(info_data)
                with c_f2:
                    st.subheader(formatar_moeda(fin["valor_total"]))
                with c_f3:
                    opcoes_pg = ["A Receber", "Parcial", "Pago"]
                    idx_pg = opcoes_pg.index(fin["status_pagamento"]) if fin["status_pagamento"] in opcoes_pg else 0
                    novo_pg = st.selectbox(
                        "Pagamento",
                        opcoes_pg,
                        index=idx_pg,
                        key=f"pg_fin_{fin['id_lancamento']}",
                        label_visibility="collapsed"
                    )
                    if novo_pg != fin["status_pagamento"]:
                        with st.spinner("Atualizando financeiro..."):
                            atualizar_pagamento_financeiro(fin["linha"], novo_pg)
                        st.toast(f"Pagamento do Orçamento #{fin['id_orcamento']:04d} marcado como {novo_pg}!")
                        st.rerun()
    else:
        st.info("Nenhum serviço aprovado ainda. Quando você aprovar um orçamento no Histórico, ele aparecerá automaticamente aqui para controle de recebimento.")

# ==============================================================================
# ABA 4: PRODUTOS E CONTROLE DE ESTOQUE
# ==============================================================================

with aba4:
    st.header("Catálogo de Serviços e Controle de Estoque")
    produtos_estoque = listar_produtos_sheets()
    
    # Alerta de estoque baixo
    itens_criticos = [
        p for p in produtos_estoque
        if p["tipo"] == "Produto" and p["estoque_atual"] <= p["estoque_minimo"]
    ]
    if itens_criticos:
        nomes_criticos = ", ".join([f"{p['nome']} ({p['estoque_atual']} un.)" for p in itens_criticos])
        st.error(f"⚠️ Alerta de Estoque Baixo: Os seguintes materiais precisam de reposição: {nomes_criticos}")
        
    sub_acao = st.radio(
        "Selecione a ação:",
        ["Tabela Geral & Reposição de Estoque", "Cadastrar Novo Produto / Serviço"],
        horizontal=True
    )
    
    if sub_acao == "Tabela Geral & Reposição de Estoque":
        if produtos_estoque:
            tabela_exib = []
            for p in produtos_estoque:
                if p["tipo"] == "Produto":
                    status_est = "🔴 Baixo" if p["estoque_atual"] <= p["estoque_minimo"] else "🟢 OK"
                    saldo_str = f"{p['estoque_atual']} un. (Mín: {p['estoque_minimo']})"
                else:
                    status_est = "➖ Serviço"
                    saldo_str = "Não controla"
                    
                tabela_exib.append({
                    "ID": p["id"],
                    "Tipo": p["tipo"],
                    "Descrição": p["nome"],
                    "Preço": formatar_moeda(p["preco"]),
                    "Estoque": saldo_str,
                    "Situação": status_est
                })
            st.dataframe(tabela_exib, use_container_width=True, hide_index=True)
            
            st.markdown("---")
            st.subheader("Atualizar Preço ou Dar Entrada no Estoque")
            
            opcoes_ed = {f"[{p['tipo']}] {p['nome']} ({formatar_moeda(p['preco'])})": p for p in produtos_estoque}
            sel_ed = st.selectbox("Selecione o item para atualizar:", list(opcoes_ed.keys()))
            prod_sel = opcoes_ed[sel_ed]
            
            c_ed1, c_ed2, c_ed3 = st.columns(3)
            with c_ed1:
                novo_preco_ed = st.number_input(
                    "Preço Comercial (R$)",
                    min_value=0.0,
                    value=float(prod_sel["preco"]),
                    step=10.0,
                    key=f"ed_p_{prod_sel['id']}"
                )
            with c_ed2:
                novo_est_ed = st.number_input(
                    "Saldo Atual em Estoque",
                    min_value=0,
                    value=int(prod_sel["estoque_atual"]),
                    step=1,
                    disabled=(prod_sel["tipo"] != "Produto"),
                    key=f"ed_e_{prod_sel['id']}"
                )
            with c_ed3:
                novo_min_ed = st.number_input(
                    "Estoque Mínimo (Alerta)",
                    min_value=0,
                    value=int(prod_sel["estoque_minimo"]),
                    step=1,
                    disabled=(prod_sel["tipo"] != "Produto"),
                    key=f"ed_m_{prod_sel['id']}"
                )
                
            col_btn_upd, col_btn_del = st.columns(2)
            with col_btn_upd:
                if st.button("🔄 Salvar Alterações na Planilha", type="primary", use_container_width=True):
                    with st.spinner("Salvando..."):
                        atualizar_produto_sheets(prod_sel["linha"], novo_preco_ed, novo_est_ed, novo_min_ed)
                    st.success(f"'{prod_sel['nome']}' atualizado com sucesso!")
                    st.rerun()
            with col_btn_del:
                if st.button("🗑️ Excluir Item", use_container_width=True):
                    with st.spinner("Excluindo..."):
                        excluir_produto_sheets(prod_sel["linha"])
                    st.warning(f"'{prod_sel['nome']}' removido do catálogo.")
                    st.rerun()
        else:
            st.info("Nenhum item cadastrado.")
            
    elif sub_acao == "Cadastrar Novo Produto / Serviço":
        st.subheader("Incluir Novo Item no Catálogo")
        with st.form("form_novo_item_sheets", clear_on_submit=True):
            novo_tipo = st.radio(
                "Categoria do Item:",
                ["Serviço", "Produto"],
                horizontal=True,
                help="Escolha 'Produto' para materiais físicos que terão baixa automática no estoque (ex: Gás, Tubulação, Suporte)."
            )
            novo_nome = st.text_input("Nome descritivo (Ex: Carga de Gás R32 ou Instalação 12000 BTUs)")
            
            c_cad1, c_cad2, c_cad3 = st.columns(3)
            with c_cad1:
                novo_preco = st.number_input("Preço de Tabela (R$)", min_value=0.0, value=0.0, step=10.0)
            with c_cad2:
                novo_est = st.number_input("Estoque Inicial (se Produto)", min_value=0, value=0, step=1)
            with c_cad3:
                novo_min = st.number_input("Estoque Mínimo para Alerta", min_value=0, value=2, step=1)
                
            salvar_btn = st.form_submit_button("💾 Cadastrar na Planilha Google", type="primary")
            if salvar_btn:
                if novo_nome.strip():
                    with st.spinner("Gravando na planilha..."):
                        ok = adicionar_produto_sheets(novo_nome, novo_tipo, novo_preco, novo_est, novo_min)
                    if ok:
                        st.success(f"'{novo_nome}' cadastrado com sucesso!")
                        st.rerun()
                    else:
                        st.error("Já existe um item cadastrado com esse nome!")
                else:
                    st.warning("Preencha o nome do item.")
