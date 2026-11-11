import sqlite3
import streamlit as st
from fpdf import FPDF
from datetime import datetime
import os

# 1. CONFIGURAÇÃO E CRIAÇÃO DO BANCO DE DADOS
def conectar_banco():
    conn = sqlite3.connect("orcamentos.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS produtos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT UNIQUE,
            preco REAL
        )
    """)
    conn.commit()
    return conn, cursor

conn, cursor = conectar_banco()

# Insere os dados padrão se o banco estiver vazio
cursor.execute("SELECT COUNT(*) FROM produtos")
if cursor.fetchone() == 0:
    produtos_iniciais = [
        ("Instalação Ar-Condicionado 9000 BTUs", 350.0),
        ("Higienização e Limpeza Completa", 150.0),
        ("Carga de Gás R410A", 220.0),
        ("Manutenção Corretiva Sistema de Climatização", 250.0)
    ]
    cursor.executemany("INSERT INTO produtos (nome, preco) VALUES (?, ?)", produtos_iniciais)
    conn.commit()

# Inicializa o "carrinho" de itens do orçamento atual se não existir
if "itens_orcamento" not in st.session_state:
    st.session_state.itens_orcamento = []

# 2. CLASSE PARA GERAR O PDF DA ELTECH COM MÚLTIPLOS ITENS
class PDF_Eltech(FPDF):
    def header(self):
        if os.path.exists("logo.png"):
            self.image("logo.png", 10, 8, 65)
        else:
            self.set_font("Arial", "B", 16)
            self.set_text_color(16, 44, 116)
            self.cell(0, 10, "ELTECH CLIMATIZAÇÃO", ln=False)
        
        self.set_font("Arial", "", 9)
        self.set_text_color(50, 50, 50)
        self.set_x(-100)
        self.multi_cell(90, 4, "ELTECH CLIMATIZAÇÃO \nE-mail: contato@eltechclimatizacao.com.br\nTelefone/WhatsApp: (51) 99232-9339\nSoluções em Conforto Térmico", align="R")
        
        self.ln(12)
        self.set_draw_color(16, 44, 116)
        self.set_line_width(0.5)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.set_text_color(120, 120, 120)
        self.cell(0, 10, f"ELTECH CLIMATIZAÇÃO - Página {self.page_no()}", align="C")

def gerar_pdf_documento(cliente, contato, endereco, documento, itens, total, obs_adicional):
    pdf = PDF_Eltech()
    pdf.add_page()
    
    pdf.ln(5)
    pdf.set_font("Arial", "B", 14)
    pdf.set_text_color(16, 44, 116)
    pdf.cell(0, 10, "ORÇAMENTO DE SERVIÇOS / PRODUTOS", ln=True, align="L")
    
    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(80, 80, 80)
    data_atual = datetime.now().strftime("%d/%m/%Y")
    pdf.cell(0, 5, f"Data de Emissão: {data_atual}", ln=True, align="L")
    pdf.ln(5)
    
    pdf.set_font("Arial", "B", 11)
    pdf.set_text_color(255, 255, 255)
    pdf.set_fill_color(16, 44, 116)
    pdf.cell(0, 6, "  DADOS DO CLIENTE", ln=True, fill=True)
    
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", "", 11)
    pdf.ln(2)
    pdf.cell(0, 6, f"Nome / Razão Social: {cliente}", ln=True)
    pdf.cell(0, 6, f"CPF / CNPJ: {documento}", ln=True)
    pdf.cell(0, 6, f"Celular / WhatsApp: {contato}", ln=True)
    pdf.cell(0, 6, f"Endereço: {endereco}", ln=True)
    pdf.ln(5)
    
    # Cabeçalho da Tabela
    pdf.set_font("Arial", "B", 10)
    pdf.set_text_color(255, 255, 255)
    pdf.set_fill_color(40, 40, 40)
    
    pdf.cell(95, 7, " Descrição Técnica do Serviço / Equipamento", border=1, fill=True)
    pdf.cell(20, 7, "Qtd", border=1, align="C", fill=True)
    pdf.cell(35, 7, "Preço Unitário", border=1, align="R", fill=True)
    pdf.cell(40, 7, "Total Item", border=1, align="R", fill=True, ln=True)
    
    # Listagem de todos os itens adicionados
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", "", 10)
    for item in itens:
        pdf.cell(95, 8, f" {item['produto']}", border=1)
        pdf.cell(20, 8, str(item['qtd']), border=1, align="C")
        pdf.cell(35, 8, f"R$ {item['preco_unit']:,.2f}", border=1, align="R")
        pdf.cell(40, 8, f"R$ {item['total_item']:,.2f}", border=1, align="R", ln=True)
    
    if obs_adicional:
        pdf.ln(4)
        pdf.set_font("Arial", "B", 10)
        pdf.cell(0, 5, "Observações Técnicas / Escopo do Serviço:", ln=True)
        pdf.set_font("Arial", "", 10)
        pdf.multi_cell(0, 5, obs_adicional)
        
    pdf.ln(6)
    
    pdf.set_font("Arial", "B", 12)
    pdf.set_fill_color(230, 240, 255)
    pdf.cell(130, 10, " VALOR TOTAL DO INVESTIMENTO:", border=1, fill=True)
    pdf.cell(60, 10, f"R$ {total:,.2f} ", border=1, align="R", fill=True, ln=True)
    
    pdf.ln(10)
    pdf.set_font("Arial", "I", 9)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 5, "* Condições de pagamento: A combinar conforme execução.", ln=True)
    pdf.cell(0, 5, f"* Validade deste orçamento: 10 dias a contar de {data_atual}.", ln=True)
    pdf.cell(0, 5, "* Garantia de instalação de equipamentos novos e usados: 3 meses.", ln=True)
    
    return pdf.output()

# 3. INTERFACE COMPLETA DO SISTEMA
st.set_page_config(page_title="ELTECH Orçamentos", layout="centered")
# --- BLOQUEIO DE SEGURANÇA POR SENHA ---
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.title("🔒 Acesso Restrito - ELTECH")
    senha_digitada = st.text_input("Digite a senha administrativa:", type="password")
    if st.button("Entrar no Sistema"):
        if senha_digitada == "21122018": 
            st.session_state.autenticado = True
            st.rerun()
        else:
            st.error("Senha incorreta!")
    st.stop()  # Impede o restante do sistema de carregar sem a senha

st.title("❄️ ELTECH Climatização - Painel Administrativo")

aba1, aba2 = st.tabs(["📄 Central de Orçamentos", "⚙️ Banco de Dados de Produtos"])

with aba1:
    st.header("Construir Novo Orçamento")    
    cursor.execute("SELECT nome, preco FROM produtos")
    dados_banco = cursor.fetchall()
    dicionario_produtos = {nome: preco for nome, preco in dados_banco}
    
    col1, col2 = st.columns(2)
    with col1:
        cliente = st.text_input("Nome do Cliente / Empresa")
    with col2:
        contato = st.text_input("Celular / WhatsApp")
        
    col3, col4 = st.columns(2)
    with col3:
        endereco = st.text_input("Endereço")
    with col4:
        documento = st.text_input("CPF / CNPJ")

    st.markdown("### Adicionar Itens")
    if dados_banco:
        produto_selecionado = st.selectbox("Selecione o Serviço/Produto", list(dicionario_produtos.keys()), key="produto_atual")
        preco_unitario = dicionario_produtos[st.session_state.produto_atual]
        quantidade = st.number_input("Quantidade", min_value=1, value=1, step=1)
            
        if st.button("➕ Adicionar Item ao Orçamento"):
            total_item = preco_unitario * quantidade
            st.session_state.itens_orcamento.append({
                "produto": produto_selecionado,
                "qtd": quantidade,
                "preco_unit": preco_unitario,
                "total_item": total_item
            })
            st.success(f"Item '{produto_selecionado}' adicionado!")
            st.rerun()

        # Se houver itens no carrinho, exibe a lista e calcula o total
        if st.session_state.itens_orcamento:
            st.markdown("### Itens Adicionados:")
            total_calculado = 0.0
            
            for idx, item in enumerate(st.session_state.itens_orcamento):
                st.text(f"🔹 {item['produto']} (x{item['qtd']}) - R$ {item['total_item']:,.2f}")
                total_calculado += item['total_item']
            
            if st.button("❌ Limpar Itens"):
                st.session_state.itens_orcamento = []
                st.rerun()
                
            obs_tecnica = st.text_area("Notas ou Detalhes adicionais do orçamento (Opcional)", placeholder="Ex: Inclui infraestrutura de cobre.")
            
            st.markdown("---")
            st.subheader(f"💰 Valor Total Geral: R$ {total_calculado:,.2f}")
            
            if cliente:
                st.info("🔍 Tudo pronto! Revise a pré-visualização abaixo antes de gerar o documento definitivo.")
                
                with st.container(border=True):
                    st.markdown(f"**EMPRESA EMISSORA:** ELTECH CLIMATIZAÇÃO")
                    st.markdown(f"**DESTINATÁRIO:** {cliente} | **CPF/CNPJ:** {documento}")
                    st.markdown(f"**CONTATO:** {contato} | **ENDEREÇO:** {endereco}")
                    st.markdown(f"**QUANTIDADE DE ITENS:** {len(st.session_state.itens_orcamento)}")
                
                pdf_resultado = gerar_pdf_documento(cliente, contato, endereco, documento, st.session_state.itens_orcamento, total_calculado, obs_tecnica)
                dados_binarios_pdf = bytes(pdf_resultado)
                
                st.download_button(
                    label="🖨️ Direcionar para Impressão / Salvar PDF",
                    data=dados_binarios_pdf,
                    file_name=f"Orcamento_ELTECH_{cliente.replace(' ', '_')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            else:
                st.warning("Preencha o nome do cliente no topo para liberar o botão de impressão.")
    else:
        st.error("Nenhum produto encontrado no banco de dados. Cadastre itens na aba ao lado.")

with aba2:
    st.header("Gerenciamento do Banco de Dados")
    sub_acao = st.radio("Selecione a ação com o Banco de Dados:", ["Incluir Novo Produto", "Visualizar e Alterar Preços"], horizontal=True)
    
    if sub_acao == "Incluir Novo Produto":
        st.subheader("Cadastrar Novo Serviço ou Item")
        novo_nome = st.text_input("Nome descritivo do produto/serviço")
        novo_preco = st.number_input("Preço de tabela (R$)", min_value=0.0, value=0.0, step=50.0)
        
        if st.button("Salvar Inclusão no Banco de Dados"):
            if novo_nome:
                try:
                    cursor.execute("INSERT INTO produtos (nome, preco) VALUES (?, ?)", (novo_nome, novo_preco))
                    conn.commit()
                    st.success(f"💾 Sucesso! '{novo_nome}' foi incluído e salvo no banco de dados!")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("Erro: Esse produto já está cadastrado no banco de dados!")
            else:
                st.warning("Preencha o nome do produto.")
                
    elif sub_acao == "Visualizar e Alterar Preços":

        st.subheader("Produtos Armazenados no Sistema")
        cursor.execute("SELECT id, nome, preco FROM produtos")
        itens_armazenados = cursor.fetchall()
        
        if itens_armazenados:
            st.markdown("---")
            st.subheader("Alterar valor existente")
            opcoes_edicao = {f"{n} (R$ {p:,.2f})": (i, n) for i, n, p in itens_armazenados}
            selecao_edicao = st.selectbox("Selecione qual deseja atualizar:", list(opcoes_edicao.keys()))
            id_para_atualizar, nome_para_atualizar = opcoes_edicao[selecao_edicao]
            novo_valor_editado = st.number_input("Digite o novo valor comercial (R$)", min_value=0.0, step=10.0, key="edit_val")
            
            if st.button("Gravar Alteração de Preço"):
                cursor.execute("UPDATE produtos SET preco = ? WHERE id = ?", (novo_valor_editado, id_para_atualizar))
                conn.commit()
                st.success(f"🔄 Preço de '{nome_para_atualizar}' atualizado com sucesso!")
                st.rerun()
        else:
            st.info("Nenhum produto cadastrado.")

conn.close()
