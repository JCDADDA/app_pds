import streamlit as st
import streamlit.components.v1 as components # NOVO: Importação para injetar o bloqueio do teclado
from pathlib import Path
from datetime import datetime
import pandas as pd 
import gspread
from google.oauth2.service_account import Credentials

# =============================================================
# CARREGANDO DADOS REGISTRADOS NO BANCO DE DADOS
# =============================================================


# =============================================================
# CARREGANDO DADOS DA PRESTAÇÃO DE CONTAS DE 2024
# =============================================================
extrato_bancario_pca = "extrato_partido_.xlsx"
extrato_bancario_pca = pd.read_excel(extrato_bancario_pca)

# =============================================================
# CONFIGURAÇÃO INICIAL DA PÁGINA WEB
# =============================================================
st.set_page_config(page_title="Gestão Partidária", page_icon="🏛️", layout="centered")

# =============================================================
# FUNÇÃO PARA BLOQUEAR O "ENTER" NOS FORMULÁRIOS
# =============================================================
def bloquear_enter():
    """Injeta um script no navegador para impedir o envio de forms com a tecla Enter"""
    components.html(
        """
        <script>
        const doc = window.parent.document;
        doc.addEventListener('keydown', function(e) {
            // Se a tecla for Enter e o usuário estiver digitando em um campo (INPUT)
            if (e.key === 'Enter' && e.target.nodeName === 'INPUT') {
                e.preventDefault(); // Cancela o envio
            }
        });
        </script>
        """,
        height=0, width=0 # Mantém o script invisível na tela
    )

# =============================================================
# CRIANDO ESTRUTURA DO BANCO EM EXCEL (Lógica Intacta) - NOVA LÓGICA DE BANCO DE DADOS: GOOGLE SHEETS
# =============================================================
TABELAS = {
    "receitas": ["id", "data_registro_sistema", "data_receita", "agencia", "conta_corrente", "origem_receita", "agencia_origem", "conta_origem", "natureza_receita", "valor", "nota_explicativa"],
    "despesas": ["id", "data_registro_sistema", "data_despesa", "nota_fiscal", "contrato", "cnpj_cpf", "agencia_pagadora", "conta_pagadora", "natureza_despesa", "valor", "nota_explicativa"],
    "usuarios": ["id", "data_registro_sistema", "nome_completo", "data_nascimento", "cpf", "estado", "cidade", "rua", "numero", "bairro", "cep"],
    "bancos": ["id", "data_registro_sistema", "nome_banco", "numero_agencia", "numero_conta_corrente", "nome_conta"],
    "credores": ["id", "data_registro_sistema", "cnpj_cpf", "nome_credor", "estado", "cidade", "rua", "numero", "cep", "agencia_credor", "conta_credor"],
    "colaboradores": ["id", "data_registro_sistema", "cpf", "data_nascimento", "nome_completo", "estado", "cidade", "rua", "numero", "cep", "carteira_trabalho"]
}

@st.cache_resource # Isso evita que o app reconecte com o Google a cada clique
def conectar_planilha():
    dados_credenciais = st.secrets["gcp_service_account"]
    escopos = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    credenciais = Credentials.from_service_account_info(dados_credenciais, scopes=escopos)
    cliente = gspread.authorize(credenciais)
    return cliente.open("base_dados_partido") # O nome exato da sua planilha no Google Drive

def proximo_id_nuvem(aba):
    valores = aba.col_values(1) # Pega todos os valores da coluna A (id)
    if len(valores) <= 1:
        return 1
    try:
        return int(valores[-1]) + 1
    except:
        return len(valores)

# Mantivemos o nome 'salvar_registro_excel' para você não precisar mudar os formulários
def salvar_registro_excel(nome_aba, dados):
    planilha = conectar_planilha()
    aba = planilha.worksheet(nome_aba)
    novo_id = proximo_id_nuvem(aba)

    dados_completos = {
        "id": novo_id,
        "data_registro_sistema": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        **dados
    }
    
    linha = [str(dados_completos.get(coluna, "")) for coluna in TABELAS[nome_aba]]
    aba.append_row(linha)

def carregar_contas_bancarias():
    planilha = conectar_planilha()
    aba = planilha.worksheet("bancos")
    registros = aba.get_all_records()
    return [r for r in registros if r.get("numero_agencia")]

def carregar_credores():
    planilha = conectar_planilha()
    aba = planilha.worksheet("credores")
    registros = aba.get_all_records()
    return [r for r in registros if r.get("cnpj_cpf")]

def carregar_dataframe(nome_aba):
    """Função extra para exibir as tabelas na tela do Streamlit"""
    planilha = conectar_planilha()
    aba = planilha.worksheet(nome_aba)
    dados = aba.get_all_records()
    if not dados:
        return pd.DataFrame(columns=TABELAS[nome_aba])
    return pd.DataFrame(dados)


# =============================================================
# INTERFACES WEB (AS TELAS DO STREAMLIT)
# =============================================================

def tela_receita():
    st.header("Registrar Receita")
    
    contas_cadastradas = carregar_contas_bancarias()
    if not contas_cadastradas:
        st.warning("Não existe conta bancária cadastrada. Cadastre uma conta no menu 'Cadastrar Banco' primeiro.")
        return

    opcoes_contas = [f"Agência: {c['numero_agencia']} | Conta: {c['numero_conta_corrente']} | {c['nome_banco']} - {c['nome_conta']}" for c in contas_cadastradas]
    
    conta_selecionada = st.selectbox("Selecione a Conta Creditada:", opcoes_contas)
    
    with st.form("form_receita", clear_on_submit=True):
        col1, col2 = st.columns(2) 
        
        with col1:
            st.subheader("Dados da Receita")
            data_rec = st.date_input("Data da Receita", format="DD/MM/YYYY")
            natureza_rec = st.text_input("Tipo (natureza) da Receita")
            valor_rec = st.number_input("Valor (R$)", min_value=0.0, format="%.2f")
            
        with col2:
            st.subheader("Fonte Pagadora")
            origem_rec = st.text_input("Origem da Receita")
            agencia_origem = st.text_input("Agência de Origem")
            conta_origem = st.text_input("Conta Corrente de Origem")
            
        nota_explicativa = st.text_area("Informações Adicionais", height=100)
        
        submit = st.form_submit_button("Salvar Receita", type="primary")
        
        if submit:
            partes_conta = conta_selecionada.split(" | ")
            agencia = partes_conta[0].replace("Agência: ", "")
            conta = partes_conta[1].replace("Conta: ", "")
            
            dados = {
                "data_receita": data_rec.strftime("%d/%m/%Y"),
                "agencia": agencia,
                "conta_corrente": conta,
                "origem_receita": origem_rec,
                "agencia_origem": agencia_origem,
                "conta_origem": conta_origem,
                "natureza_receita": natureza_rec,
                "valor": f"R$ {valor_rec:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                "nota_explicativa": nota_explicativa
            }
            salvar_registro_excel("receitas", dados)
            st.success("Receita registrada com sucesso!")


def tela_despesa():
    st.header("Registrar Despesa")
    
    # 1. CARREGAMENTO E VALIDAÇÃO DE DADOS BÁSICOS
    credores_cadastrados = carregar_credores()
    if not credores_cadastrados: # Verifica se há credores cadastrados, caso contrário, exibe um aviso de usuário não cadastrado.
        st.warning("Não existe credor cadastrado no sistema. Cadastre um no menu lateral.") 
        return

    contas_cadastradas = carregar_contas_bancarias()
    if not contas_cadastradas:
        st.warning("Não existe conta bancária cadastrada. Cadastre uma conta no menu 'Cadastrar Banco' primeiro.")
        return

    # 2. PREPARAÇÃO DAS LISTAS DE OPÇÕES
    opcoes_credores = [f"{c['cnpj_cpf']} - {c['nome_credor']}" for c in credores_cadastrados]
    opcoes_contas = [f"Agência: {c['numero_agencia']} | Conta: {c['numero_conta_corrente']} | {c['nome_banco']} - {c['nome_conta']}" for c in contas_cadastradas]

    # Seleção do Credor
    # Note que o selectbox de seleção de credor é construído ANTES do formulário.
    credor_selecionado = st.selectbox("Selecione o Credor (CNPJ/CPF):", opcoes_credores)
    
    # 3. CONSTRUÇÃO DO FORMULÁRIO
    with st.form("form_despesa", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            data_desp = st.date_input("Data da Despesa", format="DD/MM/YYYY")
            nota_fiscal = st.text_input("Nota fiscal nº")
            contrato = st.text_input("Contrato nº")
            
        with col2:
            natureza_desp = st.text_input("Tipo (natureza) da Despesa")
            valor_desp = st.number_input("Valor (R$)", min_value=0.0, format="%.2f")
            
        st.divider() # Adiciona uma linha horizontal para separar os assuntos visualmente
        st.subheader("Pagamento")
        
        # A CAIXA DE SELEÇÃO DO BANCO ENTRA AQUI
        conta_selecionada = st.selectbox("Selecione a Conta Debitada (Saída do Dinheiro):", opcoes_contas)

        # Caixa de texto para Justificativa, que pode ser usada para detalhar a despesa, justificar o motivo, ou qualquer informação adicional relevante.
        nota_explicativa = st.text_area("Justificativa")
        
        # O BOTÃO DE ENVIO DO FORMULÁRIO FICA POR ÚLTIMO, APÓS TODAS AS SELEÇÕES E CAMPOS DE TEXTO
        submit = st.form_submit_button("Salvar Despesa", type="primary")
        
        # 4. SALVAMENTO UNIFICADO
        if submit:
            # Extrai apenas o CNPJ
            cnpj_puro = credor_selecionado.split(" - ")[0]
            
            # Extrai apenas a Agência e a Conta da opção selecionada
            partes_conta = conta_selecionada.split(" | ")
            agencia_deb = partes_conta[0].replace("Agência: ", "")
            conta_deb = partes_conta[1].replace("Conta: ", "")
            
            # Reúne TODOS os dados em um único dicionário para a tabela
            dados = {
                "data_despesa": data_desp.strftime("%d/%m/%Y"),
                "nota_fiscal": nota_fiscal,
                "contrato": contrato,
                "cnpj_cpf": cnpj_puro,
                "agencia_pagadora": agencia_deb,
                "conta_pagadora": conta_deb,
                "natureza_despesa": natureza_desp,
                "valor": f"R$ {valor_desp:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                "nota_explicativa": nota_explicativa
            }
            
            # Salva na aba 'despesas'
            salvar_registro_excel("despesas", dados)
            st.success("Despesa registrada com sucesso!")


def tela_cadastro_usuario():
    st.header("Cadastrar Usuário")
    with st.form("form_usuario", clear_on_submit=True):
        nome = st.text_input("Nome Completo")
        col1, col2 = st.columns(2)
        with col1:
            data_nasc = st.date_input("Data de Nascimento", format="DD/MM/YYYY")
            cpf = st.text_input("CPF (Apenas números)", max_chars=11)
            cep = st.text_input("CEP (Apenas números)", max_chars=8)
        with col2:
            estado = st.text_input("Estado (UF)", max_chars=2)
            cidade = st.text_input("Cidade")
            bairro = st.text_input("Bairro")
            
        col3, col4 = st.columns([3, 1]) 
        with col3:
            rua = st.text_input("Rua")
        with col4:
            numero = st.text_input("Número")
            
        if st.form_submit_button("Cadastrar Usuário", type="primary"):
            dados = {
                "nome_completo": nome, "data_nascimento": data_nasc.strftime("%d/%m/%Y"),
                "cpf": cpf, "estado": estado, "cidade": cidade,
                "rua": rua, "numero": numero, "bairro": bairro, "cep": cep
            }
            salvar_registro_excel("usuarios", dados)
            st.success("Usuário cadastrado com sucesso!")
    
    st.header("Usuários Cadastrados")
    st.dataframe(carregar_dataframe("usuarios"))


def tela_cadastro_banco():
    st.header("Cadastrar Dados Bancários")
    with st.form("form_banco", clear_on_submit=True):
        nome_banco = st.text_input("Nome do Banco")
        agencia = st.text_input("Número da Agência")
        conta = st.text_input("Nº Conta Corrente")
        nome_conta = st.text_input("Nome da Conta (Ex: Doações, Fundo)")
        
        if st.form_submit_button("Cadastrar Banco", type="primary"):
            salvar_registro_excel("bancos", {"nome_banco": nome_banco, "numero_agencia": agencia, "numero_conta_corrente": conta, "nome_conta": nome_conta})
            st.success("Banco cadastrado com sucesso!")
    
    st.header("Bancos Cadastrados")
    st.dataframe(carregar_dataframe("bancos"))


def tela_cadastro_credor():
    st.header("Cadastrar Credor")
    with st.form("form_credor", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            cnpj = st.text_input("CNPJ/CPF (Apenas números)")
            nome = st.text_input("Nome do Credor")
            agencia = st.text_input("Agência (Credor)")
            conta = st.text_input("Conta Corrente (Credor)")
        with col2:
            estado = st.text_input("Estado (UF)")
            cidade = st.text_input("Cidade")
            rua = st.text_input("Rua")
            numero = st.text_input("Número")
            cep = st.text_input("CEP")
            
        if st.form_submit_button("Cadastrar Credor", type="primary"):
            dados = {"cnpj_cpf": cnpj, "nome_credor": nome, "estado": estado, "cidade": cidade, "rua": rua, "numero": numero, "cep": cep, "agencia_credor": agencia, "conta_credor": conta}
            salvar_registro_excel("credores", dados)
            st.success("Credor cadastrado com sucesso!")

    st.header("Credores Cadastrados")
    st.dataframe(carregar_dataframe("credores"))

def tela_cadastro_colaborador():
    st.header("Cadastrar Colaborador")
    with st.form("form_colab", clear_on_submit=True):
        nome = st.text_input("Nome Completo")
        col1, col2 = st.columns(2)
        with col1:
            cpf = st.text_input("CPF")
            data_nasc = st.date_input("Data de Nascimento", format="DD/MM/YYYY")
            ctb = st.text_input("Carteira de Trabalho")
            cep = st.text_input("CEP")
        with col2:
            estado = st.text_input("Estado (UF)")
            cidade = st.text_input("Cidade")
            rua = st.text_input("Rua")
            numero = st.text_input("Número")
            
        if st.form_submit_button("Cadastrar Colaborador", type="primary"):
            dados = {"cpf": cpf, "data_nascimento": data_nasc.strftime("%d/%m/%Y"), "nome_completo": nome, "estado": estado, "cidade": cidade, "rua": rua, "numero": numero, "cep": cep, "carteira_trabalho": ctb}
            salvar_registro_excel("colaboradores", dados)
            st.success("Colaborador cadastrado com sucesso!")


# =============================================================
# O MOTOR DO PROGRAMA (MENU LATERAL)
# =============================================================
def main():
    bloquear_enter() # Chama a função que trava o Enter nos formulários do site
        
    st.sidebar.title("PSD")
    st.sidebar.markdown("Contábil - (Receita x Despesa)")
    
    menu = st.sidebar.radio(
        "Selecione uma opção:",
        ["Registrar Receita", "Registrar Despesa", "Cadastrar Usuário", "Cadastrar Banco", "Cadastrar Credor", "Cadastro de Colaboradores"]
    )
    
    if menu == "Registrar Receita":
        tela_receita()
    elif menu == "Registrar Despesa":
        tela_despesa()
    elif menu == "Cadastrar Usuário":
        tela_cadastro_usuario()
    elif menu == "Cadastrar Banco":
        tela_cadastro_banco()
    elif menu == "Cadastrar Credor":
        tela_cadastro_credor()
    elif menu == "Cadastro de Colaboradores":
        tela_cadastro_colaborador()

# Comando para iniciar o Streamlit
if __name__ == "__main__":
    main()
