import streamlit as st
import streamlit.components.v1 as components # NOVO: Importação para injetar o bloqueio do teclado
from pathlib import Path
from datetime import datetime
import io
import mimetypes
import re

import pandas as pd 
import gspread
#from google.oauth2.service_account import Credentials
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

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
    "despesas": ["id", "data_registro_sistema", "data_despesa", "doc_comprovante", "contrato", "cnpj_cpf", "agencia_pagadora", "conta_pagadora", "natureza_despesa", "valor", "nota_explicativa", "comprovante_nome_arquivo", "comprovante_drive_id", "comprovante_drive_link"],
    "usuarios": ["id", "data_registro_sistema", "nome_completo", "data_nascimento", "cpf", "estado", "cidade", "rua", "numero", "bairro", "cep"],
    "bancos": ["id", "data_registro_sistema", "nome_banco", "numero_agencia", "numero_conta_corrente", "nome_conta"],
    "credores": ["id", "data_registro_sistema", "cnpj_cpf", "nome_credor", "estado", "cidade", "rua", "numero", "cep", "agencia_credor", "conta_credor"],
    "colaboradores": ["id", "data_registro_sistema", "cpf", "data_nascimento", "nome_completo", "estado", "cidade", "rua", "numero", "cep", "carteira_trabalho"]
}

#@st.cache_resource
#def criar_credenciais_google():
#    """Cria uma credencial única para Google Sheets e Google Drive."""
#    dados_credenciais = st.secrets["gcp_service_account"]
#    escopos = [
#        "https://www.googleapis.com/auth/spreadsheets",
#        "https://www.googleapis.com/auth/drive",
#    ]
#    return Credentials.from_service_account_info(dados_credenciais, scopes=escopos)

@st.cache_resource(ttl="1h") # No dia 28/05/2026 o Aplicativo deixou de funcionar. Descobri que era o token que precisava ser renovado a cada 1 hora.
def criar_credenciais_google():
    """Cria credencial usando o Token OAuth do usuário (Conta de 5TB)."""
    # Puxamos o dicionário do secrets e transformamos a chave
    dados_token = dict(st.secrets["google_oauth_token"])
    
    escopos = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    
    # Gera as credenciais dizendo: "Eu sou o dono da conta"
    return Credentials.from_authorized_user_info(dados_token, scopes=escopos)


@st.cache_resource(ttl="1h") # No dia 28/05/2026 o Aplicativo deixou de funcionar. Descobri que era o token que precisava ser renovado a cada 1 hora.
def conectar_planilha():
    credenciais = criar_credenciais_google()
    cliente = gspread.authorize(credenciais)
    return cliente.open("base_dados_partido") # O nome exato da sua planilha no Google Drive


@st.cache_resource(ttl="1h") # No dia 28/05/2026 o Aplicativo deixou de funcionar. Descobri que era o token que precisava ser renovado a cada 1 hora.
def conectar_drive():
    """Conecta na API do Google Drive usando a mesma conta de serviço."""
    credenciais = criar_credenciais_google()
    return build("drive", "v3", credentials=credenciais)


def obter_id_pasta_comprovantes():
    """Lê dos secrets o ID da pasta do Google Drive onde ficarão os comprovantes."""
    try:
        return st.secrets["google_drive"]["comprovantes_folder_id"]
    except Exception:
        return st.secrets["GOOGLE_DRIVE_COMPROVANTES_FOLDER_ID"]


def limpar_texto_arquivo(texto):
    """Remove caracteres que podem atrapalhar o nome do arquivo no Drive."""
    texto = str(texto or "sem_numero").strip().lower()
    texto = re.sub(r"[^a-z0-9_-]+", "_", texto)
    return texto.strip("_") or "sem_numero"


def salvar_comprovante_drive(arquivo, id_despesa, data_registro, comprovante):
    """
    Salva o comprovante no Google Drive e retorna os dados necessários
    para vincular o arquivo à linha da despesa no Google Sheets.
    """
    drive = conectar_drive()
    pasta_id = obter_id_pasta_comprovantes()

    extensao = Path(arquivo.name).suffix.lower() or ".pdf"
    data_nome = data_registro.strftime("%Y%m%d_%H%M%S")
    nf_nome = limpar_texto_arquivo(comprovante)
    nome_arquivo = f"despesa_{int(id_despesa):06d}_{data_nome}_nf_{nf_nome}{extensao}"

    mime_type = arquivo.type or mimetypes.guess_type(nome_arquivo)[0] or "application/octet-stream"
    media = MediaIoBaseUpload(
        io.BytesIO(arquivo.getvalue()),
        mimetype=mime_type,
        resumable=False,
    )

    metadados = {
        "name": nome_arquivo,
        "parents": [pasta_id],
    }

    arquivo_drive = drive.files().create(
        body=metadados,
        media_body=media,
        fields="id, name, webViewLink",
        supportsAllDrives=True,
    ).execute()

    return {
        "comprovante_nome_arquivo": arquivo_drive.get("name", nome_arquivo),
        "comprovante_drive_id": arquivo_drive.get("id", ""),
        "comprovante_drive_link": arquivo_drive.get("webViewLink", ""),
    }


def garantir_cabecalho_planilha(aba, nome_aba):
    """Garante que a aba possui todas as colunas esperadas no cabeçalho."""
    colunas_esperadas = TABELAS[nome_aba]
    cabecalho_atual = aba.row_values(1)

    if not cabecalho_atual:
        aba.append_row(colunas_esperadas)
        return

    colunas_faltantes = [col for col in colunas_esperadas if col not in cabecalho_atual]
    if colunas_faltantes:
        novo_cabecalho = cabecalho_atual + colunas_faltantes
        aba.update("A1", [novo_cabecalho])


def proximo_id_nuvem(aba):
    valores = aba.col_values(1) # Pega todos os valores da coluna A (id)
    if len(valores) <= 1:
        return 1
    try:
        return int(valores[-1]) + 1
    except:
        return len(valores)

# Mantivemos o nome 'salvar_registro_excel' para você não precisar mudar os formulários
def salvar_registro_excel(nome_aba, dados, id_predefinido=None, data_registro_sistema=None):
    planilha = conectar_planilha()
    aba = planilha.worksheet(nome_aba)
    garantir_cabecalho_planilha(aba, nome_aba)

    novo_id = id_predefinido if id_predefinido is not None else proximo_id_nuvem(aba)
    data_sistema = data_registro_sistema or datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    dados_completos = {
        "id": novo_id,
        "data_registro_sistema": data_sistema,
        **dados
    }
    
    linha = [str(dados_completos.get(coluna, "")) for coluna in TABELAS[nome_aba]]
    aba.append_row(linha)
    return dados_completos

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
            doc_comprovante = st.text_input("Nota fiscal/comprovante nº:")
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

        st.divider()
        st.subheader("Comprovante da Despesa")
        comprovante = st.file_uploader(
            "Anexar comprovante / Comprovante",
            type=["pdf", "png", "jpg", "jpeg"],
            help="Envie preferencialmente em PDF. Também são aceitas imagens PNG/JPG.",
        )
        
        # O BOTÃO DE ENVIO DO FORMULÁRIO FICA POR ÚLTIMO, APÓS TODAS AS SELEÇÕES E CAMPOS DE TEXTO
        submit = st.form_submit_button("Salvar Despesa", type="primary")
        
        # 4. SALVAMENTO UNIFICADO
        if submit:
            if comprovante is None:
                st.warning("Anexe o comprovante da despesa antes de salvar o registro.")
                st.stop()

            # Extrai apenas o CNPJ
            cnpj_puro = credor_selecionado.split(" - ")[0]
            
            # Extrai apenas a Agência e a Conta da opção selecionada
            partes_conta = conta_selecionada.split(" | ")
            agencia_deb = partes_conta[0].replace("Agência: ", "")
            conta_deb = partes_conta[1].replace("Conta: ", "")

            # Define ID e data antes de salvar, para que a linha da planilha
            # e o arquivo no Google Drive tenham o mesmo identificador.
            planilha = conectar_planilha()
            aba_despesas = planilha.worksheet("despesas")
            garantir_cabecalho_planilha(aba_despesas, "despesas")
            novo_id = proximo_id_nuvem(aba_despesas)
            data_registro = datetime.now()
            data_registro_texto = data_registro.strftime("%d/%m/%Y %H:%M:%S")

            try:
                dados_comprovante = salvar_comprovante_drive(
                    arquivo=comprovante,
                    id_despesa=novo_id,
                    data_registro=data_registro,
                    comprovante=comprovante,
                )
            except Exception as erro:
                st.error(f"Não foi possível salvar o comprovante no Google Drive: {erro}")
                st.stop()
            
            # Reúne TODOS os dados em um único dicionário para a tabela
            dados = {
                "data_despesa": data_desp.strftime("%d/%m/%Y"),
                "doc_comprovante": doc_comprovante,
                "contrato": contrato,
                "cnpj_cpf": str(cnpj_puro),
                #"cnpj_cpf": cnpj_puro,
                "agencia_pagadora": agencia_deb,
                "conta_pagadora": conta_deb,
                "natureza_despesa": natureza_desp,
                "valor": f"R$ {valor_desp:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                "nota_explicativa": nota_explicativa,
                **dados_comprovante,
            }
            
            # Salva na aba 'despesas'
            registro_salvo = salvar_registro_excel(
                "despesas",
                dados,
                id_predefinido=novo_id,
                data_registro_sistema=data_registro_texto,
            )
            st.success(f"Despesa ID {registro_salvo['id']} registrada com sucesso!")
            st.markdown(f"[Abrir comprovante no Google Drive]({registro_salvo['comprovante_drive_link']})")


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

def tela_relatorio_despesa():
    st.header("Relatório da Despesa")
    
    # 1. Carrega os dados usando a função que já usa o CACHE (muito mais rápido!)
    df_despesa = carregar_dataframe("despesas")

    # 2. Verifica se o DataFrame não está vazio antes de tentar tratar
    if not df_despesa.empty: # Se o DataFrame não estiver vazio faça:

        # Garante que a coluna valor é texto antes de limpar
        df_despesa['valor'] = df_despesa['valor'].astype(str)

        # Tratamento dos dados contidos na Variável "Valor"
        df_despesa['valor'] = df_despesa['valor'].str.replace('R$', "", regex=False).str.strip()
        df_despesa['valor'] = df_despesa['valor'].str.replace('.', "", regex=False)
        df_despesa['valor'] = df_despesa['valor'].str.replace(',', '.', regex=False)
        df_despesa['valor'] = pd.to_numeric(df_despesa['valor'])

        # Calcular o somatório
        soma_total = df_despesa['valor'].sum()

        # Criar a linha de total
        linha_total = pd.DataFrame([{
            'id': '', 
            'data_registro_sistema': 'TOTAL DA DESPESA',
            'data_despesa': "",
            'doc_comprovante': "",
            'contrato': "",
            'cnpj_cpf': "",
            'agencia_pagadora': "",
            'conta_pagadora': "",
            'natureza_despesa': "",
            'valor': soma_total,
            'nota_explicativa': "",
            'comprovante_nome_arquivo': "",
            'comprovante_drive_id': "",
            'comprovante_drive_link': ""
        }])

        # Juntar a linha do total no topo do DataFrame
        df_despesa = pd.concat([linha_total, df_despesa], ignore_index=True)

        # Exibe a tabela na tela (sem precisar de formulário)
        if "comprovante_drive_link" in df_despesa.columns:
            st.dataframe(
                df_despesa,
                use_container_width=True,
                column_config={
                    "comprovante_drive_link": st.column_config.LinkColumn("Comprovante")
                },
            )
        else:
            st.dataframe(df_despesa, use_container_width=True)

    else:
        st.warning("Nenhuma despesa registrada até o momento.")
    

# =============================================================
# SISTEMA DE LOGIN
# =============================================================
def checar_login():
    # Se a variável 'logado' não existir na sessão, cria como Falso
    if "logado" not in st.session_state:
        st.session_state["logado"] = False

    # Se não estiver logado, mostra a tela de login
    if not st.session_state["logado"]:
        st.title("Acesso Restrito")
        st.markdown("Por favor, faça o login para acessar o sistema de Gestão Partidária.")
        
        # Formulário de Login
        usuario = st.text_input("Usuário")
        senha = st.text_input("Senha", type="password") # type="password" esconde a senha
        
        if st.button("Entrar", type="primary"):
            # Puxa os usuários cadastrados no cofre do Streamlit
            usuarios_cadastrados = st.secrets["usuarios"]
            
            # Verifica se o usuário existe e se a senha está correta
            if usuario in usuarios_cadastrados and usuarios_cadastrados[usuario] == senha:
                st.session_state["logado"] = True
                st.session_state["usuario_atual"] = usuario
                st.rerun() # Recarrega a página agora logado
            else:
                st.error("Usuário ou senha incorretos.")
        
        # Retorna Falso para impedir que o resto do app carregue
        return False
    
    # Se já estiver logado, retorna Verdadeiro
    return True


# =============================================================
# O MOTOR DO PROGRAMA (MENU LATERAL)
# =============================================================
def main():
    bloquear_enter() # Chama a função que trava o Enter nos formulários do site
        
    st.sidebar.title("PSD")
    st.sidebar.markdown("Contábil - (Receita x Despesa)")
    
    menu = st.sidebar.selectbox(
        "Selecione uma opção:",
        ["Registrar Receita", 
         "Registrar Despesa", 
         #"Cadastrar Usuário", # Ao inserir o "#" desabilitei a tela "Cadastrar Usuário"
         "Cadastrar Banco", 
         "Cadastrar Credor", 
         #"Cadastro de Colaboradores", # Ao inserir o "#" desabilitei a tela "Cadastro de Colaboradores"
         "Relatório da Despesa"]
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
    elif menu == "Relatório da Despesa":
        tela_relatorio_despesa()

# Comando para iniciar o Streamlit
if __name__ == "__main__":
    if checar_login():
        main()



