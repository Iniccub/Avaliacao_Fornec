import streamlit as st
import pandas as pd
import os
from openpyxl import load_workbook
from streamlit_js_eval import streamlit_js_eval
from io import BytesIO
from fornecedores import fornecedores
from unidades import unidades
from perguntas_por_fornecedor import perguntas_por_fornecedor

st.set_page_config(
    page_title='Avaliação de Fornecedores - SUP',
    page_icon='CSA.png',
    layout='wide'
)

# Listas fixas
meses = ['31/01/2025', '28/02/2025', '31/03/2025', '30/04/2025', '31/05/2025', '30/06/2025', '31/07/2025', '31/08/2025',
         '30/09/2025', '31/10/2025', '30/11/2025', '31/12/2025']
opcoes = ['Atende Totalmente', 'Atende Parcialmente', 'Não Atende', 'Não se Aplica']


def carregar_fornecedores():
    if os.path.exists(CAMINHO_FORNECEDORES):
        try:
            from fornecedores import fornecedores
            return fornecedores
        except ImportError:
            return []
    return []

CAMINHO_FORNECEDORES = 'fornecedores.py'

def salvar_fornecedores(lista):
    with open(CAMINHO_FORNECEDORES, 'w', encoding='utf-8') as f:
        f.write('fornecedores = [\n')
        for item in lista:
            f.write(f"    '{item}',\n")
        f.write(']\n')

with st.sidebar:
    st.image("CSA.png", width=150)

# Aplicar estilo CSS para centralizar imagens na sidebar
st.markdown(
    """
    <style>
        [data-testid="stSidebar"] [data-testid="stImage"] {
            display: block;
            margin-left: 70px;
            margin-right: auto;
        }
    </style>
    """,
    unsafe_allow_html=True
)

st.sidebar.write('---')

# Sidebar, Caixas de seleção da unidade, período e fornecedor
unidade = st.sidebar.selectbox('Selecione a unidade', index=None, options=unidades, placeholder='Escolha a unidade')
periodo = st.sidebar.selectbox('Selecione o período avaliado', index=None, options=meses, placeholder='Defina o período de avaliação')
fornecedor = st.sidebar.selectbox('Selecione o fornecedor a ser avaliado', index=None, options=fornecedores, placeholder='Selecione o prestador/fornecedor')

st.sidebar.write('---')

with st.sidebar:
    # Cadastrar novo fornecedor
    novo_fornecedor = st.text_input('Cadastrar novo fornecedor: ')
    if st.button('Cadastrar fornecedor'):
        novo_fornecedor = novo_fornecedor.strip()
        if novo_fornecedor:
            if novo_fornecedor not in fornecedores:
                fornecedores.append(novo_fornecedor)
                salvar_fornecedores(fornecedores)
                st.toast(f'Fornecedor "{novo_fornecedor}" adicionado com sucesso!', icon='✅')
            else:
                st.warning('Fornecedor já existe na lista')
        else:
            st.warning('Por Favor, insira um nome válido')
        
# Tela para cadastrar nova pergunta
@st.dialog("Cadastrar Nova Pergunta", width="large")
def cadastrar_pergunta():
    st.subheader("Cadastro de Nova Pergunta")
    fornecedor = st.selectbox("Selecione o fornecedor", options=fornecedores)
    categoria = st.text_input("Categoria", placeholder="Ex: Documentação")
    nova_pergunta = st.text_area("Nova pergunta", placeholder="Digite a nova pergunta aqui")

    if st.button("Salvar"):
        if fornecedor and categoria and nova_pergunta:
            # Carregar perguntas existentes
            from perguntas_por_fornecedor import perguntas_por_fornecedor

            # Adicionar nova pergunta
            if fornecedor not in perguntas_por_fornecedor:
                perguntas_por_fornecedor[fornecedor] = {}
            if categoria not in perguntas_por_fornecedor[fornecedor]:
                perguntas_por_fornecedor[fornecedor][categoria] = []
            perguntas_por_fornecedor[fornecedor][categoria].append(nova_pergunta)

            # Salvar de volta no arquivo
            with open('perguntas_por_fornecedor.py', 'w', encoding='utf-8') as f:
                f.write('perguntas_por_fornecedor = {\n')
                for forn, cats in perguntas_por_fornecedor.items():
                    f.write(f"    '{forn}': {{\n")
                    for cat, perguntas in cats.items():
                        f.write(f"        '{cat}': [\n")
                        for pergunta in perguntas:
                            f.write(f"            '{pergunta}',\n")
                        f.write("        ],\n")
                    f.write("    },\n")
                f.write('}\n')
            
            st.success("Pergunta adicionada com sucesso!")
        else:
            st.warning("Por favor, preencha todos os campos.")

if st.sidebar.button("Cadastrar nova pergunta"):
    cadastrar_pergunta()

# Título
st.markdown(
    "<h1 style='text-align: left; font-family: Open Sauce; color: #104D73;'>"
    'ADFS - AVALIAÇÃO DE DESEMPENHO DE FORNECEDORES DE SERVIÇOS</h1>',
    unsafe_allow_html=True
)

st.write('---')

# Subtitulo
if fornecedor and unidade and periodo:
    st.subheader(f'Contratada/Fornecedor: {fornecedor}')
    st.write('Vigência: 02/01/2025 a 31/12/2025')
    st.write(f'Unidade: {unidade}')
    st.write(f'Período avaliado: {periodo}')
    st.write('---')

    # Determinação das abas
    tab1, tab2, tab3 = st.tabs(['Atividades Operacionais', 'Segurança', 'Qualidade'])

    respostas = []
    perguntas = []

    # Obter perguntas específicas do fornecedor
    perguntas_fornecedor = perguntas_por_fornecedor.get(fornecedor, {})

    with tab1:
        perguntas_tab1 = perguntas_fornecedor.get('Atividades Operacionais', [])
        for pergunta in perguntas_tab1:
            resposta = st.selectbox(pergunta, options=opcoes, index=None, placeholder='Selecione uma opção', key=pergunta)
            respostas.append(resposta)
            perguntas.append(pergunta)

    with tab2:
        perguntas_tab2 = perguntas_fornecedor.get('Segurança', [])
        for pergunta in perguntas_tab2:
            resposta = st.selectbox(pergunta, options=opcoes, index=None, placeholder='Selecione uma opção', key=pergunta)
            respostas.append(resposta)
            perguntas.append(pergunta)

    with tab3:
        perguntas_tab3 = perguntas_fornecedor.get('Qualidade', [])
        for pergunta in perguntas_tab3:
            resposta = st.selectbox(pergunta, options=opcoes, index=None, placeholder='Selecione uma opção', key=pergunta)
            respostas.append(resposta)
            perguntas.append(pergunta)

    st.sidebar.write('---')

    # Após coletar as perguntas e respostas de cada aba
    categorias = (
            ['Atividades Operacionais'] * len(perguntas_tab1) +
            ['Segurança'] * len(perguntas_tab2) +
            ['Qualidade'] * len(perguntas_tab3)
    )

    if st.sidebar.button('Salvar pesquisa'):
        # Verifica se todas as perguntas foram respondidas
        if None in respostas:
            st.warning('Por favor, responda todas as perguntas antes de salvar.')
        else:
            # Cria DataFrame com as respostas
            df_respostas = pd.DataFrame({
                'Unidade': unidade,
                'Período': periodo,
                'Fornecedor': fornecedor,
                'categorias': categorias,
                'Pergunta': perguntas,
                'Resposta': respostas
            })

            # Formata o nome do arquivo com base no fornecedor e período
            nome_fornecedor = fornecedor.replace(' ', '_')
            nome_periodo = periodo.replace('/', '-')
            nome_unidade = unidade
            nome_arquivo = f'{nome_fornecedor}_{nome_periodo}_{unidade}.xlsx'

            # Salva o DataFrame em um objeto BytesIO
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_respostas.to_excel(writer, index=False)
            output.seek(0)

            # Cria um botão de download no Streamlit
            st.download_button(
                label='Clique aqui para baixar o arquivo Excel com as respostas',
                data=output,
                file_name=nome_arquivo,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

            st.success('Respostas processadas com sucesso! Você pode baixar o arquivo acima')
    else:
        st.warning('Por favor, selecione a unidade, o período e o fornecedor para iniciar a avaliação.')

    if st.sidebar.button("Preencher nova pesquisa"):
        streamlit_js_eval(js_expressions='parent.window.location.reload()')

# Rodapé com copyright
st.sidebar.markdown("""
    <style>
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #f0f0f0;
        color: #333;
        text-align: center;
        padding: 10px;
        font-size: 14px;
    }
    </style>
    <div class="footer">
        © 2025 FP&A e Orçamento - Rede Lius. Todos os direitos reservados.
    </div>
    """, unsafe_allow_html=True)
