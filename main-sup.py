import streamlit as st
import pandas as pd
import os
from openpyxl import load_workbook
from streamlit_js_eval import streamlit_js_eval
from io import BytesIO

st.set_page_config(
    page_title='Avaliação de Fornecedores - SUP',
    page_icon='CSA.png',
    layout='wide'
)

# Listas fixas
unidades = ['CSA-BH', 'CSA-CT', 'CSA-NL', 'CSA-GZ', 'CSA-DV', 'EPSA', 'ESA', 'AIACOM', 'ILALI', 'ADEODATO', 'SIC SEDE']
meses = ['31/01/2025', '28/02/2025', '31/03/2025', '30/04/2025', '31/05/2025', '30/06/2025', '31/07/2025', '31/08/2025',
         '30/09/2025', '31/10/2025', '30/11/2025', '31/12/2025']
fornecedores = ['CANTINA FREITAS',
                'EXPRESSA TURISMO LTDA',
                'ACREDITE EXCURSÕES E EXPOSIÇÕES INTINERANTE LTDA',
                'LEAL VIAGENS E TURISMO',
                'MINASCOPY NACIONAL EIRELI',
                'OTIMIZA VIGILÂNCIA E SEG. PATRIMONIAL',
                'PETRUS LOCACAO E SERVICOS LTDA',
                'REAL VANS LOCAÇÕES',
                'AC TRANSPORTES E SERVIÇOS LTDA - ACTUR',
                'TRANSCELO TRANSPORTES LTDA',
                'AC Transportes e Serviços LTDA',
                'GULP SÃO TOMAS',
                'NUTRIMIX - EXCELÊNCIA EM ALIMENTAÇÃO',
                'SALADA & TAL ( PAOLA OLIVEIRA COSTA )',
                'ELEVADORES ATLAS SCHINDLER LTDA',
                'TK ELEVADORES BRASIL LTDA',
                'ELEVAÇO LTDA',
                'JD CONSERVAÇÃO E SERVIÇOS',
                'QA - IT ANSWER - CONSULTORIA - N1',
                'QA - IT ANSWER - CONSULTORIA - N2',
                'MODERNA TURISMO LTDA',
                'XINGU ELEVADORES',
                'PHP SERVICE EIRELI',
                'CAMPOS DE MINAS SERV. ORG. PROG.TURÍSTICOS',
                'ACCESS GESTÃO DE DOCUMENTOS LTDA',
                'BOCAINA CIENCIAS NATURAIS & EDUCACAO AMBIENTAL',
                'NOVA FORMA VIAGENS E TURISMO',
                'CONSERVADORA CIDADE LC',
                'CONSERVADORA CIDADE PC',
                'OTIS ELEVADORES'
                ]
opcoes = ['Atende Totalmente', 'Atende Parcialmente', 'Não Atende', 'Não se Aplica']

# Inserir a imagem na sidebar
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

# Dicionário de perguntas por fornecedor
perguntas_por_fornecedor = {
    'CANTINA FREITAS': {
        'Documentação': [
            '1 - Os documentos obrigatórios para análise e faturamento foram entregues dentro do prazo acordado em contrato?',
            '2 - A contratada apresentou todas as documentações exigidas, conforme contrato com os devidos recolhimentos e pagamentos?',
            '3 - A Nota Fiscal foi emitida com dados corretos?',
            '4 - CND-FGTS e CRT, relativos à Regularidade Fiscal e Trabalhista, estão atualizados?',
            '5 - Apresentam à SIC cópia autenticada do ALVARÁ DE FUNCIONAMENTO, expedido pelos órgãos competentes, por meio do qual a CONTRATADA ficará autorizada a realizar suas atividades comerciais'
        ]
    },
    'EXPRESSA TURISMO LTDA': {
        'Documentação': [
            '1 - Os documentos obrigatórios para análise e faturamento foram entregues dentro do prazo acordado em contrato?',
            '2 - A contratada apresentou todas as documentações exigidas, conforme contrato com os devidos recolhimentos e pagamentos?',
            '3 - A Nota Fiscal foi emitida com dados corretos?',
            '4 - A contratada possui autorizações atualizadas (licenças específicas) que a habilitem para a prestação de serviços de transporte, expedidas pelos órgãos competentes (ANTT, DER, DETRAN, BHTRANS e/ou quaisquer outros);'
            '5 - A Contratada possui seguro de Acidentes pessoais, extensivo aos passageiros, e outros exigidos por Lei?'
        ]
    },
'LEAL VIAGENS E TURISMO': {
        'Documentação': [
            '1 - Os documentos obrigatórios para análise e faturamento foram entregues dentro do prazo acordado em contrato?',
            '2 - A contratada apresentou todas as documentações exigidas, conforme contrato com os devidos recolhimentos e pagamentos?',
            '3 - A Nota Fiscal foi emitida com dados corretos?',
            '4 - A contratada possui autorizações atualizadas (licenças específicas) que a habilitem para a prestação de serviços de transporte, expedidas pelos órgãos competentes (ANTT, DER, DETRAN, BHTRANS e/ou quaisquer outros);'
            '5 - A Contratada possui seguro de Acidentes pessoais, extensivo aos passageiros, e outros exigidos por Lei?'
        ]
    },
'ACREDITE EXCURSÕES E EXPOSIÇÕES INTINERANTE LTDA': {
        'Documentação': [
            '1 - Os documentos obrigatórios para análise e faturamento foram entregues dentro do prazo acordado em contrato?',
            '2 - A contratada apresentou todas as documentações exigidas, conforme contrato com os devidos recolhimentos e pagamentos?',
            '3 - A Nota Fiscal foi emitida com dados corretos?',
            '4 - A contratada possui autorizações atualizadas (licenças específicas) que a habilitem para a prestação de serviços de transporte, expedidas pelos órgãos competentes (ANTT, DER, DETRAN, BHTRANS e/ou quaisquer outros);'
            '5 - A Contratada possui seguro de Acidentes pessoais, extensivo aos passageiros, e outros exigidos por Lei?'
        ]
    },
'REAL VANS LOCAÇÕES': {
        'Documentação': [
            '1 - Os documentos obrigatórios para análise e faturamento foram entregues dentro do prazo acordado em contrato?',
            '2 - A contratada apresentou todas as documentações exigidas, conforme contrato com os devidos recolhimentos e pagamentos?',
            '3 - A Nota Fiscal foi emitida com dados corretos?',
            '4 - A contratada possui autorizações atualizadas (licenças específicas) que a habilitem para a prestação de serviços de transporte, expedidas pelos órgãos competentes (ANTT, DER, DETRAN, BHTRANS e/ou quaisquer outros);'
            '5 - A Contratada possui seguro de Acidentes pessoais, extensivo aos passageiros, e outros exigidos por Lei?'
        ]
    },
'AC TRANSPORTES E SERVIÇOS LTDA - ACTUR': {
        'Documentação': [
            '1 - Os documentos obrigatórios para análise e faturamento foram entregues dentro do prazo acordado em contrato?',
            '2 - A contratada apresentou todas as documentações exigidas, conforme contrato com os devidos recolhimentos e pagamentos?',
            '3 - A Nota Fiscal foi emitida com dados corretos?',
            '4 - A contratada possui autorizações atualizadas (licenças específicas) que a habilitem para a prestação de serviços de transporte, expedidas pelos órgãos competentes (ANTT, DER, DETRAN, BHTRANS e/ou quaisquer outros);'
            '5 - A Contratada possui seguro de Acidentes pessoais, extensivo aos passageiros, e outros exigidos por Lei?'
        ]
    },
'TRANSCELO TRANSPORTES LTDA': {
        'Documentação': [
            '1 - Os documentos obrigatórios para análise e faturamento foram entregues dentro do prazo acordado em contrato?',
            '2 - A contratada apresentou todas as documentações exigidas, conforme contrato com os devidos recolhimentos e pagamentos?',
            '3 - A Nota Fiscal foi emitida com dados corretos?',
            '4 - A contratada possui autorizações atualizadas (licenças específicas) que a habilitem para a prestação de serviços de transporte, expedidas pelos órgãos competentes (ANTT, DER, DETRAN, BHTRANS e/ou quaisquer outros);'
            '5 - A Contratada possui seguro de Acidentes pessoais, extensivo aos passageiros, e outros exigidos por Lei?'
        ]
    },
'MINASCOPY NACIONAL EIRELI': {
        'Documentação': [
            '1 - Os documentos obrigatórios para análise e faturamento foram entregues dentro do prazo acordado em contrato?',
            '2 - A contratada apresentou todas as documentações exigidas, conforme contrato com os devidos recolhimentos e pagamentos?',
            '3 - A Nota Fiscal foi emitida com dados corretos?',
            '4 - CND e CRT, relativos à Regularidade Fiscal e Trabalhista, estão atualizados?'
        ]
    },
'OTIMIZA VIGILÂNCIA E SEG. PATRIMONIAL': {
        'Documentação': [
            '1 - Os documentos obrigatórios para análise e faturamento foram entregues dentro do prazo acordado em contrato?',
            '2 - A contratada apresentou todas as documentações exigidas, conforme contrato com os devidos recolhimentos e pagamentos?',
            '3 - A Nota Fiscal foi emitida com dados corretos?',
            '4 - CND e CRT, relativos à Regularidade Fiscal e Trabalhista, estão atualizados?'
        ]
    },
    # Adicione outros fornecedores conforme necessário
}

# Título
st.markdown(
    "<h1 style='text-align: left; font-family: Open Sauce; color: #104D73;'>"
    'ADFS - AVALIAÇÃO DE DESEMPENHO DE FORNECEDORES DE SERVIÇOS</h1>',
    unsafe_allow_html=True
)

st.subheader('Categoria: Documentação')

st.write('---')

# Subtitulo
if fornecedor and unidade and periodo:
    st.subheader(f'Contratada/Fornecedor: {fornecedor}')
    st.write('Vigência: 02/01/2025 a 31/12/2025')
    st.write(f'Unidade: {unidade}')
    st.write(f'Período avaliado: {periodo}')
    st.write('---')

    # Determinação das abas
    tab1, = st.tabs(['Documentação'])

    respostas = []
    perguntas = []

    # Obter perguntas específicas do fornecedor
    perguntas_fornecedor = perguntas_por_fornecedor.get(fornecedor, {})

    with tab1:
        perguntas_tab1 = perguntas_fornecedor.get('Documentação', [])
        for pergunta in perguntas_tab1:
            resposta = st.selectbox(pergunta, options=opcoes, index=None, placeholder='Selecione uma opção', key=pergunta)
            respostas.append(resposta)
            perguntas.append(pergunta)

    st.sidebar.write('---')

    # Após coletar as perguntas e respostas de cada aba
    categorias = (
            ['Documentação'] * len(perguntas_tab1)
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
            nome_arquivo = f'{nome_fornecedor}_{nome_periodo}_{unidade}_SUP.xlsx'

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
