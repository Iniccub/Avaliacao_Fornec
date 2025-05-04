import streamlit as st
import pandas as pd
import os
from openpyxl import load_workbook
from streamlit_js_eval import streamlit_js_eval
from io import BytesIO

st.set_page_config(
    page_title='Avaliação de Fornecedores.',
    page_icon='CSA.png',
    layout='wide'
)

# Listas fixas
unidades = ['CSA - BH', 'CSA - CTG', 'CSA - NL', 'CSA - GZ', 'CSA - DV', 'EPSA', 'ESA', 'AIACOM']
meses = ['31/01/2025', '28/02/2025', '31/03/2025', '30/04/2025', '31/05/2025', '30/06/2025', '31/07/2025', '31/08/2025',
         '30/09/2025', '31/10/2025', '30/11/2025', '31/12/2025']
fornecedores = ['Cantina Freitas', 'Expressa Turismo', 'Acredite Excursões e Exposição itinerante', 'Leal Viagens e Turismo',
                'MinasCopy Nacional EIRELI', 'Otimiza Vigilância e Segurança', 'Petrus', 'Real Vans', 'Actur Turismo LTDA', 'Transcelo']
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
    tab1, tab2, tab3, tab4 = st.tabs(['Atividades Operacionais', 'Segurança', 'Documentação', 'Qualidade'])

    respostas = []

    with tab1:
        perguntas_tab1 = [
            'O quantitativo (quadro efetivo) de funcionários da contratada está conforme a necessidade exigida para o atendimento?',
            'O prestador/fornecedor cumpre a escala de horarios conforme acordado em contrato, observando pontualmente os horários de entrada e saída?',
            'O preposto da contratada atua de maneira presente, efetiva, orientando e zelando pelos seus funcionários?',
            'O prestador/fornecedor mantem a área cedida sempre em boas condições de conservação e higiene, realizando a limpeza diária de toda a área interna e externa;',
            'Aceitam, sem restrições, a fiscalização por parte do colégio, no que diz respeito ao fiel cumprimento das condições e cláusulas pactuadas?',
            'Mantem profissional NUTRICIONISTA supervisionando permanentemente a prestação dos serviços e, inclusive, promovendo entre os alunos, pais/responsáveis de aluno e empregados do COLÉGIO, a divulgação de bons hábitos alimentares?',
            'Cumprem todas as exigências da Secretaria Municipal de Vigilância Sanitária e da Secretaria Municipal de Posturas e/ou Regulação Urbana e demais órgãos públicos de fiscalização e de normatização?',
            'Recolhem todo o lixo produzido durante o desempenho de sua atividade, de forma a descartá-lo adequadamente no ponto de coleta, obedecendo e cumprindo todas as exigências da Secretaria Municipal?',
            'Procedem a desinsetização e desratização da área cedida durante o período de férias do COLÉGIO, na data estabelecida, previamente, pela CONTRATANTE?'
        ]

        for pergunta in perguntas_tab1:
            resposta = st.selectbox(pergunta, options=opcoes, index=None, placeholder='Selecione uma opção')
            respostas.append(resposta)


    with tab2:
        perguntas_tab2 = [
            'Fornece aos seus empregados os equipamentos/materiais, uniformes e EPI’s (Equipamentos de Proteção Individual) necessários para a realização dos serviços?',
            'Os funcionários da cantina seguem as normas internas e orientações de segurança da SIC?',
            'Os funcionários zelam pela segurança e cuidado com os funcionarios e alunos do colégio?',
            'Os Funcionarios comunicam , qualquer anormalidade em relação ao andamento dos serviços, prestando à SIC os esclarecimentos, que julgar necessários?',
            'Obedecem às normas internas do colégio e desenvolvem suas atividades sem perturbar as atividades escolares normais?',
            'Os funcionários da contratada transmitem segurança na execução de suas tarefas?'
        ]

        for pergunta in perguntas_tab2:
            resposta = st.selectbox(pergunta, options=opcoes, index=None, placeholder='Selecione uma opção', key=pergunta)
            respostas.append(resposta)

    with tab3:
        perguntas_tab3 = [
            'Os documentos obrigatórios para análise e faturamento foram entregues dentro do prazo acordado em contrato?',
            'A contratada apresentou todas as documentações exigidas, conforme contrato com os devidos recolhimentos e pagamentos?',
            'A Nota Fiscal foi emitida com dados corretos?',
            'CND-FGTS e CRT, relativos à Regularidade Fiscal e Trabalhista, estão atualizados?',
            'Apresentam à SIC cópia autenticada do ALVARÁ DE FUNCIONAMENTO, expedido pelos órgãos competentes, por meio do qual a CONTRATADA ficará autorizada a realizar suas atividades comerciais'
        ]

        for pergunta in perguntas_tab3:
            resposta = st.selectbox(pergunta, options=opcoes, index=None, placeholder='Selecione uma opção', key=pergunta)
            respostas.append(resposta)

    with tab4:
        perguntas_tab4 = [
            'Os funcionários da contratada atendem com cortesia e presteza quando solicitados?',
            'Os funcionários da contratada comunicam-se com eficácia?',
            'Cumprem rigorosamente todas as normas técnicas relacionadas ao transporte e armazenamento de todo o tipo de ALIMENTO, especialmente as relativas a embalagens, volumes, etc?',
            'A Cantina garante e zela pela boa qualidade dos produtos fornecidos aos usuários(alunos, pais/responsáveis de alunos e empregados do COLÉGIO e terceiros visitantes), em consonância com os parâmetros de qualidade fixados e exigidos pelas normas técnicas pertinentes, expedidas pelo Poder Público e/ou por órgãos e/ou entidades competentes?',
            'A cantina Não comercializa bebidas alcoólicas, cigarros, chicletes, balas, pirulitos, laranjinhas, “chup-chup” e tudo o mais que possa contrariar o bom andamento escolar e/ou causar dano a terceiro, sobretudo, mas não exclusivamente, aos alunos, pais de aluno e empregados do COLÉGIO?',
            'Oferecem e fornecem, sempre que necessário, alimentação especial para os alunos, pais de aluno e/ou empregados que possuam alguma restrição alimentar ou dieta especial, recomendada por profissional de saúde?'
        ]

        for pergunta in perguntas_tab4:
            resposta = st.selectbox(pergunta, options=opcoes, index=None, placeholder='Selecione uma opção', key=pergunta)
            respostas.append(resposta)

    st.sidebar.write('---')

    if st.sidebar.button('Salvar pesquisa'):
        # Verifica se todas as perguntas foram respondidas
        if None in respostas:
            st.warning('Por favor, responda todas as perguntas antes de salvar.')
        else:
            # Cria DataFrame com as respostas
            perguntas = perguntas_tab1 + perguntas_tab2 + perguntas_tab3 + perguntas_tab4
            df_respostas = pd.DataFrame({
                'Unidade': unidade,
                'Período': periodo,
                'Fornecedor': fornecedor,
                'Pergunta': perguntas,
                'Resposta': respostas
            })

            # Formata o noem do arquivo com base no fornecedor e periodo
            nome_fornecedor = fornecedor.replace(' ', '_')
            nome_periodo = periodo.replace('/', '-')
            nome_arquivo = f'{nome_fornecedor}_{nome_periodo}.xlsx'

            # Salva o DataFrame em um objeto BytesIO
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_respostas.to_excel(writer, index=False)
            output.seek(0)

            # Cria um botão de download no streamlit
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
