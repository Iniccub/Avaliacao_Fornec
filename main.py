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
unidades = ['CSA - BH', 'CSA - CTG', 'CSA - NL', 'CSA - GZ', 'CSA - DV', 'EPSA', 'ESA', 'AIACOM', 'ILALI', 'ADEODATO', 'SIC']
meses = ['31/01/2025', '28/02/2025', '31/03/2025', '30/04/2025', '31/05/2025', '30/06/2025', '31/07/2025', '31/08/2025',
         '30/09/2025', '31/10/2025', '30/11/2025', '31/12/2025']
fornecedores = ['Cantina Freitas',
                'Expressa Turismo',
                'Acredite Excursões e Exposição itinerante',
                'Leal Viagens e Turismo',
                'MinasCopy Nacional EIRELI',
                'Otimiza Vigilância e Segurança Patrimonial',
                'Petrus',
                'Real Vans',
                'Actur Turismo LTDA',
                'Transcelo',
                'AC Transportes e Serviços LTDA'
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
    'Cantina Freitas': {
        'Atividades Operacionais': [
            '1 - O quantitativo (quadro efetivo) de funcionários da contratada está conforme a necessidade exigida para o atendimento?',
            '2 - A Cantina cumpre a escala de horarios conforme acordado em contrato, observando pontualmente os horários de entrada e saída?',
            '3 - O preposto da contratada atua de maneira presente, efetiva, orientando e zelando pelos seus funcionários?',
            '4 - A Cantina mantem a área cedida sempre em boas condições de conservação e higiene, realizando a limpeza diária de toda a área interna e externa;',
            '5 - Aceitam, sem restrições, a fiscalização por parte do colégio, no que diz respeito ao fiel cumprimento das condições e cláusulas pactuadas?',
            '6 - Mantem profissional NUTRICIONISTA supervisionando permanentemente a prestação dos serviços e, inclusive, promovendo entre os alunos, pais/responsáveis de aluno e empregados do COLÉGIO, a divulgação de bons hábitos alimentares?',
            '7 - Cumprem todas as exigências da Secretaria Municipal de Vigilância Sanitária e da Secretaria Municipal de Posturas e/ou Regulação Urbana e demais órgãos públicos de fiscalização e de normatização?',
            '8 - Recolhem todo o lixo produzido durante o desempenho de sua atividade, de forma a descartá-lo adequadamente no ponto de coleta, obedecendo e cumprindo todas as exigências da Secretaria Municipal?',
            '9 - Procedem a desinsetização e desratização da área cedida durante o período de férias do COLÉGIO, na data estabelecida, previamente, pela CONTRATANTE?'
        ],
        'Segurança': [
            '1 - Fornece aos seus empregados os equipamentos/materiais, uniformes e EPI’s (Equipamentos de Proteção Individual) necessários para a realização dos serviços?',
            '2 - Os funcionários da cantina seguem as normas internas e orientações de segurança da SIC?',
            '3 - Os funcionários zelam pela segurança e cuidado com os funcionarios e alunos do colégio?',
            '4 - Os Funcionarios comunicam , qualquer anormalidade em relação ao andamento dos serviços, prestando à SIC os esclarecimentos, que julgar necessários?',
            '5 - Obedecem às normas internas do colégio e desenvolvem suas atividades sem perturbar as atividades escolares normais?',
            '6 - Os funcionários da contratada transmitem segurança na execução de suas tarefas?'
        ],
        'Documentação': [
            '1 - Os documentos obrigatórios para análise e faturamento foram entregues dentro do prazo acordado em contrato?',
            '2 - A contratada apresentou todas as documentações exigidas, conforme contrato com os devidos recolhimentos e pagamentos?',
            '3 - A Nota Fiscal foi emitida com dados corretos?',
            '4 - CND-FGTS e CRT, relativos à Regularidade Fiscal e Trabalhista, estão atualizados?',
            '5 - Apresentam à SIC cópia autenticada do ALVARÁ DE FUNCIONAMENTO, expedido pelos órgãos competentes, por meio do qual a CONTRATADA ficará autorizada a realizar suas atividades comerciais'
        ],
        'Qualidade': [
            '1 - Os funcionários da contratada atendem com cortesia e presteza quando solicitados?',
            '2 - Os funcionários da contratada comunicam-se com eficácia?',
            '3 - Cumprem rigorosamente todas as normas técnicas relacionadas ao transporte e armazenamento de todo o tipo de ALIMENTO, especialmente as relativas a embalagens, volumes, etc?',
            '4 - A Cantina garante e zela pela boa qualidade dos produtos fornecidos aos usuários(alunos, pais/responsáveis de alunos e empregados do COLÉGIO e terceiros visitantes), em consonância com os parâmetros de qualidade fixados e exigidos pelas normas técnicas pertinentes, expedidas pelo Poder Público e/ou por órgãos e/ou entidades competentes?',
            '5 - A cantina Não comercializa bebidas alcoólicas, cigarros, chicletes, balas, pirulitos, laranjinhas, “chup-chup” e tudo o mais que possa contrariar o bom andamento escolar e/ou causar dano a terceiro, sobretudo, mas não exclusivamente, aos alunos, pais de aluno e empregados do COLÉGIO?',
            '6 - Oferecem e fornecem, sempre que necessário, alimentação especial para os alunos, pais de aluno e/ou empregados que possuam alguma restrição alimentar ou dieta especial, recomendada por profissional de saúde?'
        ]
    },
    'Expressa Turismo': {
        'Atividades Operacionais': [
            '1 - A Contratada disponibiliza veículos em perfeitas condições de conservação e funcionamento mecânico, limpeza externa e interna e de segurança, em conformidade com as exigências legais e demais normas existentes?',
            '2 - A Contratada disponibiliza os veículos após o recebimento da autorização de início dos serviços, nos locais e horários fixados pelo Colégio, cumprindo pontualmente os horários  acordados de saída e retorno, conforme alinhado previamente?',
            '3 - Os profissionais indicados para execução dos serviços objeto do presente contrato, possum conhecimento técnico para sua função e estão devidamente habilitados  para conduzir veículo de transporte de passageiros?',
            '4 - A Contratada cumpre com antecedência e em tempo hábil, informando qualquer motivo que a impossibilite de assumir os serviços conforme estabelecido?',
            '5 - Os Funcionarios da Contratada comunicam, qualquer anormalidade em relação ao andamento dos serviços, prestando à SIC os esclarecimentos, que julgar necessários '
        ],
        'Segurança': [
            '1 - Os funcionários da Contratada durante a prestação de serviço no ambiente interno da escola, seguem as normas internas e orientações demandadas pelo responsável do colégio?',
            '2 - O funcionário da Contratada durante a prestação de serviço segue as normas e regras do Código Brasileiro de Trânsito, respeitando principalmente as regras de limite de velocidade?',
            '3 - Os funcionários zelam pela segurança e cuidado com os alunos e funcionários durante toda a prestação de serviço, apresentando ao serviço sem sinais de embriaguez ou sob efeito de substancia tóxica?',
            '4 - Os funcionários mantem sigilo das informações das quais possuem acesso?',
            '5 - Os veículos disponibilizados para a prestação de serviços estão com os cintos de segurança adequados e em funcionamento, conforme regulamentação específica?',
            '6 - Os veículos enviados para a prestação de serviços, estão equipados com tacógrafos calibrados e aferidos pelo INMETRO?'
        ],
        'Documentação': [
            '1 - Os documentos obrigatórios para análise e faturamento foram entregues dentro do prazo acordado em contrato?',
            '2 - A contratada apresentou todas as documentações exigidas, conforme contrato com os devidos recolhimentos e pagamentos?',
            '3 - A Nota Fiscal foi emitida com dados corretos?',
            '4 - A contratada possui autorizações atualizadas (licenças específicas) que a habilitem para a prestação de serviços de transporte, expedidas pelos órgãos competentes (ANTT, DER, DETRAN, BHTRANS e/ou quaisquer outros);'
            '5 - A Contratada possui seguro de Acidentes pessoais, extensivo aos passageiros, e outros exigidos por Lei?'
        ],
        'Qualidade': [
            '1 - Os profissionais enviados pela Contratada para prestação de serviços são capacitados, qualificados e devidamente treinados, seguindo todas as normas e exigências da legislação brasileira, sobretudo, da legislação de trânsito brasileira?',
            '2 - O profissional da contratada estão devidamente identificados (crachá e/ou uniforme), apresentado com cortesia e presteza, prestando uma boa relação quando solicitado?',
            '3 - Os profissionais da Contratada prestam os esclarecimentos desejados, bem como comunicam, por meio de líder ou diretamente, quaisquer fatos ou anormalidades que porventura possam prejudicar o bom andamento ou o resultado final dos serviços?',
            '4 - Os profissionais da contratada transmitem segurança e conhecimento técnico na execução de suas tarefas?'
        ]
    },
'Leal Viagens e Turismo': {
        'Atividades Operacionais': [
            '1 - A Contratada disponibiliza veículos em perfeitas condições de conservação e funcionamento mecânico, limpeza externa e interna e de segurança, em conformidade com as exigências legais e demais normas existentes?',
            '2 - A Contratada disponibiliza os veículos após o recebimento da autorização de início dos serviços, nos locais e horários fixados pelo Colégio, cumprindo pontualmente os horários  acordados de saída e retorno, conforme alinhado previamente?',
            '3 - Os profissionais indicados para execução dos serviços objeto do presente contrato, possum conhecimento técnico para sua função e estão devidamente habilitados  para conduzir veículo de transporte de passageiros?',
            '4 - Em relação aos locais de destino para os respectivos passeios, onde há necessidade de entrada no local, almoço, lanche e demais infra-estrtura, a Contratada tem entregado esta estrutura conforme acordado previamente?',
            '5 - Os Funcionarios da Contratada comunicam, qualquer anormalidade em relação ao andamento dos serviços, prestando à SIC os esclarecimentos, que julgar necessários?'
        ],
        'Segurança': [
            '1 - Os funcionários da Contratada durante a prestação de serviço no ambiente interno da escola, seguem as normas internas e orientações demandadas pelo responsável do colégio?',
            '2 - O funcionário da Contratada durante a prestação de serviço segue as normas e regras do Código Brasileiro de Trânsito, respeitando principalmente as regras de limite de velocidade?',
            '3 - Os funcionários zelam pela segurança e cuidado com os alunos e funcionários durante toda a prestação de serviço, apresentando ao serviço sem sinais de embriaguez ou sob efeito de substancia tóxica?',
            '4 - Os funcionários mantem sigilo das informações das quais possuem acesso?',
            '5 - Os veículos disponibilizados para a prestação de serviços estão com os cintos de segurança adequados e em funcionamento, conforme regulamentação específica?',
            '6 - Os veículos enviados para a prestação de serviços, estão equipados com tacógrafos calibrados e aferidos pelo INMETRO?'
        ],
        'Documentação': [
            '1 - Os documentos obrigatórios para análise e faturamento foram entregues dentro do prazo acordado em contrato?',
            '2 - A contratada apresentou todas as documentações exigidas, conforme contrato com os devidos recolhimentos e pagamentos?',
            '3 - A Nota Fiscal foi emitida com dados corretos?',
            '4 - A contratada possui autorizações atualizadas (licenças específicas) que a habilitem para a prestação de serviços de transporte, expedidas pelos órgãos competentes (ANTT, DER, DETRAN, BHTRANS e/ou quaisquer outros);'
            '5 - A Contratada possui seguro de Acidentes pessoais, extensivo aos passageiros, e outros exigidos por Lei?'
        ],
        'Qualidade': [
            '1 - Os profissionais enviados pela Contratada para prestação de serviços são capacitados, qualificados e devidamente treinados, seguindo todas as normas e exigências da legislação brasileira, sobretudo, da legislação de trânsito brasileira?',
            '2 - O profissional da contratada estão devidamente identificados (crachá e/ou uniforme), apresentado com cortesia e presteza, prestando uma boa relação quando solicitado?',
            '3 - Os profissionais da Contratada prestam os esclarecimentos desejados, bem como comunicam, por meio de líder ou diretamente, quaisquer fatos ou anormalidades que porventura possam prejudicar o bom andamento ou o resultado final dos serviços?',
            '4 - Os profissionais da contratada transmitem segurança e conhecimento técnico na execução de suas tarefas?'
        ]
    },
'Acredite Excursões e Exposição itinerante': {
        'Atividades Operacionais': [
            '1 - A Contratada disponibiliza veículos em perfeitas condições de conservação e funcionamento mecânico, limpeza externa e interna e de segurança, em conformidade com as exigências legais e demais normas existentes?',
            '2 - A Contratada disponibiliza os veículos após o recebimento da autorização de início dos serviços, nos locais e horários fixados pelo Colégio, cumprindo pontualmente os horários  acordados de saída e retorno, conforme alinhado previamente?',
            '3 - Os profissionais indicados para execução dos serviços objeto do presente contrato, possum conhecimento técnico para sua função e estão devidamente habilitados  para conduzir veículo de transporte de passageiros?',
            '4 - Em relação aos locais de destino para os respectivos passeios, onde há necessidade de entrada no local, almoço, lanche e demais infra-estrtura, a Contratada tem entregado esta estrutura conforme acordado previamente?',
            '5 - Os Funcionarios da Contratada comunicam, qualquer anormalidade em relação ao andamento dos serviços, prestando à SIC os esclarecimentos, que julgar necessários?'
        ],
        'Segurança': [
            '1 - Os funcionários da Contratada durante a prestação de serviço no ambiente interno da escola, seguem as normas internas e orientações demandadas pelo responsável do colégio?',
            '2 - O funcionário da Contratada durante a prestação de serviço segue as normas e regras do Código Brasileiro de Trânsito, respeitando principalmente as regras de limite de velocidade?',
            '3 - Os funcionários zelam pela segurança e cuidado com os alunos e funcionários durante toda a prestação de serviço, apresentando ao serviço sem sinais de embriaguez ou sob efeito de substancia tóxica?',
            '4 - Os funcionários mantem sigilo das informações das quais possuem acesso?',
            '5 - Os veículos disponibilizados para a prestação de serviços estão com os cintos de segurança adequados e em funcionamento, conforme regulamentação específica?',
            '6 - Os veículos enviados para a prestação de serviços, estão equipados com tacógrafos calibrados e aferidos pelo INMETRO?'
        ],
        'Documentação': [
            '1 - Os documentos obrigatórios para análise e faturamento foram entregues dentro do prazo acordado em contrato?',
            '2 - A contratada apresentou todas as documentações exigidas, conforme contrato com os devidos recolhimentos e pagamentos?',
            '3 - A Nota Fiscal foi emitida com dados corretos?',
            '4 - A contratada possui autorizações atualizadas (licenças específicas) que a habilitem para a prestação de serviços de transporte, expedidas pelos órgãos competentes (ANTT, DER, DETRAN, BHTRANS e/ou quaisquer outros);'
            '5 - A Contratada possui seguro de Acidentes pessoais, extensivo aos passageiros, e outros exigidos por Lei?'
        ],
        'Qualidade': [
            '1 - Os profissionais enviados pela Contratada para prestação de serviços são capacitados, qualificados e devidamente treinados, seguindo todas as normas e exigências da legislação brasileira, sobretudo, da legislação de trânsito brasileira?',
            '2 - O profissional da contratada estão devidamente identificados (crachá e/ou uniforme), apresentado com cortesia e presteza, prestando uma boa relação quando solicitado?',
            '3 - Os profissionais da Contratada prestam os esclarecimentos desejados, bem como comunicam, por meio de líder ou diretamente, quaisquer fatos ou anormalidades que porventura possam prejudicar o bom andamento ou o resultado final dos serviços?',
            '4 - Os profissionais da contratada transmitem segurança e conhecimento técnico na execução de suas tarefas?'
        ]
    },
'Real Vans': {
        'Atividades Operacionais': [
            '1 - A Contratada disponibiliza veículos em perfeitas condições de conservação e funcionamento mecânico, limpeza externa e interna e de segurança, em conformidade com as exigências legais e demais normas existentes?',
            '2 - A Contratada disponibiliza os veículos após o recebimento da autorização de início dos serviços, nos locais e horários fixados pelo Colégio, cumprindo pontualmente os horários  acordados de saída e retorno, conforme alinhado previamente?',
            '3 - Os profissionais indicados para execução dos serviços objeto do presente contrato, possum conhecimento técnico para sua função e estão devidamente habilitados  para conduzir veículo de transporte de passageiros?',
            '4 - Em relação aos locais de destino para os respectivos passeios, onde há necessidade de entrada no local, almoço, lanche e demais infra-estrtura, a Contratada tem entregado esta estrutura conforme acordado previamente?',
            '5 - Os Funcionarios da Contratada comunicam, qualquer anormalidade em relação ao andamento dos serviços, prestando à SIC os esclarecimentos, que julgar necessários?'
        ],
        'Segurança': [
            '1 - Os funcionários da Contratada durante a prestação de serviço no ambiente interno da escola, seguem as normas internas e orientações demandadas pelo responsável do colégio?',
            '2 - O funcionário da Contratada durante a prestação de serviço segue as normas e regras do Código Brasileiro de Trânsito, respeitando principalmente as regras de limite de velocidade?',
            '3 - Os funcionários zelam pela segurança e cuidado com os alunos e funcionários durante toda a prestação de serviço, apresentando ao serviço sem sinais de embriaguez ou sob efeito de substancia tóxica?',
            '4 - Os funcionários mantem sigilo das informações das quais possuem acesso?',
            '5 - Os veículos disponibilizados para a prestação de serviços estão com os cintos de segurança adequados e em funcionamento, conforme regulamentação específica?',
            '6 - Os veículos enviados para a prestação de serviços, estão equipados com tacógrafos calibrados e aferidos pelo INMETRO?'
        ],
        'Documentação': [
            '1 - Os documentos obrigatórios para análise e faturamento foram entregues dentro do prazo acordado em contrato?',
            '2 - A contratada apresentou todas as documentações exigidas, conforme contrato com os devidos recolhimentos e pagamentos?',
            '3 - A Nota Fiscal foi emitida com dados corretos?',
            '4 - A contratada possui autorizações atualizadas (licenças específicas) que a habilitem para a prestação de serviços de transporte, expedidas pelos órgãos competentes (ANTT, DER, DETRAN, BHTRANS e/ou quaisquer outros);'
            '5 - A Contratada possui seguro de Acidentes pessoais, extensivo aos passageiros, e outros exigidos por Lei?'
        ],
        'Qualidade': [
            '1 - Os profissionais enviados pela Contratada para prestação de serviços são capacitados, qualificados e devidamente treinados, seguindo todas as normas e exigências da legislação brasileira, sobretudo, da legislação de trânsito brasileira?',
            '2 - O profissional da contratada estão devidamente identificados (crachá e/ou uniforme), apresentado com cortesia e presteza, prestando uma boa relação quando solicitado?',
            '3 - Os profissionais da Contratada prestam os esclarecimentos desejados, bem como comunicam, por meio de líder ou diretamente, quaisquer fatos ou anormalidades que porventura possam prejudicar o bom andamento ou o resultado final dos serviços?',
            '4 - Os profissionais da contratada transmitem segurança e conhecimento técnico na execução de suas tarefas?'
        ]
    },
'AC Transportes e Serviços LTDA': {
        'Atividades Operacionais': [
            '1 - A Contratada disponibiliza veículos em perfeitas condições de conservação e funcionamento mecânico, limpeza externa e interna e de segurança, em conformidade com as exigências legais e demais normas existentes?',
            '2 - A Contratada disponibiliza os veículos após o recebimento da autorização de início dos serviços, nos locais e horários fixados pelo Colégio, cumprindo pontualmente os horários  acordados de saída e retorno, conforme alinhado previamente?',
            '3 - Os profissionais indicados para execução dos serviços objeto do presente contrato, possum conhecimento técnico para sua função e estão devidamente habilitados  para conduzir veículo de transporte de passageiros?',
            '4 - Em relação aos locais de destino para os respectivos passeios, onde há necessidade de entrada no local, almoço, lanche e demais infra-estrtura, a Contratada tem entregado esta estrutura conforme acordado previamente?',
            '5 - Os Funcionarios da Contratada comunicam, qualquer anormalidade em relação ao andamento dos serviços, prestando à SIC os esclarecimentos, que julgar necessários?'
        ],
        'Segurança': [
            '1 - Os funcionários da Contratada durante a prestação de serviço no ambiente interno da escola, seguem as normas internas e orientações demandadas pelo responsável do colégio?',
            '2 - O funcionário da Contratada durante a prestação de serviço segue as normas e regras do Código Brasileiro de Trânsito, respeitando principalmente as regras de limite de velocidade?',
            '3 - Os funcionários zelam pela segurança e cuidado com os alunos e funcionários durante toda a prestação de serviço, apresentando ao serviço sem sinais de embriaguez ou sob efeito de substancia tóxica?',
            '4 - Os funcionários mantem sigilo das informações das quais possuem acesso?',
            '5 - Os veículos disponibilizados para a prestação de serviços estão com os cintos de segurança adequados e em funcionamento, conforme regulamentação específica?',
            '6 - Os veículos enviados para a prestação de serviços, estão equipados com tacógrafos calibrados e aferidos pelo INMETRO?'
        ],
        'Documentação': [
            '1 - Os documentos obrigatórios para análise e faturamento foram entregues dentro do prazo acordado em contrato?',
            '2 - A contratada apresentou todas as documentações exigidas, conforme contrato com os devidos recolhimentos e pagamentos?',
            '3 - A Nota Fiscal foi emitida com dados corretos?',
            '4 - A contratada possui autorizações atualizadas (licenças específicas) que a habilitem para a prestação de serviços de transporte, expedidas pelos órgãos competentes (ANTT, DER, DETRAN, BHTRANS e/ou quaisquer outros);'
            '5 - A Contratada possui seguro de Acidentes pessoais, extensivo aos passageiros, e outros exigidos por Lei?'
        ],
        'Qualidade': [
            '1 - Os profissionais enviados pela Contratada para prestação de serviços são capacitados, qualificados e devidamente treinados, seguindo todas as normas e exigências da legislação brasileira, sobretudo, da legislação de trânsito brasileira?',
            '2 - O profissional da contratada estão devidamente identificados (crachá e/ou uniforme), apresentado com cortesia e presteza, prestando uma boa relação quando solicitado?',
            '3 - Os profissionais da Contratada prestam os esclarecimentos desejados, bem como comunicam, por meio de líder ou diretamente, quaisquer fatos ou anormalidades que porventura possam prejudicar o bom andamento ou o resultado final dos serviços?',
            '4 - Os profissionais da contratada transmitem segurança e conhecimento técnico na execução de suas tarefas?'
        ]
    },
'Transcelo': {
        'Atividades Operacionais': [
            '1 - A Contratada disponibiliza veículos em perfeitas condições de conservação e funcionamento mecânico, limpeza externa e interna e de segurança, em conformidade com as exigências legais e demais normas existentes?',
            '2 - A Contratada disponibiliza os veículos após o recebimento da autorização de início dos serviços, nos locais e horários fixados pelo Colégio, cumprindo pontualmente os horários  acordados de saída e retorno, conforme alinhado previamente?',
            '3 - Os profissionais indicados para execução dos serviços objeto do presente contrato, possum conhecimento técnico para sua função e estão devidamente habilitados  para conduzir veículo de transporte de passageiros?',
            '4 - A Contratada cumpre com antecedência e em tempo hábil, informando qualquer motivo que a impossibilite de assumir os serviços conforme estabelecido?',
            '5 - Os Funcionarios da Contratada comunicam, qualquer anormalidade em relação ao andamento dos serviços, prestando à SIC os esclarecimentos, que julgar necessários?'
        ],
        'Segurança': [
            '1 - Os funcionários da Contratada durante a prestação de serviço no ambiente interno da escola, seguem as normas internas e orientações demandadas pelo responsável do colégio?',
            '2 - O funcionário da Contratada durante a prestação de serviço segue as normas e regras do Código Brasileiro de Trânsito, respeitando principalmente as regras de limite de velocidade?',
            '3 - Os funcionários zelam pela segurança e cuidado com os alunos e funcionários durante toda a prestação de serviço, apresentando ao serviço sem sinais de embriaguez ou sob efeito de substancia tóxica?',
            '4 - Os funcionários mantem sigilo das informações das quais possuem acesso?',
            '5 - Os veículos disponibilizados para a prestação de serviços estão com os cintos de segurança adequados e em funcionamento, conforme regulamentação específica?',
            '6 - Os veículos enviados para a prestação de serviços, estão equipados com tacógrafos calibrados e aferidos pelo INMETRO?'
        ],
        'Documentação': [
            '1 - Os documentos obrigatórios para análise e faturamento foram entregues dentro do prazo acordado em contrato?',
            '2 - A contratada apresentou todas as documentações exigidas, conforme contrato com os devidos recolhimentos e pagamentos?',
            '3 - A Nota Fiscal foi emitida com dados corretos?',
            '4 - A contratada possui autorizações atualizadas (licenças específicas) que a habilitem para a prestação de serviços de transporte, expedidas pelos órgãos competentes (ANTT, DER, DETRAN, BHTRANS e/ou quaisquer outros);'
            '5 - A Contratada possui seguro de Acidentes pessoais, extensivo aos passageiros, e outros exigidos por Lei?'
        ],
        'Qualidade': [
            '1 - Os profissionais enviados pela Contratada para prestação de serviços são capacitados, qualificados e devidamente treinados, seguindo todas as normas e exigências da legislação brasileira, sobretudo, da legislação de trânsito brasileira?',
            '2 - O profissional da contratada estão devidamente identificados (crachá e/ou uniforme), apresentado com cortesia e presteza, prestando uma boa relação quando solicitado?',
            '3 - Os profissionais da Contratada prestam os esclarecimentos desejados, bem como comunicam, por meio de líder ou diretamente, quaisquer fatos ou anormalidades que porventura possam prejudicar o bom andamento ou o resultado final dos serviços?',
            '4 - Os profissionais da contratada transmitem segurança e conhecimento técnico na execução de suas tarefas?'
        ]
    },
'MinasCopy Nacional EIRELI': {
        'Atividades Operacionais': [
            '1 - O quantitativo (quadro efetivo) de funcionários da contratada está conforme especificação e acordado em contrato?',
            '2 - Os funcionários cumprem a escala de serviço, observando pontualmente os horários de entrada e saída, sendo assíduos e pontuais ao trabalho?',
            '3 - A empresa fornece o ponto eletrônico e o mantem em pleno funcionamento, registrando e apurando os horários registrados dos respectivos funcionários?',
            '4 - A SIC é informada previamente das eventuais substituições dos funcionários da contratada?',
            '5 - Os profissionais indicados para execução dos serviços objeto do presente Instrumento, possuem conhecimento técnico necessário para operar e manusear as máquinas, bem como executar as funções que lhe forem atribuídas?'
        ],
        'Segurança': [
            '1 - Os funcionários seguem as normas internas e orientações de segurança da SIC?',
            '2 - Os funcionários mantem sigilo das informações das quais possuem acesso?',
            '3 - Os funcionários zelam pela segurança e cuidado com suas entregas, conforme são demandados?',
            '4 - Os Funcionarios comunicam, qualquer anormalidade em relação ao andamento dos serviços, prestando à SIC os esclarecimentos, que julgar necessários?'
        ],
        'Documentação': [
            '1 - Os documentos obrigatórios para análise e faturamento foram entregues dentro do prazo acordado em contrato?',
            '2 - A contratada apresentou todas as documentações exigidas, conforme contrato com os devidos recolhimentos e pagamentos?',
            '3 - A Nota Fiscal foi emitida com dados corretos?',
            '4 - CND e CRT, relativos à Regularidade Fiscal e Trabalhista, estão atualizados?'
        ],
        'Qualidade': [
            '1 - Os funcionários da contratada executam suas atividades diárias com qualidade, atendendo todas as demandas inerentes ao objeto deste contrato?',
            '2 - Os funcionários da contratada atendem com cortesia e presteza, prestando uma boa relação quando solicitados?',
            '3 - Os funcionários da contratada comunicam-se com eficácia?',
            '4 - Os funcionários da contratada ocupam-se permanentemente no local designado para exercicio de suas funções, não se afastando deste local, salvo em situações de necessidade?',
            '5 - Os funcionários da contratada transmitem segurança na execução de suas tarefas?',
            '6 - Os funcionários da contratada zelam pelos materiais e equipamentos quando estão dentro das dependências do colégio?'
        ]
    },
'Otimiza Vigilância e Segurança Patrimonial': {
        'Atividades Operacionais': [
            '1 - O quantitativo (quadro efetivo) de funcionários da contratada está conforme especificação e acordado em contrato?',
            '2 - Os funcionários cumprem a escala de serviço, observando pontualmente os horários de entrada e saída, sendo assíduos e pontuais ao trabalho?',
            '3 - Na ocorrência de faltas, é providenciada pela contratada a reposição do funcionário no período previsto no contrato?',
            '4 - A empresa fornece o ponto eletrônico e o mantem em pleno funcionamento, registrando e apurando os horários registrados dos respectivos funcionários?',
            '5 - A SIC é  informada previamente das eventuais substituições dos funcionários da contratada?',
            '6 - O preposto da contratada atua de maneira presente, efetiva, orientando e zelando pelos seus funcionários?'
        ],
        'Segurança': [
            '1 - Os funcionários estão devidamente uniformizados (padrão único) e identificados (crachá)?',
            '2 - Os funcionários seguem as normas internas e orientações de segurança da SIC?',
            '3 - Os funcionários mantem sigilo das informações das quais possuem acesso?',
            '4 - Os funcionários zelam pela segurança e cuidado com os funcionarios e alunos do colégio?',
            '5 - Os Funcionarios comunicam, qualquer anormalidade em relação ao andamento dos serviços, prestando à SIC os esclarecimentos, que julgar necessários?'
        ],
        'Documentação': [
            '1 - Os documentos obrigatórios para análise e faturamento foram entregues dentro do prazo acordado em contrato?',
            '2 - A contratada apresentou todas as documentações exigidas, conforme contrato com os devidos recolhimentos e pagamentos?',
            '3 - A Nota Fiscal foi emitida com dados corretos?',
            '4 - CND e CRT, relativos à Regularidade Fiscal e Trabalhista, estão atualizados?'
        ],
        'Qualidade': [
            '1 - Os funcionários da contratada executam suas atividades diárias com qualidade, atendendo todas as demandas inerentes ao objeto deste contrato?',
            '2 - Os funcionários da contratada atendem com cortesia e presteza, prestando uma boa relação quando solicitados?',
            '3 - Os funcionários da contratada comunicam-se com eficácia?',
            '4 - Os funcionários da contratada ocupam-se permanentemente no local designado para exercicio de suas funções, não se afastando deste local, salvo em situações de necessidade?',
            '5 - Os funcionários da contratada transmitem segurança na execução de suas tarefas?',
            '6 - Os funcionários da contratada zelam pelos materiais e equipamentos quando estão dentro das dependências do colégio?'
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
        perguntas_tab3 = perguntas_fornecedor.get('Documentação', [])
        for pergunta in perguntas_tab3:
            resposta = st.selectbox(pergunta, options=opcoes, index=None, placeholder='Selecione uma opção', key=pergunta)
            respostas.append(resposta)
            perguntas.append(pergunta)

    with tab4:
        perguntas_tab4 = perguntas_fornecedor.get('Qualidade', [])
        for pergunta in perguntas_tab4:
            resposta = st.selectbox(pergunta, options=opcoes, index=None, placeholder='Selecione uma opção', key=pergunta)
            respostas.append(resposta)
            perguntas.append(pergunta)

    st.sidebar.write('---')

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
                'Pergunta': perguntas,
                'Resposta': respostas
            })

            # Formata o nome do arquivo com base no fornecedor e período
            nome_fornecedor = fornecedor.replace(' ', '_')
            nome_periodo = periodo.replace('/', '-')
            nome_arquivo = f'{nome_fornecedor}_{nome_periodo}.xlsx'

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
