import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
import io
from datetime import date

st.set_page_config(page_title="Gerador de Relatórios ITS", layout="wide", page_icon="📊")

st.title("📊 Automação de Relatórios Semanais")
st.markdown("Faça o upload da exportação bruta de monitorias e defina os intervalos de datas dos ciclos ativos.")

# ==========================================
# 1. UPLOAD DA BASE BRUTA
# ==========================================
uploaded_file = st.file_uploader("1. Base Bruta de Monitorias (.csv ou .xlsx)", type=["csv", "xlsx"])

# Lista oficial de monitores da Qualidade
MONITORES_QUALIDADE = {
    "Rodrigo Ramos Novais",
    "Carlos Alessandro Silva Paiva Junior",
    "Caio Gabriel Dos Santos Lima",
    "Maria Vitoria Da Rocha",
    "Taiane Moreira De Jesus",
    "Ivone Tavares De Oliveira",
    "Yuri Fernando Dallastra Eggert",
    "Pedro Alexandre Rodrigues Carneiro",
    "Marcus Braga Dos Santos",
    "Ana Paula Silva Dos Santos",
    "Lady Lahanna Dos Santos Sutil"
}

TODOS_CICLOS = ['Ciclo 1', 'Ciclo 2', 'Ciclo 3', 'Ciclo 4']

# Nomes dos templates fixos no repositório do GitHub
TEMPLATE_DIGITAL = 'Relatório - Digital.pptx'
TEMPLATE_VOZ = 'Relatório - Voz.pptx'

def carregar_base(file):
    if file.name.endswith('.csv'):
        try:
            return pd.read_csv(file, sep=None, engine='python', encoding='utf-8-sig')
        except UnicodeDecodeError:
            file.seek(0)
            return pd.read_csv(file, sep=None, engine='python', encoding='latin1')
    else:
        xls = pd.ExcelFile(file)
        sheet = 'BASE BRUTA' if 'BASE BRUTA' in xls.sheet_names else xls.sheet_names[0]
        return pd.read_excel(file, sheet_name=sheet)

# ==========================================
# 2. CONFIGURAÇÃO DE CICLOS E DATAS
# ==========================================
if uploaded_file:
    df_raw = carregar_base(uploaded_file)

    # Identifica colunas de data disponíveis para o filtro de lançamento
    colunas_data = [c for c in df_raw.columns if any(k in c.lower() for k in ['data', 'lancamento', 'lançamento', 'cadastro', 'avaliacao', 'avaliação', 'criacao', 'criação'])]
    if not colunas_data:
        colunas_data = list(df_raw.columns)

    col_data_escolhida = st.selectbox(
        "Selecione a coluna que representa a Data de Lançamento da monitoria:",
        options=colunas_data,
        index=0
    )

    st.subheader("🗓️ Configuração dos 4 Ciclos")
    st.caption("Marque apenas os ciclos que deseja atualizar. Os ciclos desmarcados ficarão zerados nos gráficos.")

    cols_ciclos = st.columns(4)
    config_ciclos = {}

    for idx, nome_ciclo in enumerate(TODOS_CICLOS):
        with cols_ciclos[idx]:
            ativo = st.checkbox(f"Atualizar {nome_ciclo}", value=(idx == 0), key=f"chk_{idx}")
            dt_ini = st.date_input(f"Início ({nome_ciclo})", value=date.today(), disabled=not ativo, key=f"ini_{idx}")
            dt_fim = st.date_input(f"Fim ({nome_ciclo})", value=date.today(), disabled=not ativo, key=f"fim_{idx}")
            config_ciclos[nome_ciclo] = {'ativo': ativo, 'ini': dt_ini, 'fim': dt_fim}

    # ==========================================
    # 3. PROCESSAMENTO E REGRAS DE NEGÓCIO
    # ==========================================
    if st.button("🚀 Gerar Relatórios Atualizados", use_container_width=True):
        with st.spinner("Aplicando fórmulas na base bruta e atualizando apresentações..."):
            df = df_raw.copy()

            # 0) Exclui monitorias lançadas na ficha "Monitoras Adicionais" (não entram no resultado)
            qtd_adicionais = 0
            if 'Questionario' in df.columns:
                eh_adicional = df['Questionario'].astype(str).str.strip().str.lower() == 'monitoras adicionais'
                qtd_adicionais = int(eh_adicional.sum())
                df = df[~eh_adicional].copy()

            # A) Identificação das colunas (por nome, com fallback para a posição original do Excel)
            col_nota = 'Nota' if 'Nota' in df.columns else df.columns[4]                 # Coluna E (índice 4)
            col_monitor = 'Monitor' if 'Monitor' in df.columns else df.columns[16]       # Coluna Q (índice 16)
            col_sla = 'Detalhe SLA' if 'Detalhe SLA' in df.columns else df.columns[-1]   # Última coluna
            col_id = 'ID' if 'ID' in df.columns else df.columns[0]                       # Coluna A (ID)

            # Coluna de Produto
            col_prod = next((c for c in df.columns if 'produto' in c.lower() or 'canal' in c.lower() or 'ilha' in c.lower()), 'Produto')

            # B) Tratamento da Nota e Regra de Classificação: =SE(E2=0;"NCG";SE(E2=100;"Excelência";"NC"))
            df['Nota_Num'] = pd.to_numeric(df[col_nota].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)

            def classificar_nota(nota):
                if nota == 0:
                    return "NCG"
                elif nota == 100:
                    return "Excelência"
                else:
                    return "NC"

            df['Classificação'] = df['Nota_Num'].apply(classificar_nota)

            # C) Regra de Time (Qualidade vs Operação, pelo nome do monitor)
            monitores_norm = {m.strip().lower() for m in MONITORES_QUALIDADE}
            df['Time'] = df[col_monitor].astype(str).str.strip().str.lower().apply(
                lambda x: "Qualidade" if x in monitores_norm else "Operação"
            )

            # D) Regra de Detalhe SLA (Substituir "Vence Hoje" por "No Prazo")
            df['Detalhe SLA'] = df[col_sla].astype(str).str.strip().replace("Vence Hoje", "No Prazo")

            # E) Mapeamento de Ciclos por Intervalo de Data
            df['_data_lancamento'] = pd.to_datetime(df[col_data_escolhida], dayfirst=True, errors='coerce').dt.date
            df['Ciclo'] = None

            for nome_ciclo, cfg in config_ciclos.items():
                if cfg['ativo']:
                    mascara = (df['_data_lancamento'] >= cfg['ini']) & (df['_data_lancamento'] <= cfg['fim'])
                    df.loc[mascara, 'Ciclo'] = nome_ciclo

            # Filtra apenas as linhas que caíram nos ciclos ativos selecionados
            df = df[df['Ciclo'].notna()].copy()

            # F) Padronização dos Nomes de Produtos
            prod_map = {
                'CRC': 'Serasa - CRC',
                '0800 Premium': 'Serasa - 0800 Premium',
                'Chat': 'Serasa - CHAT',
                'Central de Ajuda': 'Serasa - Central de Ajuda',
                'Cadastro Positivo': 'SERASA - Cadastro Positivo',
                'Reclame Aqui': 'Serasa - Reclame Aqui',
                'GOV': 'SERASA - GOV'
            }
            df['Produto_Padrao'] = df[col_prod].astype(str).str.strip().map(prod_map).fillna(df[col_prod].astype(str).str.strip())

            # G) Tratamento de Feedback Aplicado
            col_fdb = next((c for c in df.columns if 'feedback' in c.lower()), None)
            if col_fdb:
                if pd.api.types.is_numeric_dtype(df[col_fdb]):
                    df['Fdb_Count'] = df[col_fdb].fillna(0).astype(int)
                else:
                    df['Fdb_Count'] = df[col_fdb].astype(str).str.strip().str.lower().isin(['1', 'sim', 'aplicado', 'true', 'ok']).astype(int)
            else:
                df['Fdb_Count'] = 0

            # ==========================================
            # 4. FUNÇÕES DE AGREGAÇÃO (MANTENDO CICLOS INATIVOS EM ZERO)
            # ==========================================
            def get_volumes(prod, time_type):
                sub = df[(df['Produto_Padrao'] == prod) & (df['Time'] == time_type)]
                contagem = sub.groupby('Ciclo')[col_id].count()
                return [int(contagem.get(c, 0)) if config_ciclos[c]['ativo'] else 0 for c in TODOS_CICLOS]

            def get_medias(prod, time_type):
                sub = df[(df['Produto_Padrao'] == prod) & (df['Time'] == time_type)]
                medias = sub.groupby('Ciclo')['Nota_Num'].mean().round(2)
                return [float(medias.get(c, 0.0)) if config_ciclos[c]['ativo'] else 0.0 for c in TODOS_CICLOS]

            def get_nc_ncg(prod, class_type):
                sub = df[(df['Produto_Padrao'] == prod) & (df['Classificação'] == class_type)]
                contagem = sub.groupby('Ciclo')[col_id].count()
                return [int(contagem.get(c, 0)) if config_ciclos[c]['ativo'] else 0 for c in TODOS_CICLOS]

            def obter_dados_texto(prod_name):
                df_p = df[df['Produto_Padrao'] == prod_name]
                vol = len(df_p)
                fdb = int(df_p['Fdb_Count'].sum())
                med = f"{df_p['Nota_Num'].mean():.2f}".replace('.', ',') if vol > 0 else "0,00"
                just = len(df_p[df_p['Detalhe SLA'].str.lower() == 'em abono'])
                return str(vol), str(fdb), med, str(just)

            # ==========================================
            # 5. ATUALIZAÇÃO DOS POWERPOINTS
            # ==========================================
            def processar_ppt(caminho_template, mapa_produtos):
                prs = Presentation(caminho_template)
                for slide_idx_vol, dados in mapa_produtos.items():
                    slide_idx_nc = slide_idx_vol + 1
                    prod_name = dados['prod']
                    slide_vol = prs.slides[slide_idx_vol]
                    slide_nc = prs.slides[slide_idx_nc]

                    # 1. Atualizar Gráficos (Sempre com os 4 ciclos no eixo X)
                    for s in [slide_vol, slide_nc]:
                        for shape in s.shapes:
                            if shape.has_chart:
                                chart = shape.chart
                                title = chart.chart_title.text_frame.text.upper() if chart.has_title else ""
                                c_data = CategoryChartData()
                                c_data.categories = TODOS_CICLOS

                                if "QUANTIDADE" in title:
                                    c_data.add_series('Operação', get_volumes(prod_name, 'Operação'))
                                    c_data.add_series('Qualidade', get_volumes(prod_name, 'Qualidade'))
                                    chart.replace_data(c_data)
                                elif "MÉDIA" in title:
                                    c_data.add_series('Operação', get_medias(prod_name, 'Operação'))
                                    c_data.add_series('Qualidade', get_medias(prod_name, 'Qualidade'))
                                    chart.replace_data(c_data)
                                elif "VOLUME" in title:
                                    c_data.add_series('NC', get_nc_ncg(prod_name, 'NC'))
                                    c_data.add_series('NCG', get_nc_ncg(prod_name, 'NCG'))
                                    chart.replace_data(c_data)

                    # 2. Atualizar Textos (preservando formatação)
                    novo_vol, novo_fdb, nova_med, novas_just = obter_dados_texto(prod_name)

                    sub_exatas = []
                    for v_antigo in dados['vol_antigos']: sub_exatas.append((v_antigo, novo_vol))
                    for v_antigo in dados['fdb_antigos']: sub_exatas.append((v_antigo, novo_fdb))
                    for v_antigo in dados['med_antigos']: sub_exatas.append((v_antigo, nova_med))

                    sub_parciais = []
                    for v_antigo in dados['just_antigos']:
                        sub_parciais.append((v_antigo, f"{novas_just} monitorias"))

                    for s in [slide_vol, slide_nc]:
                        for shape in s.shapes:
                            if shape.has_text_frame:
                                for paragraph in shape.text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        texto_limpo = run.text.strip()
                                        for de, para in sub_exatas:
                                            if texto_limpo == de or texto_limpo == f" {de} ":
                                                run.text = run.text.replace(de, para)
                                        for de, para in sub_parciais:
                                            if de in run.text:
                                                run.text = run.text.replace(de, para)

                output = io.BytesIO()
                prs.save(output)
                output.seek(0)
                return output

            mapa_dig = {
                3: { 'prod': 'Serasa - Central de Ajuda', 'vol_antigos': ["43", "140", "[VOL]"], 'fdb_antigos': ["41", "137", "[FDB]"], 'med_antigos': ["85,47", "82,57", "[MED]"], 'just_antigos': ["4 monitorias", "5 monitorias", "[JUST] monitorias"] },
                6: { 'prod': 'Serasa - CHAT', 'vol_antigos': ["78", "316", "[VOL]"], 'fdb_antigos': ["76", "308", "[FDB]"], 'med_antigos': ["88,92", "83,87", "[MED]"], 'just_antigos': ["8 monitorias", "15 monitorias", "0 monitorias", "[JUST] monitorias"] },
                9: { 'prod': 'SERASA - GOV', 'vol_antigos': ["9", "33", "[VOL]"], 'fdb_antigos': ["9", "31", "[FDB]"], 'med_antigos': ["77,78", "88,97", "[MED]"], 'just_antigos': ["1 monitorias", "0 monitorias", "[JUST] monitorias"] },
                12:{ 'prod': 'Serasa - Reclame Aqui', 'vol_antigos': ["17", "65", "[VOL]"], 'fdb_antigos': ["17", "62", "[FDB]"], 'med_antigos': ["72,41", "76,11", "[MED]"], 'just_antigos': ["3 monitoria", "3 monitorias", "0 monitorias", "[JUST] monitorias"] }
            }

            mapa_voz = {
                3: { 'prod': 'Serasa - 0800 Premium', 'vol_antigos': ["63", "176", "[VOL]"], 'fdb_antigos': ["61", "173", "[FDB]"], 'med_antigos': ["81,73", "78,97", "[MED]"], 'just_antigos': ["1 monitoria", "10 monitorias", "0 feedbacks", "10 feedbacks", "[JUST] monitorias"] },
                6: { 'prod': 'Serasa - CRC', 'vol_antigos': ["76", "447", "[VOL]"], 'fdb_antigos': ["52", "407", "[FDB]"], 'med_antigos': ["74,80", "72,13", "[MED]"], 'just_antigos': ["2 monitorias", "43 monitorias", "10 monitorias", "[JUST] monitorias"] },
                9: { 'prod': 'SERASA - Cadastro Positivo', 'vol_antigos': ["7", "23", "[VOL]"], 'fdb_antigos': ["7", "21", "[FDB]"], 'med_antigos': ["90,57", "89,04", "90,523", "[MED]"], 'just_antigos': ["0 monitorias", "[JUST] monitorias"] }
            }

            try:
                ppt_dig_final = processar_ppt(TEMPLATE_DIGITAL, mapa_dig)
                ppt_voz_final = processar_ppt(TEMPLATE_VOZ, mapa_voz)

                st.success(
                    f"Relatórios gerados! {len(df)} monitorias processadas dentro dos ciclos selecionados. "
                    f"{qtd_adicionais} monitorias da ficha 'Monitoras Adicionais' foram desconsideradas."
                )

                col_down1, col_down2, col_down3 = st.columns(3)
                with col_down1:
                    st.download_button("📥 Baixar Relatório Digital", data=ppt_dig_final, file_name="Relatório_Digital_Atualizado.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
                with col_down2:
                    st.download_button("📥 Baixar Relatório de Voz", data=ppt_voz_final, file_name="Relatório_Voz_Atualizado.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
                with col_down3:
                    # Permite baixar a base já tratada (sem as monitorias adicionais) com as colunas criadas para conferência
                    buffer_excel = io.BytesIO()
                    df.to_excel(buffer_excel, index=False, sheet_name='BASE BRUTA')
                    buffer_excel.seek(0)
                    st.download_button("📄 Baixar Base Tratada (.xlsx)", data=buffer_excel, file_name="Base_Monitorias_Tratada.xlsx")

            except Exception as e:
                st.error(f"Erro ao processar os templates: {e}")
