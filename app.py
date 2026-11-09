import io
import re

import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.chart.data import CategoryChartData

from regras_monitorias import (
    COLUNAS_OBRIGATORIAS,
    COLUNAS_PESSOAIS,
    GRUPOS_SLA,
    hoje_sp,
    preparar_base,
)

st.set_page_config(page_title="Relatórios ITS", layout="wide")

# ==========================================
# CONSTANTES
# ==========================================
TODOS_CICLOS = ['Ciclo 1', 'Ciclo 2', 'Ciclo 3', 'Ciclo 4']
TEMPLATE_DIGITAL = 'Relatório - Digital.pptx'
TEMPLATE_VOZ = 'Relatório - Voz.pptx'
MIME_PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Escala fixa dos gráficos de média (eixo começando em 0)
EIXO_MEDIA_MIN = 0
EIXO_MEDIA_MAX = 100

MESES = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
         'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
ORDINAIS = {1: 'primeiro', 2: 'segundo', 3: 'terceiro', 4: 'quarto'}
PADRAO_MES_CAPA = re.compile(r'^(' + '|'.join(MESES) + r') \d{4}$')


def fmt_int(n):
    return f"{int(n):,}".replace(",", ".")


def carregar_base(file):
    if file.name.endswith('.csv'):
        try:
            return pd.read_csv(file, sep=None, engine='python', encoding='utf-8-sig')
        except UnicodeDecodeError:
            file.seek(0)
            return pd.read_csv(file, sep=None, engine='python', encoding='latin1')
    xls = pd.ExcelFile(file)
    aba = next((a for a in ('Monitorias', 'BASE BRUTA') if a in xls.sheet_names), xls.sheet_names[0])
    return pd.read_excel(xls, sheet_name=aba)


def substituir_parcial(texto, de, para):
    # Só troca o número inteiro: "0 monitorias" não pode casar dentro de "10 monitorias"
    # Também não pode terminar no meio de uma palavra: "3 monitoria" não casa com "3 monitorias"
    padrao = r'(?<![\d,.])' + re.escape(de) + r'(?![\dA-Za-zÀ-ÿ])'
    return re.sub(padrao, para, texto)


# ==========================================
# CABEÇALHO
# ==========================================
st.title("Relatórios Semanais de Monitorias")
st.caption("Atualize os relatórios Digital e de Voz a partir da exportação de monitorias.")

# ==========================================
# 1. BASE BRUTA
# ==========================================
st.header("1. Base bruta")
uploaded_file = st.file_uploader(
    "Exportação de monitorias (.csv ou .xlsx)",
    type=["csv", "xlsx"],
    label_visibility="collapsed"
)

if uploaded_file:
    df_raw = carregar_base(uploaded_file)

    faltantes = [c for c in COLUNAS_OBRIGATORIAS if c not in df_raw.columns]
    if faltantes:
        st.error("A base não tem as colunas obrigatórias: " + ", ".join(faltantes))
        st.stop()

    st.caption(f"{len(df_raw):,} linhas carregadas · arquivo: {uploaded_file.name}".replace(",", "."))

    # ==========================================
    # 2. CICLOS
    # ==========================================
    st.header("2. Ciclos a atualizar")
    st.caption("Marque os ciclos que serão atualizados e informe o período de cada um. "
               "Ciclos desmarcados não aparecem nos gráficos. As datas são comparadas com a Data Monitoria.")

    cols_ciclos = st.columns(4)
    config_ciclos = {}
    hoje = hoje_sp()

    for idx, nome_ciclo in enumerate(TODOS_CICLOS):
        with cols_ciclos[idx]:
            with st.container(border=True):
                st.markdown(f"**{nome_ciclo}**")
                ativo = st.checkbox("Atualizar", value=(idx == 0), key=f"chk_{idx}")
                dt_ini = st.date_input("Início", value=hoje, disabled=not ativo, key=f"ini_{idx}", format="DD/MM/YYYY")
                dt_fim = st.date_input("Fim", value=hoje, disabled=not ativo, key=f"fim_{idx}", format="DD/MM/YYYY")
                config_ciclos[nome_ciclo] = {'ativo': ativo, 'ini': dt_ini, 'fim': dt_fim}

    # ==========================================
    # 3. PROCESSAMENTO
    # ==========================================
    st.header("3. Gerar relatórios")

    if st.button("Gerar relatórios", type="primary", use_container_width=True):
        ciclos_ativos = [c for c in TODOS_CICLOS if config_ciclos[c]['ativo']]
        if not ciclos_ativos:
            st.warning("Marque pelo menos um ciclo para atualizar.")
            st.stop()

        with st.spinner("Processando a base e atualizando as apresentações..."):
            try:
                df, resumo, problemas = preparar_base(df_raw, config_ciclos, hoje=hoje)
            except ValueError as e:
                st.error(str(e))
                st.stop()

            # Funções de agregação (somente ciclos ativos)
            def get_volumes(prod, time_type):
                sub = df[(df['Produto_Padrao'] == prod) & (df['Time'] == time_type)]
                contagem = sub.groupby('Ciclo')['ID'].count()
                return [int(contagem.get(c, 0)) for c in ciclos_ativos]

            def get_medias(prod, time_type):
                sub = df[(df['Produto_Padrao'] == prod) & (df['Time'] == time_type)]
                medias = sub.groupby('Ciclo')['Nota_Num'].mean().round(2)
                return [float(medias.get(c, 0.0)) for c in ciclos_ativos]

            def get_nc_ncg(prod, class_type):
                sub = df[(df['Produto_Padrao'] == prod) & (df['Classificação'] == class_type)]
                contagem = sub.groupby('Ciclo')['ID'].count()
                return [int(contagem.get(c, 0)) for c in ciclos_ativos]

            def get_sla_pct(prod, class_type, grupo):
                # Percentual (0 a 1) de cada grupo de SLA dentro de NC ou NCG, por ciclo.
                # "Inconsistente" fica fora do cálculo (não há data confiável para classificar).
                sub = df[(df['Produto_Padrao'] == prod) & (df['Classificação'] == class_type)]
                valores = []
                for c in ciclos_ativos:
                    s = sub[(sub['Ciclo'] == c) & (sub['SLA_Grupo'].isin(GRUPOS_SLA))]
                    total = len(s)
                    if total == 0:
                        valores.append(0.0)
                    else:
                        valores.append(round(float((s['SLA_Grupo'] == grupo).sum()) / total, 4))
                return valores

            def obter_dados_texto(prod_name):
                # Canal sem registros (ex.: Cadastro Positivo) sai tudo zerado
                df_p = df[df['Produto_Padrao'] == prod_name]
                vol = len(df_p)
                fdb = int(df_p['Fdb_Count'].sum())
                med = f"{df_p['Nota_Num'].mean():.2f}".replace('.', ',') if vol > 0 else "0,00"
                just = int((df_p['SLA_Grupo'] == 'Justificado').sum())
                return str(vol), str(fdb), med, str(just)

            # ==========================================
            # MAPAS DE PRODUTOS E TEXTOS DOS TEMPLATES
            # ==========================================
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

            # ==========================================
            # PERÍODO (MÊS E CICLOS) PARA OS TEXTOS DO TEMPLATE
            # ==========================================
            ativos_idx = [TODOS_CICLOS.index(c) + 1 for c in ciclos_ativos]
            inicio_ref = config_ciclos[ciclos_ativos[0]]['ini']
            mes_ref = MESES[inicio_ref.month - 1]
            ano_ref = inicio_ref.year

            meses_envolvidos = {config_ciclos[c]['ini'].month for c in ciclos_ativos}
            if len(meses_envolvidos) > 1:
                st.warning("Os ciclos marcados caem em meses diferentes. O texto usará o mês do primeiro ciclo.")

            ords = [ORDINAIS[i] for i in ativos_idx]
            lista_ords = ords[0] if len(ords) == 1 else ", ".join(ords[:-1]) + " e " + ords[-1]
            palavra_ciclo = 'ciclo' if len(ords) == 1 else 'ciclos'
            trecho_ciclos = f"do {lista_ords} {palavra_ciclo} de {mes_ref.lower()}"

            def atualizar_periodo(prs):
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if not shape.has_text_frame:
                            continue
                        for p in shape.text_frame.paragraphs:
                            for run in p.runs:
                                if PADRAO_MES_CAPA.match(run.text.strip()):
                                    run.text = f"{mes_ref} {ano_ref}"
                                elif 'do primeiro ciclo de setembro' in run.text:
                                    run.text = run.text.replace('do primeiro ciclo de setembro', trecho_ciclos)

            # ==========================================
            # ATUALIZAÇÃO DOS POWERPOINTS
            # ==========================================
            def processar_ppt(caminho_template, mapa_produtos):
                prs = Presentation(caminho_template)
                for slide_idx_vol, dados in mapa_produtos.items():
                    slide_idx_nc = slide_idx_vol + 1
                    prod_name = dados['prod']
                    slide_vol = prs.slides[slide_idx_vol]
                    slide_nc = prs.slides[slide_idx_nc]

                    # 1. Gráficos (somente os ciclos ativos no eixo X)
                    for s in [slide_vol, slide_nc]:
                        for shape in s.shapes:
                            if shape.has_chart:
                                chart = shape.chart
                                title = chart.chart_title.text_frame.text.upper() if chart.has_title else ""

                                if "CUMPRIMENTO DO SLA" in title:
                                    classe = 'NCG' if title.strip().endswith('NCG') else 'NC'
                                    c_data = CategoryChartData(number_format='0%')
                                    c_data.categories = ciclos_ativos
                                    for grupo in ['No Prazo', 'Fora do Prazo', 'Justificado']:
                                        c_data.add_series(grupo, get_sla_pct(prod_name, classe, grupo))
                                    chart.replace_data(c_data)

                                elif "QUANTIDADE" in title:
                                    c_data = CategoryChartData()
                                    c_data.categories = ciclos_ativos
                                    c_data.add_series('Operação', get_volumes(prod_name, 'Operação'))
                                    c_data.add_series('Qualidade', get_volumes(prod_name, 'Qualidade'))
                                    chart.replace_data(c_data)

                                elif "MÉDIA" in title:
                                    c_data = CategoryChartData()
                                    c_data.categories = ciclos_ativos
                                    c_data.add_series('Operação', get_medias(prod_name, 'Operação'))
                                    c_data.add_series('Qualidade', get_medias(prod_name, 'Qualidade'))
                                    chart.replace_data(c_data)
                                    chart.value_axis.minimum_scale = EIXO_MEDIA_MIN
                                    chart.value_axis.maximum_scale = EIXO_MEDIA_MAX

                                elif "VOLUME" in title:
                                    c_data = CategoryChartData()
                                    c_data.categories = ciclos_ativos
                                    c_data.add_series('NC', get_nc_ncg(prod_name, 'NC'))
                                    c_data.add_series('NCG', get_nc_ncg(prod_name, 'NCG'))
                                    chart.replace_data(c_data)

                    # 2. Textos (preservando formatação)
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
                                            if texto_limpo == de:
                                                run.text = run.text.replace(de, para)
                                        for de, para in sub_parciais:
                                            run.text = substituir_parcial(run.text, de, para)

                # 3. Mês da capa/cabeçalhos e ciclos do texto de abertura
                atualizar_periodo(prs)

                output = io.BytesIO()
                prs.save(output)
                output.seek(0)
                return output

            try:
                ppt_dig_final = processar_ppt(TEMPLATE_DIGITAL, mapa_dig)
                ppt_voz_final = processar_ppt(TEMPLATE_VOZ, mapa_voz)

                # Base tratada para conferência (sem dados pessoais)
                colunas_export = [c for c in df.columns if not c.startswith('_') and c not in COLUNAS_PESSOAIS]
                buffer_excel = io.BytesIO()
                with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
                    df[colunas_export].to_excel(writer, index=False, sheet_name='BASE TRATADA')
                    problemas.to_excel(writer, index=False, sheet_name='SINALIZADOS')
                buffer_excel.seek(0)

                # ==========================================
                # 4. RESULTADOS
                # ==========================================
                st.header("4. Resultados")
                m1, m2, m3 = st.columns(3)
                m1.metric("Monitorias processadas", fmt_int(resumo['processadas']))
                m2.metric("Ciclos atualizados", len(ciclos_ativos))
                m3.metric("Registros inconsistentes (feedback)", fmt_int(resumo['inconsistencias']))

                with st.expander("Conferência da base"):
                    st.write(f"- Linhas lidas: {fmt_int(resumo['linhas_lidas'])}")
                    st.write(f"- Ficha 'Monitoras Adicionais' excluída: {fmt_int(resumo['monitoras_adicionais'])}")
                    st.write(f"- IDs repetidos ignorados: {fmt_int(resumo['ids_duplicados'])}")
                    st.write(f"- Datas de monitoria inválidas (ignoradas): {fmt_int(resumo['datas_invalidas'])}")
                    st.write(f"- Notas inválidas (ignoradas): {fmt_int(resumo['notas_invalidas'])}")
                    st.write(f"- Fora dos ciclos selecionados: {fmt_int(resumo['fora_dos_ciclos'])}")
                    st.write(f"- Canais não reconhecidos: {fmt_int(resumo['canais_nao_reconhecidos'])}")
                    if not problemas.empty:
                        st.dataframe(problemas, hide_index=True, use_container_width=True)

                st.caption(f"Período: {mes_ref} {ano_ref} · Ciclos: " + ", ".join(ciclos_ativos))

                st.subheader("Downloads")
                d1, d2, d3 = st.columns(3)
                with d1:
                    st.download_button(
                        "Relatório Digital (.pptx)",
                        data=ppt_dig_final,
                        file_name="Relatório_Digital_Atualizado.pptx",
                        mime=MIME_PPTX,
                        use_container_width=True
                    )
                with d2:
                    st.download_button(
                        "Relatório de Voz (.pptx)",
                        data=ppt_voz_final,
                        file_name="Relatório_Voz_Atualizado.pptx",
                        mime=MIME_PPTX,
                        use_container_width=True
                    )
                with d3:
                    st.download_button(
                        "Base tratada (.xlsx)",
                        data=buffer_excel,
                        file_name="Base_Monitorias_Tratada.xlsx",
                        mime=MIME_XLSX,
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"Erro ao processar os templates: {e}")
