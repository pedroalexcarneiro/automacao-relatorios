import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
import io

st.set_page_config(page_title="Gerador de Relatórios ITS", layout="wide", page_icon="📊")

st.title("📊 Automação de Relatórios Semanais")
st.markdown("Faça o upload da sua base de dados atualizada e das apresentações da semana passada (como molde) para gerar os ficheiros novos mantendo a formatação.")

# ==========================================
# UPLOAD DE FICHEIROS NO SITE
# ==========================================
col1, col2, col3 = st.columns(3)
with col1:
    excel_file = st.file_uploader("1. Base de Dados (.xlsx)", type=["xlsx"])
with col2:
    template_dig_file = st.file_uploader("2. PPT Digital da Semana Passada (.pptx)", type=["pptx"])
with col3:
    template_voz_file = st.file_uploader("3. PPT Voz da Semana Passada (.pptx)", type=["pptx"])

ciclos_selecionados = st.multiselect(
    "Selecione os ciclos que deseja incluir:",
    ['Ciclo 1', 'Ciclo 2', 'Ciclo 3', 'Ciclo 4', 'Ciclo 5'],
    default=['Ciclo 1', 'Ciclo 2', 'Ciclo 3', 'Ciclo 4']
)

# ==========================================
# LÓGICA DE PROCESSAMENTO INTELIGENTE
# ==========================================
if excel_file and template_dig_file and template_voz_file and ciclos_selecionados:
    if st.button("🚀 Gerar Relatórios Atualizados", use_container_width=True):
        
        with st.spinner('Processando dados e desenhando gráficos sem quebrar o layout...'):
            df = pd.read_excel(excel_file, sheet_name='BASE BRUTA')
            df = df[df['Ciclo'].isin(ciclos_selecionados)].copy()

            prod_map = {
                'CRC': 'Serasa - CRC',
                '0800 Premium': 'Serasa - 0800 Premium',
                'Chat': 'Serasa - CHAT',
                'Central de Ajuda': 'Serasa - Central de Ajuda',
                'Cadastro Positivo': 'SERASA - Cadastro Positivo',
                'Reclame Aqui': 'Serasa - Reclame Aqui',
                'GOV': 'SERASA - GOV'
            }
            df['Produto'] = df['Produto'].map(prod_map).fillna(df['Produto'])

            def get_volumes(prod, time_type):
                return [df[(df['Produto'] == prod) & (df['Time'] == time_type)].groupby('Ciclo')['ID'].count().get(c, 0) for c in ciclos_selecionados]

            def get_medias(prod, time_type):
                return [float(df[(df['Produto'] == prod) & (df['Time'] == time_type)].groupby('Ciclo')['Nota'].mean().round(2).get(c, 0)) for c in ciclos_selecionados]

            def get_nc_ncg(prod, class_type):
                return [df[(df['Produto'] == prod) & (df['Classificação'] == class_type)].groupby('Ciclo')['ID'].count().get(c, 0) for c in ciclos_selecionados]

            def obter_dados_texto(prod_name):
                df_p = df[df['Produto'] == prod_name]
                vol = len(df_p)
                fdb = df_p['Feedback Aplicado'].sum()
                med = f"{df_p['Nota'].mean():.2f}".replace('.', ',') if vol > 0 else "0,00"
                just = len(df_p[df_p['Detalhe SLA'] == 'Em Abono'])
                return str(vol), str(fdb), med, str(just)

            def processar_ppt(template_bytes, mapa_produtos):
                prs = Presentation(template_bytes)
                for slide_idx_vol, dados in mapa_produtos.items():
                    slide_idx_nc = slide_idx_vol + 1
                    prod_name = dados['prod']
                    slide_vol = prs.slides[slide_idx_vol]
                    slide_nc = prs.slides[slide_idx_nc]
                    
                    # 1. Atualizar Gráficos
                    for s in [slide_vol, slide_nc]:
                        for shape in s.shapes:
                            if shape.has_chart:
                                chart = shape.chart
                                title = chart.chart_title.text_frame.text.upper() if chart.has_title else ""
                                c_data = CategoryChartData()
                                c_data.categories = ciclos_selecionados
                                
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

                    # 2. Atualizar Textos (preservando formato)
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

            # Mapas com as referências numéricas passadas
            mapa_dig = {
                3: { 'prod': 'Serasa - Central de Ajuda', 'vol_antigos': ["43", "140", "[VOL]"], 'fdb_antigos': ["41", "137", "[FDB]"], 'med_antigos': ["85,47", "82,57", "[MED]"], 'just_antigos': ["4 monitorias", "5 monitorias", "[JUST] monitorias"] },
                6: { 'prod': 'Serasa - CHAT', 'vol_antigos': ["78", "316", "[VOL]"], 'fdb_antigos': ["76", "308", "[FDB]"], 'med_antigos': ["88,92", "83,87", "[MED]"], 'just_antigos': ["8 monitorias", "15 monitorias", "0 monitorias", "[JUST] monitorias"] },
                9: { 'prod': 'SERASA - GOV', 'vol_antigos': ["9", "33", "[VOL]"], 'fdb_antigos': ["9", "31", "[FDB]"], 'med_antigos': ["77,78", "88,97", "[MED]"], 'just_antigos': ["1 monitorias", "0 monitorias", "[JUST] monitorias"] },
                12:{ 'prod': 'Serasa - Reclame Aqui', 'vol_antigos': ["17", "65", "[VOL]"], 'fdb_antigos': ["17", "62", "[FDB]"], 'med_antigos': ["72,41", "76,11", "[MED]"], 'just_antigos': ["3 monitoria", "3 monitorias", "0 monitorias", "[JUST] monitorias"] }
            }

            mapa_voz = {
                3: { 'prod': 'Serasa - 0800 Premium', 'vol_antigos': ["63", "176", "[VOL]"], 'fdb_antigos': ["61", "173", "[FDB]"], 'med_antigos': ["81,73", "78,97", "[MED]"], 'just_antigos': ["1 monitoria", "10 monitorias", "0 feedbacks", "10 feedbacks", "[JUST] monitorias"] },
                6: { 'prod': 'Serasa - CRC', 'vol_antigos': ["76", "447", "[VOL]"], 'fdb_antigos': ["52", "407", "[FDB]"], 'med_antigos': ["74,80", "72,13", "[MED]"], 'just_antigos': ["2 monitorias", "43 monitorias", "10 monitorias", "[JUST] monitorias"] },
                9: { 'prod': 'SERASA - Cadastro Positivo', 'vol_antigos': ["7", "23", "[VOL]"], 'fdb_antigos': ["7", "21", "[FDB]"], 'med_antigos': ["90,57", "89,04", "[MED]"], 'just_antigos': ["0 monitorias", "[JUST] monitorias"] }
            }
            
            ppt_dig_final = processar_ppt(template_dig_file, mapa_dig)
            ppt_voz_final = processar_ppt(template_voz_file, mapa_voz)

        st.success("Tudo pronto! Relatórios gerados com a formatação intacta.")
        
        col_down1, col_down2 = st.columns(2)
        with col_down1:
            st.download_button("📥 Baixar Relatório Digital", data=ppt_dig_final, file_name="Relatório_Digital_NOVO.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        with col_down2:
            st.download_button("📥 Baixar Relatório de Voz", data=ppt_voz_final, file_name="Relatório_Voz_NOVO.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
