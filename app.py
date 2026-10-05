import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
import io

st.set_page_config(page_title="Gerador de Relatórios ITS", layout="wide", page_icon="📊")

st.title("📊 Automação de Relatórios Semanais")
st.markdown("Faça o upload da sua base de dados atualizada e dos templates para gerar as apresentações automaticamente.")

# ==========================================
# UPLOAD DE ARQUIVOS NO SITE
# ==========================================
col1, col2, col3 = st.columns(3)
with col1:
    excel_file = st.file_uploader("1. Base de Dados (.xlsx)", type=["xlsx"])
with col2:
    template_dig_file = st.file_uploader("2. Template Digital (.pptx)", type=["pptx"])
with col3:
    template_voz_file = st.file_uploader("3. Template de Voz (.pptx)", type=["pptx"])

# Seleção de Ciclos
ciclos_selecionados = st.multiselect(
    "Selecione os ciclos que deseja incluir:",
    ['Ciclo 1', 'Ciclo 2', 'Ciclo 3', 'Ciclo 4', 'Ciclo 5'],
    default=['Ciclo 1', 'Ciclo 2', 'Ciclo 3', 'Ciclo 4']
)

# ==========================================
# LÓGICA DE PROCESSAMENTO
# ==========================================
if excel_file and template_dig_file and template_voz_file and ciclos_selecionados:
    if st.button("🚀 Gerar Relatórios Atualizados", use_container_width=True):
        
        with st.spinner('Processando dados e desenhando gráficos...'):
            # Lendo a base
            df = pd.read_excel(excel_file, sheet_name='BASE BRUTA')
            df = df[df['Ciclo'].isin(ciclos_selecionados)].copy()

            # Padronizando nomes
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

            # Funções de Cálculo
            def get_volumes(prod, time_type):
                return [df[(df['Produto'] == prod) & (df['Time'] == time_type)].groupby('Ciclo')['ID'].count().get(c, 0) for c in ciclos_selecionados]

            def get_medias(prod, time_type):
                return [float(df[(df['Produto'] == prod) & (df['Time'] == time_type)].groupby('Ciclo')['Nota'].mean().round(2).get(c, 0)) for c in ciclos_selecionados]

            def get_nc_ncg(prod, class_type):
                return [df[(df['Produto'] == prod) & (df['Classificação'] == class_type)].groupby('Ciclo')['ID'].count().get(c, 0) for c in ciclos_selecionados]

            def get_sla(prod, class_type, sla_type):
                df_p = df[(df['Produto'] == prod) & (df['Classificação'] == class_type)]
                res = []
                for c in ciclos_selecionados:
                    df_c = df_p[df_p['Ciclo'] == c]
                    tot = len(df_c)
                    res.append(round(len(df_c[df_c['Detalhe SLA'] == sla_type]) / tot, 4) if tot > 0 else 0.0)
                return res

            def obter_dados_texto(prod_name):
                df_p = df[df['Produto'] == prod_name]
                vol = len(df_p)
                fdb = df_p['Feedback Aplicado'].sum()
                med = f"{df_p['Nota'].mean():.2f}".replace('.', ',') if vol > 0 else "0,00"
                just = len(df_p[df_p['Detalhe SLA'] == 'Em Abono'])
                return str(vol), str(fdb), med, str(just)

            def processar_ppt(template_bytes, mapa_produtos):
                prs = Presentation(template_bytes)
                
                for slide_idx_vol, prod_name in mapa_produtos.items():
                    slide_idx_nc = slide_idx_vol + 1
                    slide_vol = prs.slides[slide_idx_vol]
                    slide_nc = prs.slides[slide_idx_nc]
                    
                    # 1. ATUALIZA GRÁFICOS
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
                                elif "SLA(%) - NC" in title and "NCG" not in title:
                                    c_data.add_series('No Prazo', get_sla(prod_name, 'NC', 'No Prazo'))
                                    c_data.add_series('Fora do Prazo', get_sla(prod_name, 'NC', 'Fora do Prazo'))
                                    c_data.add_series('Justificado', get_sla(prod_name, 'NC', 'Em Abono'))
                                    chart.replace_data(c_data)
                                elif "SLA(%)" in title and "NCG" in title:
                                    c_data.add_series('No Prazo', get_sla(prod_name, 'NCG', 'No Prazo'))
                                    c_data.add_series('Fora do Prazo', get_sla(prod_name, 'NCG', 'Fora do Prazo'))
                                    c_data.add_series('Justificado', get_sla(prod_name, 'NCG', 'Em Abono'))
                                    chart.replace_data(c_data)
                                    
                    # 2. ATUALIZA TEXTOS (TAGS)
                    novo_vol, novo_fdb, nova_med, novas_just = obter_dados_texto(prod_name)
                    tags = {'[VOL]': novo_vol, '[FDB]': novo_fdb, '[MED]': nova_med, '[JUST]': novas_just}
                    
                    for s in [slide_vol, slide_nc]:
                        for shape in s.shapes:
                            if shape.has_text_frame:
                                for p in shape.text_frame.paragraphs:
                                    for r in p.runs:
                                        for tag, valor in tags.items():
                                            if tag in r.text:
                                                r.text = r.text.replace(tag, valor)
                
                # Salva em memória para o usuário baixar
                output = io.BytesIO()
                prs.save(output)
                output.seek(0)
                return output

            # Mapas de localização dos slides
            mapa_dig = {3: 'Serasa - Central de Ajuda', 6: 'Serasa - CHAT', 9: 'SERASA - GOV', 12: 'Serasa - Reclame Aqui'}
            mapa_voz = {3: 'Serasa - 0800 Premium', 6: 'Serasa - CRC', 9: 'SERASA - Cadastro Positivo'}
            
            # Gera os arquivos em memória
            ppt_dig_final = processar_ppt(template_dig_file, mapa_dig)
            ppt_voz_final = processar_ppt(template_voz_file, mapa_voz)

        st.success("Tudo pronto! Seus relatórios foram gerados e estão prontos para download.")
        
        col_down1, col_down2 = st.columns(2)
        with col_down1:
            st.download_button(label="📥 Baixar Relatório Digital", data=ppt_dig_final, file_name="Relatório_Digital_Atualizado.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        with col_down2:
            st.download_button(label="📥 Baixar Relatório de Voz", data=ppt_voz_final, file_name="Relatório_Voz_Atualizado.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")