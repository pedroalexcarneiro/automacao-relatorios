from datetime import date
import pandas as pd
import pytest
from regras_monitorias import (
    preparar_base, parse_data, parse_nota, classificar_nota, somar_dias_uteis,
    feriados_sp, _feriados_fallback, canal_padrao, normalizar, COLUNAS_OBRIGATORIAS)

def base_minima(**over):
    linha = {"ID": 1, "Questionario": "SERASA_2", "Nota": 100, "Operador": "x", "Monitor": "Joao Carlos Bezerra",
             "Canal": "CRC", "Data Monitoria": "09/11/2026 00:00", "Feedback Aplicado": "Não",
             "Data Feedback": None, "Status": "concluida"}
    linha.update(over)
    return linha

def df_de(*linhas):
    return pd.DataFrame(list(linhas))

CICLO1 = {"Ciclo 1": {"ativo": True, "ini": date(2026,10,1), "fim": date(2026,11,30)}}
HOJE = date(2026,11,9)

# ---- Notas
def test_classificacao_notas():
    assert classificar_nota(0.0) == "NCG"
    assert classificar_nota(1.0) == "NC"
    assert classificar_nota(99.0) == "NC"
    assert classificar_nota(100.0) == "Excelência"
def test_nota_invalida_nao_vira_zero():
    assert parse_nota("") is None and parse_nota(None) is None
    assert parse_nota("abc") is None and parse_nota(150) is None and parse_nota(-1) is None
    assert parse_nota("87,5") == 87.5

# ---- Dias úteis e feriados
def test_ncg_segunda_vence_terca_e_nc_quarta():
    f = feriados_sp([2026])
    assert somar_dias_uteis(date(2026,11,9), 1, f) == date(2026,11,10)
    assert somar_dias_uteis(date(2026,11,9), 2, f) == date(2026,11,11)
def test_pula_fim_de_semana_e_feriado_consciencia_negra():
    f = feriados_sp([2026])
    # Quinta 19/11: NCG -> sexta 20 é feriado -> segunda 23
    assert somar_dias_uteis(date(2026,11,19), 1, f) == date(2026,11,23)
    # Quinta 19/11: NC -> 23 e 24
    assert somar_dias_uteis(date(2026,11,19), 2, f) == date(2026,11,24)
    # Sexta 13/11: NCG -> segunda 16; NC -> terça 17
    assert somar_dias_uteis(date(2026,11,13), 1, f) == date(2026,11,16)
    assert somar_dias_uteis(date(2026,11,13), 2, f) == date(2026,11,17)
def test_sabado_conta_a_partir_da_segunda():
    f = feriados_sp([2026])
    assert somar_dias_uteis(date(2026,11,7), 1, f) == date(2026,11,9)
def test_feriados_sp_inclui_9_julho_e_fallback_pascoa():
    f = feriados_sp([2026])
    assert date(2026,7,9) in f          # SP estadual
    assert date(2026,4,3) in f          # Sexta-feira Santa 2026
    assert date(2026,1,25) not in f     # feriado municipal de SP não conta
    assert date(2026,6,4) in _feriados_fallback(2026)  # Corpus Christi 2026

# ---- Datas
def test_parse_data_formatos_e_invalidas():
    assert parse_data("09/11/2026 00:00") == pd.Timestamp("2026-11-09")
    assert parse_data("09/11/2026").date() == date(2026,11,9)
    assert pd.isna(parse_data("31/02/2026 10:00"))
    assert pd.isna(parse_data("2026-11-09"))
    assert pd.isna(parse_data(None))

# ---- Canais
def test_canal_case_insensitive_e_cadastro_pelo_nome():
    assert canal_padrao("Gov")[0] == "SERASA - GOV"
    assert canal_padrao("  central  de ajuda ")[0] == "Serasa - Central de Ajuda"
    assert canal_padrao("Cadastro Positivo") == ("SERASA - Cadastro Positivo", True)
    assert canal_padrao("Desconhecido")[1] is False

# ---- Regras de feedback
def test_feedback_sim_dentro_e_fora_do_prazo():
    df = df_de(
        base_minima(ID=1, Nota=0, **{"Data Monitoria":"09/11/2026 10:00", "Feedback Aplicado":"Sim", "Data Feedback":"10/11/2026 09:00"}),  # NCG dentro
        base_minima(ID=2, Nota=0, **{"Data Monitoria":"09/11/2026 10:00", "Feedback Aplicado":"Sim", "Data Feedback":"11/11/2026 09:00"}),  # NCG fora
    )
    out, _, _ = preparar_base(df, CICLO1, hoje=HOJE)
    g = dict(zip(out.ID, out.SLA_Grupo))
    assert g == {"1": "No Prazo", "2": "Fora do Prazo"} or g == {1: "No Prazo", 2: "Fora do Prazo"}
def test_nao_aplicado_no_prazo_e_vencido():
    df = df_de(base_minima(ID=1, Nota=50, **{"Data Monitoria":"09/11/2026 10:00"}),   # NC venc 11/11
               base_minima(ID=2, Nota=0,  **{"Data Monitoria":"02/11/2026 10:00"}))   # NCG venc 03/11
    out, _, _ = preparar_base(df, CICLO1, hoje=HOJE)
    g = {str(k): v for k, v in zip(out.ID, out.SLA_Grupo)}
    assert g == {"1": "No Prazo", "2": "Fora do Prazo"}
def test_abonado_é_justificado_e_não_feedback_realizado():
    df = df_de(base_minima(ID=1, **{"Feedback Aplicado":"Abonado"}))
    out, _, _ = preparar_base(df, CICLO1, hoje=HOJE)
    assert out.SLA_Grupo.iloc[0] == "Justificado" and out.Fdb_Count.iloc[0] == 0
def test_sim_sem_data_feedback_vira_inconsistente_e_sinalizado():
    df = df_de(base_minima(ID=7, **{"Feedback Aplicado":"Sim", "Data Feedback":None}))
    out, resumo, problemas = preparar_base(df, CICLO1, hoje=HOJE)
    assert out.SLA_Grupo.iloc[0] == "Inconsistente"
    assert resumo["inconsistencias"] == 1
    assert "sem Data Feedback" in problemas.Motivo.iloc[0]
def test_contestada_conta_normal():
    df = df_de(base_minima(ID=1, Status="contestada", **{"Feedback Aplicado":"Sim", "Data Feedback":"10/11/2026 08:00"}))
    out, _, _ = preparar_base(df, CICLO1, hoje=HOJE)
    assert len(out) == 1 and out.Fdb_Count.iloc[0] == 1 and out.SLA_Grupo.iloc[0] == "No Prazo"

# ---- Estrutura: IDs, datas, notas, ciclos, adicionais
def test_ids_repetidos_ignorados_com_aviso():
    df = df_de(base_minima(ID=5), base_minima(ID=5))
    out, resumo, problemas = preparar_base(df, CICLO1, hoje=HOJE)
    assert len(out) == 1 and resumo["ids_duplicados"] == 1
    assert any("repetido" in m for m in problemas.Motivo)
def test_data_e_nota_invalidas_sinalizadas():
    df = df_de(base_minima(ID=1, **{"Data Monitoria":"xx"}), base_minima(ID=2, Nota="")) 
    out, resumo, problemas = preparar_base(df, CICLO1, hoje=HOJE)
    assert len(out) == 0
    assert resumo["datas_invalidas"] == 1 and resumo["notas_invalidas"] == 1
    assert len(problemas) == 2
def test_ciclo_pela_data_monitoria_ignora_coluna_ciclo():
    df = df_de(base_minima(ID=1, Ciclo=1, **{"Data Monitoria":"05/10/2026 10:00"}),
               base_minima(ID=2, Ciclo=1, **{"Data Monitoria":"20/11/2026 10:00"}))
    ciclos = {"Ciclo 1": {"ativo": True, "ini": date(2026,10,1), "fim": date(2026,10,31)},
              "Ciclo 2": {"ativo": True, "ini": date(2026,11,1), "fim": date(2026,11,30)}}
    out, resumo, _ = preparar_base(df, ciclos, hoje=HOJE)
    assert dict(zip(out.ID, out.Ciclo)) == {1: "Ciclo 1", 2: "Ciclo 2"} or dict(zip(out.ID.astype(str), out.Ciclo)) == {"1":"Ciclo 1","2":"Ciclo 2"}
    assert resumo["fora_dos_ciclos"] == 0
def test_ciclos_sobrepostos_erro():
    ciclos = {"Ciclo 1": {"ativo": True, "ini": date(2026,10,1), "fim": date(2026,10,20)},
              "Ciclo 2": {"ativo": True, "ini": date(2026,10,15), "fim": date(2026,10,31)}}
    with pytest.raises(ValueError, match="sobrepõem"):
        preparar_base(df_de(base_minima()), ciclos, hoje=HOJE)
def test_monitoras_adicionais_excluidas():
    df = df_de(base_minima(ID=1), base_minima(ID=2, Questionario="Monitoras Adicionais"))
    out, resumo, _ = preparar_base(df, CICLO1, hoje=HOJE)
    assert len(out) == 1 and resumo["monitoras_adicionais"] == 1
def test_colunas_obrigatorias_ausentes_erro():
    with pytest.raises(ValueError, match="obrigatórias"):
        preparar_base(df_de(base_minima()).drop(columns=["Canal"]), CICLO1, hoje=HOJE)

# ---- Base real
def test_base_real_fechamento():
    raw = pd.read_excel("base.xlsx", sheet_name="Monitorias")
    ciclos = {"Ciclo 1": {"ativo": True, "ini": date(2026,10,1), "fim": date(2026,11,30)}}
    out, resumo, problemas = preparar_base(raw, ciclos, hoje=date(2026,11,9))
    assert resumo["processadas"] == 303 == len(out)
    assert resumo["datas_invalidas"] == 0 and resumo["notas_invalidas"] == 0
    assert out.ID.is_unique
    assert out.Produto_Padrao.value_counts().to_dict() == {
        "Serasa - CRC": 104, "Serasa - CHAT": 90, "Serasa - Central de Ajuda": 48,
        "Serasa - 0800 Premium": 47, "Serasa - Reclame Aqui": 9, "SERASA - GOV": 5}
    assert (out.Produto_Padrao == "SERASA - Cadastro Positivo").sum() == 0
    assert out.SLA_Grupo.value_counts().to_dict().get("Justificado") == 5
    assert resumo["inconsistencias"] == 0
    assert set(out.Classificação) <= {"NC","NCG","Excelência"}
