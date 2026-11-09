"""Regras de negócio do relatório semanal de Monitorias e Feedbacks.

Módulo sem dependência do Streamlit, para poder ser testado isoladamente.
"""
import re
import unicodedata
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
from dateutil.easter import easter

FUSO = ZoneInfo("America/Sao_Paulo")
FORMATO_DATA_HORA = "%d/%m/%Y %H:%M"
FORMATO_DATA = "%d/%m/%Y"

# Colunas que a base precisa ter para o processamento
COLUNAS_OBRIGATORIAS = [
    "ID", "Nota", "Monitor", "Canal", "Data Monitoria",
    "Feedback Aplicado", "Data Feedback", "Status",
]

# Dados pessoais: nunca usados nos relatórios e removidos da base exportada
COLUNAS_PESSOAIS = ["CPF Operador", "Telefone Cliente", "Nome Cliente"]

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
    "Lady Lahanna Dos Santos Sutil",
    "Catiane Cavanhol",
}

# Canal da base (normalizado) -> nome de exibição usado nos templates
CANAIS = {
    "crc": "Serasa - CRC",
    "0800 premium": "Serasa - 0800 Premium",
    "cadastro positivo": "SERASA - Cadastro Positivo",
    "central de ajuda": "Serasa - Central de Ajuda",
    "chat": "Serasa - CHAT",
    "reclame aqui": "Serasa - Reclame Aqui",
    "gov": "SERASA - GOV",
}

# Dias úteis até o vencimento do feedback, contados a partir da Data Monitoria
PRAZO_DIAS_UTEIS = {"NCG": 1, "NC": 2, "Excelência": 2}

GRUPOS_SLA = ("No Prazo", "Fora do Prazo", "Justificado")


# ==========================================
# NORMALIZAÇÃO
# ==========================================
def _vazio(valor):
    return valor is None or (not isinstance(valor, str) and pd.isna(valor))


def normalizar(valor):
    """Minúsculas, sem acento e com espaços colapsados. Usado para comparar textos."""
    if _vazio(valor):
        return ""
    texto = unicodedata.normalize("NFKD", str(valor)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", texto).strip().lower()


def normalizar_id(valor):
    if _vazio(valor):
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()


def parse_data(valor):
    """Converte 'dd/MM/yyyy HH:mm' (ou 'dd/MM/yyyy' / datetime) em Timestamp. Inválida -> NaT."""
    if isinstance(valor, datetime):
        return pd.Timestamp(valor)
    if isinstance(valor, date):
        return pd.Timestamp(valor)
    if _vazio(valor):
        return pd.NaT
    texto = str(valor).strip()
    for formato in (FORMATO_DATA_HORA, FORMATO_DATA):
        try:
            return pd.Timestamp(datetime.strptime(texto, formato))
        except ValueError:
            continue
    return pd.NaT


def parse_nota(valor):
    """Nota numérica entre 0 e 100. Qualquer outro valor -> None (não é inventado)."""
    if _vazio(valor):
        return None
    try:
        nota = float(str(valor).strip().replace(",", "."))
    except ValueError:
        return None
    if pd.isna(nota) or nota < 0 or nota > 100:
        return None
    return nota


def classificar_nota(nota):
    if nota == 0:
        return "NCG"
    if nota == 100:
        return "Excelência"
    return "NC"


def canal_padrao(valor):
    """Retorna (nome_de_exibicao, reconhecido). Canal desconhecido mantém o texto original."""
    chave = normalizar(valor)
    if chave in CANAIS:
        return CANAIS[chave], True
    return (str(valor).strip() if chave else "(sem canal)"), False


# ==========================================
# CALENDÁRIO (dias úteis de São Paulo)
# ==========================================
def hoje_sp():
    return datetime.now(FUSO).date()


def _feriados_fallback(ano):
    """Feriados nacionais e estaduais de SP, usados quando a lib `holidays` não está instalada."""
    pascoa = easter(ano)
    fixos = {(1, 1), (4, 21), (5, 1), (7, 9), (9, 7), (10, 12), (11, 2), (11, 15), (12, 25)}
    if ano >= 2024:  # Consciência Negra virou feriado nacional em 2024
        fixos.add((11, 20))
    datas = {date(ano, mes, dia) for mes, dia in fixos}
    datas.add(pascoa - timedelta(days=2))   # Sexta-feira Santa
    datas.add(pascoa + timedelta(days=60))  # Corpus Christi
    return datas


def feriados_sp(anos):
    """Feriados nacionais e estaduais de SP (feriado municipal da capital não entra)."""
    anos = sorted(set(anos))
    try:
        import holidays  # opcional: mais completo e atualizado
        feriados = set(holidays.Brazil(subdiv="SP", years=anos).keys())
    except ImportError:
        feriados = set()
        for ano in anos:
            feriados |= _feriados_fallback(ano)
    return feriados


def somar_dias_uteis(data_ref, n, feriados):
    """Soma n dias úteis a data_ref. A própria data não conta como dia do prazo."""
    atual = data_ref
    contados = 0
    while contados < n:
        atual += timedelta(days=1)
        if atual.weekday() < 5 and atual not in feriados:
            contados += 1
    return atual


# ==========================================
# PROCESSAMENTO
# ==========================================
def _registrar(problemas, ids, motivo):
    for id_ in ids:
        problemas.append({"ID": id_, "Motivo": motivo})


def _classificar_feedback(status_fb, data_fb, vencimento, hoje):
    """Retorna (grupo_sla, motivo_de_inconsistencia_ou_None)."""
    if status_fb == "abonado":
        return "Justificado", None
    if status_fb == "sim":
        if pd.isna(data_fb):
            return "Inconsistente", "Feedback marcado como aplicado sem Data Feedback"
        return ("No Prazo" if data_fb.date() <= vencimento else "Fora do Prazo"), None
    if status_fb == "nao":
        # Sem feedback aplicado: ainda no prazo ou já vencido
        return ("No Prazo" if hoje <= vencimento else "Fora do Prazo"), None
    return "Inconsistente", f"Feedback Aplicado com valor não reconhecido: '{status_fb or 'vazio'}'"


def preparar_base(df_raw, ciclos, hoje=None):
    """Trata a base bruta de monitorias.

    ciclos: {nome: {'ativo': bool, 'ini': date, 'fim': date}}
    Retorna (df_tratado, resumo, problemas), em que problemas é um DataFrame com ID e Motivo
    de cada registro sinalizado ou descartado.
    """
    faltantes = [c for c in COLUNAS_OBRIGATORIAS if c not in df_raw.columns]
    if faltantes:
        raise ValueError("Colunas obrigatórias ausentes na base: " + ", ".join(faltantes))

    ativos = {nome: cfg for nome, cfg in ciclos.items() if cfg["ativo"]}
    if not ativos:
        raise ValueError("Marque pelo menos um ciclo para atualizar.")
    for nome, cfg in ativos.items():
        if cfg["fim"] < cfg["ini"]:
            raise ValueError(f"{nome}: a data final é anterior à data inicial.")
    ordenados = sorted(ativos.items(), key=lambda kv: kv[1]["ini"])
    for (nome_a, a), (nome_b, b) in zip(ordenados, ordenados[1:]):
        if b["ini"] <= a["fim"]:
            raise ValueError(f"Os ciclos {nome_a} e {nome_b} se sobrepõem.")

    hoje = hoje or hoje_sp()
    problemas = []
    resumo = {
        "linhas_lidas": len(df_raw),
        "monitoras_adicionais": 0,
        "ids_duplicados": 0,
        "datas_invalidas": 0,
        "notas_invalidas": 0,
        "fora_dos_ciclos": 0,
        "inconsistencias": 0,
        "canais_nao_reconhecidos": 0,
    }

    df = df_raw.copy().reset_index(drop=True)

    # 0) Ficha "Monitoras Adicionais" fica fora do resultado (regra já existente)
    if "Questionario" in df.columns:
        eh_adicional = df["Questionario"].map(normalizar) == "monitoras adicionais"
        resumo["monitoras_adicionais"] = int(eh_adicional.sum())
        df = df[~eh_adicional].reset_index(drop=True)

    # 1) IDs: sem ID não há como conferir; repetidos ficam só na primeira ocorrência
    df["ID"] = df["ID"].map(normalizar_id)
    _registrar(problemas, df.loc[df["ID"] == "", "ID"], "ID ausente")
    duplicado = (df["ID"] != "") & df["ID"].duplicated(keep="first")
    _registrar(problemas, df.loc[duplicado, "ID"], "ID repetido (registro ignorado)")
    resumo["ids_duplicados"] = int(duplicado.sum())
    df = df[~duplicado].reset_index(drop=True)

    # 2) Data Monitoria: sem data válida não há período nem prazo
    df["_data_mon"] = df["Data Monitoria"].map(parse_data)
    invalida = df["_data_mon"].isna()
    _registrar(problemas, df.loc[invalida, "ID"], "Data Monitoria inválida ou ausente (registro ignorado)")
    resumo["datas_invalidas"] = int(invalida.sum())
    df = df[~invalida].reset_index(drop=True)

    # 3) Nota: somente valores entre 0 e 100
    df["_nota"] = df["Nota"].map(parse_nota)
    invalida = df["_nota"].isna()
    _registrar(problemas, df.loc[invalida, "ID"], "Nota inválida ou ausente (registro ignorado)")
    resumo["notas_invalidas"] = int(invalida.sum())
    df = df[~invalida].reset_index(drop=True)

    df["Nota_Num"] = df["_nota"].astype(float)
    df["Classificação"] = df["Nota_Num"].map(classificar_nota)
    df["Prazo_Dias_Uteis"] = df["Classificação"].map(PRAZO_DIAS_UTEIS)

    # 4) Vencimento do feedback (dias úteis de SP, a partir da Data Monitoria)
    if df.empty:
        feriados = set()
        df["_data_mon_d"] = pd.Series(dtype=object)
        df["Vencimento"] = pd.Series(dtype=object)
    else:
        df["_data_mon_d"] = pd.to_datetime(df["_data_mon"]).dt.date
        ano_min = min(df["_data_mon_d"]).year
        ano_max = max(df["_data_mon_d"]).year
        feriados = feriados_sp(range(ano_min, ano_max + 2))
        df["Vencimento"] = [
            somar_dias_uteis(d, int(n), feriados)
            for d, n in zip(df["_data_mon_d"], df["Prazo_Dias_Uteis"])
        ]

    # 5) Canal -> página do template (sem diferenciar maiúsculas/espaços)
    canais = df["Canal"].map(canal_padrao)
    df["Produto_Padrao"] = [c[0] for c in canais]
    nao_reconhecido = [not c[1] for c in canais]
    resumo["canais_nao_reconhecidos"] = int(sum(nao_reconhecido))
    for id_, canal, falta in zip(df["ID"], df["Canal"], nao_reconhecido):
        if falta:
            problemas.append({"ID": id_, "Motivo": f"Canal não reconhecido: '{canal}' (não aparece nos relatórios)"})

    # 6) Time: Qualidade ou Operação, pelo nome do monitor
    monitores_norm = {normalizar(m) for m in MONITORES_QUALIDADE}
    df["Time"] = df["Monitor"].map(
        lambda v: "Qualidade" if normalizar(v) in monitores_norm else "Operação"
    )

    # 7) Feedback aplicado e situação no prazo
    fb_norm = df["Feedback Aplicado"].map(normalizar)
    fb_data = df["Data Feedback"].map(parse_data)
    grupos = []
    for id_, status_fb, data_fb, venc in zip(df["ID"], fb_norm, fb_data, df["Vencimento"]):
        grupo, motivo = _classificar_feedback(status_fb, data_fb, venc, hoje)
        grupos.append(grupo)
        if motivo:
            problemas.append({"ID": id_, "Motivo": motivo})
    df["SLA_Grupo"] = grupos
    df["Fdb_Count"] = (fb_norm == "sim").astype(int)
    resumo["inconsistencias"] = int((df["SLA_Grupo"] == "Inconsistente").sum())

    # 8) Ciclos definidos pelo usuário, sempre pela Data Monitoria (coluna Ciclo da base é ignorada)
    df["Ciclo"] = None
    for nome, cfg in ativos.items():
        mascara = (df["_data_mon_d"] >= cfg["ini"]) & (df["_data_mon_d"] <= cfg["fim"])
        df.loc[mascara, "Ciclo"] = nome
    fora = df["Ciclo"].isna()
    resumo["fora_dos_ciclos"] = int(fora.sum())
    df = df[~fora].reset_index(drop=True)

    resumo["processadas"] = len(df)
    return df, resumo, pd.DataFrame(problemas, columns=["ID", "Motivo"])
