from pathlib import Path
import io

import plotly.graph_objects as go
import streamlit as st
import pandas as pd

from carregar_eap import carregar_dados


# =============================================================================
# FORMATADORES
# =============================================================================

def fmt_money(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_percent(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    return f"{v * 100:.2f}%".replace(".", ",")


def fmt_decimal(v):
    if v is None:
        return "—"
    return f"{v:.2f}".replace(".", ",")


def as_percent(v) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    return float(v) * 100


def fmt_date(v) -> str:
    if v is None:
        return "—"
    try:
        return v.strftime("%d/%m/%Y")
    except Exception:
        return str(v)[:10]


# =============================================================================
# ADAPTADOR: converte saída de carregar_dados → dict que o app consome
# =============================================================================

ETAPA_CORES = [
    "green", "blue", "amber",
    "gray", "gray", "gray", "gray", "gray", "gray", "gray",
]


def _build_data(config: dict, df_eap: pd.DataFrame, df_totais: pd.DataFrame) -> dict:
    med_atual = int(config["medicao_atual"])
    idp_val = config["idp"]["valor"] or 1.0

    # ---- Prazos — usar valores diretos da CONFIG (calculados pela planilha) ---
    vig = config["vigencia"]
    exe = config["execucao"]
    tend = config["tendencia"]

    # Vigência: usar dias_decorridos e prazo_total_dias da CONFIG diretamente
    vig_dec  = int(vig["dias_decorridos"] or 0)
    vig_tot  = int(vig["prazo_total_dias"] or 0)
    vig_rest = max(vig_tot - vig_dec, 0)

    # Execução: usar dias_decorridos e prazo_total_dias da CONFIG diretamente
    exe_dec  = int(exe["dias_decorridos"] or 0)
    exe_tot  = int(exe["prazo_total_dias"] or 0)
    exe_rest = max(exe_tot - exe_dec, 0)

    prazos = {
        "prazo_vigencia_data":       fmt_date(vig["fim"]),
        "vigencia_total_dias":       vig_tot,
        "vigencia_decorrida_dias":   vig_dec,
        "vigencia_restante_dias":    vig_rest,
        "vigencia_decorrida_texto":  f"{vig_dec} dias decorridos",
        "vigencia_restante_texto":   f"{vig_rest} dias restantes",

        "prazo_execucao_data":       fmt_date(exe["fim"]),
        "execucao_total_dias":       exe_tot,
        "execucao_decorrida_dias":   exe_dec,
        "execucao_restante_dias":    exe_rest,
        "execucao_decorrida_texto":  f"{exe_dec} dias decorridos",
        "execucao_restante_texto":   f"{exe_rest} dias restantes",

        "nova_data_termino":         tend["nova_data_termino"],
        "tendencia_duracao_texto":   f"Nova duração: {tend['nova_duracao_dias']} dias ({tend['nova_duracao_meses']:.1f} meses)",
        "tendencia_termino_texto":   f"Tendência de término: {fmt_date(tend['nova_data_termino'])}",
    }

    # ---- Medições / financeiro ----------------------------------------
    row_atual = df_totais[df_totais["medicao"] == med_atual]
    row_ant   = df_totais[df_totais["medicao"] == med_atual - 1]

    val_med_atual = float(row_atual["total_medido"].iloc[0]) if len(row_atual) else 0.0
    val_med_acum  = float(row_atual["total_medido_acum"].iloc[0]) if len(row_atual) else 0.0
    pct_mes_atual = float(row_atual["total_pct_mes"].iloc[0]) if len(row_atual) else 0.0
    pct_acum      = float(row_atual["total_pct_acum"].iloc[0]) if len(row_atual) else 0.0

    val_med_ant   = float(row_ant["total_medido_acum"].iloc[0]) if len(row_ant) else 0.0
    pct_ant       = float(row_ant["total_pct_acum"].iloc[0]) if len(row_ant) else 0.0

    valor_contrato = float(df_eap["preco_total_rs"].max() or 0)
    # pega o valor total da EAP (preço total do item de nível 1 que tem maior valor)
    # melhor: somar todos os itens de nível 1 (item sem ponto)
    df_n1 = df_eap[df_eap["item"].apply(
        lambda x: x is not None and "." not in str(x)
    )].drop_duplicates(subset=["item"])
    valor_contrato = float(df_n1["preco_total_rs"].sum()) if len(df_n1) else valor_contrato

    saldo = valor_contrato - val_med_acum
    saldo_pct = saldo / valor_contrato if valor_contrato else 0

    medicoes = {
        "medicao_atual":                  med_atual,
        "avanco_fisico_acumulado":         pct_acum,
        "mes_atual_percentual":            pct_mes_atual,
        "medicoes_anteriores_percentual":  pct_ant,
        "valor_medido_mes_atual":          val_med_atual,
        "valor_medido_acumulado":          val_med_acum,
        "valor_medido_anterior":           val_med_ant,
    }

    # ---- IDP -----------------------------------------------------------
    idp_status = "Obra no prazo" if idp_val >= 1 else ("Atenção" if idp_val >= 0.9 else "Atraso")
    idp = {"valor": idp_val, "status": idp_status, "analise": ""}

    # Análise narrativa baseada no IDP
    nova_data_fmt = fmt_date(tend["nova_data_termino"]) if tend["nova_data_termino"] else fmt_date(exe["fim"])
    if idp_val >= 1.05:
        analise_idp = (
            f"Com IDP = {idp_val:.2f}, a obra apresenta desempenho físico <strong>acima do planejado</strong>. "
            f"O ritmo atual indica que a conclusão ocorrerá antes da data contratual. "
            f"Tendência de término: <strong>{nova_data_fmt}</strong>, sem necessidade de aditivo de prazo."
        )
    elif idp_val >= 1.0:
        analise_idp = (
            f"Com IDP = {idp_val:.2f}, a obra apresenta desempenho físico <strong>exatamente alinhado</strong> "
            f"ao cronograma contratual. A tendência de término mantém a data prevista de <strong>{nova_data_fmt}</strong>, "
            f"sem necessidade de aditivo de prazo."
        )
    elif idp_val >= 0.9:
        deficit = (1.0 - idp_val) * 100
        analise_idp = (
            f"Com IDP = {idp_val:.2f}, a obra apresenta <strong>leve atraso físico</strong> em relação ao cronograma "
            f"({deficit:.1f}% de déficit). Atenção redobrada é necessária para recuperação do ritmo. "
            f"Tendência de término: <strong>{nova_data_fmt}</strong>. Avalie a necessidade de aceleração."
        )
    elif idp_val >= 0.75:
        deficit = (1.0 - idp_val) * 100
        analise_idp = (
            f"Com IDP = {idp_val:.2f}, a obra está em <strong>atraso significativo</strong> "
            f"({deficit:.1f}% abaixo do planejado). Risco elevado de não cumprimento do prazo contratual. "
            f"Tendência de término: <strong>{nova_data_fmt}</strong>. Plano de recuperação urgente."
        )
    else:
        deficit = (1.0 - idp_val) * 100
        analise_idp = (
            f"Com IDP = {idp_val:.2f}, a obra está em situação <strong>crítica de atraso</strong> "
            f"({deficit:.1f}% abaixo do planejado). Alta probabilidade de necessidade de aditivo de prazo. "
            f"Tendência de término: <strong>{nova_data_fmt}</strong>. Intervenção imediata necessária."
        )
    idp["analise"] = analise_idp

    # ---- Etapas (nível 1 da EAP orçamento) ----------------------------
    df_orc = config["eap_orcamento"].copy()

    # pegar valor medido acumulado por etapa na medição atual
    df_med_at = df_eap[df_eap["medicao"] == med_atual].copy()
    df_med_at = df_med_at[df_med_at["item"].apply(
        lambda x: x is not None and "." not in str(x)
    )]

    # medição anterior acumulada por etapa
    df_med_pr = df_eap[df_eap["medicao"] == med_atual - 1].copy()
    df_med_pr = df_med_pr[df_med_pr["item"].apply(
        lambda x: x is not None and "." not in str(x)
    )]

    etapas = []
    for i, row in df_orc.iterrows():
        item_id = row["item"]
        nome    = row["descricao"]
        vt      = float(row["valor_rs"] or 0)
        pct_ac  = float(row["pct_concluido"] or 0)
        cor     = ETAPA_CORES[i] if i < len(ETAPA_CORES) else "gray"

        # valor medido acumulado desta etapa
        r_at = df_med_at[df_med_at["item"] == item_id]
        v_acum = float(r_at["pct_acumulado"].iloc[0]) if len(r_at) else 0.0
        vm_acum = vt * v_acum

        r_pr = df_med_pr[df_med_pr["item"] == item_id]
        v_ant = float(r_pr["pct_acumulado"].iloc[0]) if len(r_pr) else 0.0
        vm_ant = vt * v_ant

        pct_mes = float(r_at["pct_mes"].iloc[0]) if len(r_at) else 0.0

        etapas.append({
            "item":                    int(item_id) if isinstance(item_id, (int, float)) else i + 1,
            "nome":                    nome,
            "valor_total":             vt,
            "percentual_acumulado":    v_acum,
            "percentual_anterior":     v_ant,
            "percentual_mes_atual":    pct_mes,
            "valor_medido_acumulado":  vm_acum,
            "valor_medido_anterior":   vm_ant,
            "saldo_a_medir":           max(vt - vm_acum, 0),
            "cor":                     cor,
        })

    # ---- Cronograma (Curva S) — usar df_totais para medido ------
    # Config cronograma_mensal tem NaN para meses 1 e 2 (valor_medido_acum).
    # df_totais tem 0.0 correto para todos os meses medidos, inclusive os zerados.
    df_cron = config["cronograma_mensal"].copy()
    cronograma = []
    for _, r in df_cron.iterrows():
        mes_num = int(r["mes"])
        vp_acum = r["valor_planejado_acum"]
        if mes_num <= med_atual:
            row_t = df_totais[df_totais["medicao"] == mes_num]
            if len(row_t) and pd.notna(row_t["total_medido_acum"].iloc[0]):
                vm_acum_val = float(row_t["total_medido_acum"].iloc[0])
            else:
                vm_acum_val = 0.0
            pct_medido = vm_acum_val / valor_contrato if valor_contrato else 0.0
            val_medido = vm_acum_val
        else:
            pct_medido = None
            val_medido = 0.0
        cronograma.append({
            "mes_label":                      f"M{mes_num:02d}",
            "percentual_planejado_acumulado":  float(vp_acum) / valor_contrato if (valor_contrato and pd.notna(vp_acum)) else 0,
            "percentual_medido_acumulado":     pct_medido,
            "valor_medido_acumulado":          val_medido,
        })

    # ---- Contrato ------------------------------------------------------
    from openpyxl import load_workbook as _lwb  # lê metadados da EAP
    contrato = {
        "objeto":    "Construção da Nova Sede da Justiça Federal em Juazeiro do Norte/CE",
        "contratada": "Consórcio Juazeiro do Norte",
        "cnpj":      "62.009.288/0001-12",
        "periodo_inicio": vig["inicio"],
        "periodo_fim":    vig["fim"],
    }

    # ---- Fontes (para aba Dados do Contrato) ---------------------------
    fontes = {
        "avanco_fisico_acumulado":        f"EAP!col{7+(med_atual-1)*3+2} L270",
        "mes_atual_percentual":           f"EAP!col{7+(med_atual-1)*3+1} L270",
        "medicoes_anteriores_percentual": f"EAP!col{7+(med_atual-2)*3+2} L270",
        "valor_medido_acumulado":         f"EAP!col{7+(med_atual-1)*3} L271",
        "idp":                            "CONFIG!M27",
        "prazo_vigencia":                 "CONFIG!O6",
    }

    return {
        "medicoes":           medicoes,
        "idp":                idp,
        "prazos":             prazos,
        "etapas":             etapas,
        "cronograma":         cronograma,
        "valor_contrato":     valor_contrato,
        "valor_medido_acumulado": val_med_acum,
        "saldo_contratual":   saldo,
        "saldo_percentual":   saldo_pct,
        "contrato":           contrato,
        "fontes":             fontes,
        # objetos brutos (para uso avançado)
        "_config":    config,
        "_df_eap":    df_eap,
        "_df_totais": df_totais,
    }


# =============================================================================
# CONFIG
# =============================================================================
st.set_page_config(
    page_title="Dashboard de Obra — JCP Juazeiro do Norte",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

DEFAULT_EAP_FILE = DATA_DIR / "eap_atual.xlsx"
ROOT_EAP_FILE    = BASE_DIR / "EAP_MEDICAO_4_MEDICAO___OBRAS__1_ (2).xlsx"

ABAS = [
    "Visão Geral",
    "Avanço Físico",
    "Execução Financeira",
    "Prazos",
    "Dados do Contrato",
    "Upload EAP",
]


# =============================================================================
# CSS
# =============================================================================
def inject_css():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

        :root {
            --bg: #F4F2EE;
            --surface: #FFFFFF;
            --surface2: #F9F8F6;
            --border: rgba(0,0,0,0.08);
            --text: #1A1917;
            --text2: #6B6860;
            --text3: #9E9C96;
            --accent: #1B3A5C;
            --accent2: #2E6DA4;
            --green: #1D6B3E;
            --green-bg: #E8F5EE;
            --amber: #8A5300;
            --amber-bg: #FEF3DC;
            --blue-bg: #E4EDF6;
            --red: #8B1A1A;
            --red-bg: #FDEAEA;
            --radius: 12px;
            --page-x: 48px;
            --section-gap: 22px;
            --card-gap: 12px;
        }

        * { font-family: 'DM Sans', sans-serif !important; }

        #MainMenu, footer, header { visibility: hidden; }

        html, body,
        .stApp,
        [data-testid="stAppViewContainer"],
        [data-testid="stAppViewBlockContainer"],
        [data-testid="block-container"] {
            background: var(--accent) !important;
        }

        /* Faz o app ocupar altura total e esticar o conteúdo */
        html, body { height: 100% !important; }
        .stApp {
            min-height: 100vh !important;
            display: flex !important;
            flex-direction: column !important;
        }
        [data-testid="stAppViewContainer"] {
            flex: 1 !important;
            display: flex !important;
            flex-direction: column !important;
        }
        [data-testid="stAppViewBlockContainer"] {
            flex: 1 !important;
            display: flex !important;
            flex-direction: column !important;
        }
        /* O tab-panel cresce para preencher o espaço restante */
        div[data-testid="stTabs"] {
            flex: 1 !important;
            display: flex !important;
            flex-direction: column !important;
        }

        .stApp > div { padding-top: 0 !important; }

        [data-testid="stAppViewBlockContainer"],
        [data-testid="block-container"],
        .block-container {
            padding: 0 !important;
            max-width: 100% !important;
        }

        section[data-testid="stSidebar"] { display: none; }

        [data-testid="stHorizontalBlock"] {
            background: transparent !important;
            gap: var(--card-gap) !important;
        }

        /* Faz as colunas crescerem em altura */
        [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] {
            display: flex !important;
            flex-direction: column !important;
        }

        /* Os cards com classe de equalização herdam a altura total da coluna */
        .overview-equal-card,
        .fisico-equal-card,
        .prazos-equal-card,
        .contrato-equal-card {
            box-sizing: border-box;
        }

        .dash-header {
            background: var(--accent);
            color: white;
            padding: 22px 48px 0;
            margin: 0 !important;
        }

        .header-top {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding-bottom: 18px;
            border-bottom: 1px solid rgba(255,255,255,0.12);
        }

        .header-title h1 {
            font-size: 15px;
            font-weight: 600;
            color: white;
            margin: 0;
        }

        .header-title p {
            font-size: 12px;
            color: rgba(255,255,255,0.62);
            margin: 3px 0 0;
        }

        .header-meta { text-align: right; }

        .medicao-badge {
            background: rgba(255,255,255,0.15);
            border: 1px solid rgba(255,255,255,0.25);
            color: white;
            font-size: 12px;
            font-weight: 500;
            padding: 5px 14px;
            border-radius: 20px;
            display: inline-block;
            margin-bottom: 6px;
        }

        .header-meta p {
            font-size: 11px;
            color: rgba(255,255,255,0.55);
            margin: 0;
        }

        div[data-testid="stTabs"] {
            background: var(--accent) !important;
            margin: 0 !important;
            padding: 0 !important;
        }

        div[data-testid="stTabs"] div[data-baseweb="tab-list"] {
            background: var(--accent) !important;
            padding: 14px var(--page-x) !important;
            gap: 8px !important;
            border-bottom: none !important;
        }

        div[data-testid="stTabs"] button[data-baseweb="tab"] {
            background: transparent !important;
            color: rgba(255,255,255,0.62) !important;
            border: none !important;
            border-radius: 7px !important;
            padding: 9px 20px !important;
            height: auto !important;
            font-size: 13px !important;
            font-weight: 600 !important;
            letter-spacing: 0.2px !important;
            box-shadow: none !important;
            outline: none !important;
        }

        div[data-testid="stTabs"] button[data-baseweb="tab"] p {
            color: inherit !important;
            font-size: 13px !important;
            font-weight: 600 !important;
        }

        div[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"] {
            background: rgba(255,255,255,0.18) !important;
            color: #FFFFFF !important;
        }

        div[data-testid="stTabs"] div[data-baseweb="tab-highlight"],
        div[data-testid="stTabs"] [data-testid="stTabBarActiveTabHighlight"] {
            display: none !important;
        }

        div[data-testid="stTabs"] div[data-baseweb="tab-panel"] {
            background: var(--bg) !important;
            padding: var(--section-gap) var(--page-x) 48px !important;
            margin: 0 !important;
            flex: 1 !important;
            min-height: calc(100vh - 140px) !important;
        }

        div[data-testid="stTabs"] div[data-baseweb="tab-panel"] > div {
            padding: 0 !important;
        }

        .card {
            background: var(--surface);
            border-radius: var(--radius);
            border: 1px solid var(--border);
            padding: 20px;
            box-sizing: border-box;
        }

        .equal-card-row {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: var(--card-gap);
            align-items: stretch;
        }

        .equal-card-row > .card {
            height: 100%;
            min-height: 0 !important;
        }

        /* Espaçamento vertical único entre blocos do dashboard.
           Mantém o mesmo respiro da faixa azul até os KPIs. */
        div[data-testid="stVerticalBlock"] {
            gap: 0 !important;
        }

        /* Na aba Avanço Físico, as colunas precisam começar e terminar alinhadas. */
        .fisico-equal-card {
            height: 100%;
        }

        @media (max-width: 900px) {
            .equal-card-row {
                grid-template-columns: 1fr;
            }
        }

        .card-title {
            font-size: 11px;
            font-weight: 600;
            color: var(--text2);
            text-transform: uppercase;
            letter-spacing: 0.8px;
            margin-bottom: 16px;
        }

        .kpi-card {
            background: var(--surface);
            border-radius: var(--radius);
            border: 1px solid var(--border);
            padding: 18px 20px;
            height: 150px;
            box-sizing: border-box;
            display: flex;
            flex-direction: column;
            align-items: flex-start;
        }

        .kpi-label {
            font-size: 11px;
            font-weight: 500;
            color: var(--text3);
            text-transform: uppercase;
            letter-spacing: 0.8px;
            margin-bottom: 8px;
        }

        .kpi-value {
            font-size: 26px;
            font-weight: 600;
            color: var(--text);
            line-height: 1;
            font-family: 'DM Mono', monospace !important;
        }

        .kpi-value.small { font-size: 19px; }

        .kpi-sub {
            font-size: 11px;
            color: var(--text3);
            margin-top: 6px;
        }

        .kpi-sub strong {
            color: var(--text);
            font-family: 'DM Mono', monospace !important;
        }

        .mini-badge,
        .kpi-badge {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            font-size: 11px;
            font-weight: 500;
            padding: 3px 8px;
            border-radius: 12px;
            margin-top: auto;
            width: auto !important;
            max-width: fit-content;
        }

        .badge-green { background: var(--green-bg); color: var(--green); }
        .badge-blue  { background: var(--blue-bg); color: var(--accent); }
        .badge-amber { background: var(--amber-bg); color: var(--amber); }
        .badge-red   { background: var(--red-bg); color: var(--red); }

        .progress-labels {
            display: flex;
            justify-content: space-between;
            margin-bottom: 5px;
        }

        .progress-name {
            font-size: 12px;
            color: var(--text2);
        }

        .progress-pct {
            font-size: 12px;
            font-weight: 600;
            color: var(--text);
            font-family: 'DM Mono', monospace !important;
        }

        .progress-track {
            height: 8px;
            background: var(--bg);
            border-radius: 4px;
            overflow: hidden;
            margin-bottom: 12px;
        }

        .progress-bar {
            height: 100%;
            border-radius: 4px;
        }

        .bar-green { background: #2D9B63; }
        .bar-blue  { background: #2E6DA4; }
        .bar-amber { background: #D4910E; }
        .bar-gray  { background: #888780; }

        .pill-list {
            display: flex;
            flex-wrap: wrap;
            gap: 6px;
            margin-top: 6px;
        }

        .pill {
            font-size: 11px;
            padding: 3px 8px;
            border-radius: 10px;
            background: var(--surface2);
            color: var(--text3);
            border: 1px solid var(--border);
        }

        .prazo-track {
            height: 24px;
            background: var(--bg);
            border-radius: 6px;
            overflow: hidden;
            position: relative;
            margin: 10px 0;
        }

        .prazo-decorrido {
            height: 100%;
            background: var(--accent2);
            border-radius: 6px 0 0 6px;
            display: flex;
            align-items: center;
            padding-left: 8px;
        }

        .prazo-decorrido span {
            font-size: 11px;
            font-weight: 500;
            color: white;
            white-space: nowrap;
        }

        .prazo-restante-label {
            position: absolute;
            right: 8px;
            top: 50%;
            transform: translateY(-50%);
            font-size: 11px;
            color: var(--text3);
        }

        .medicoes-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 6px;
        }

        .med-card {
            text-align: center;
            padding: 8px;
            border-radius: 6px;
            font-size: 10px;
        }

        .med-card strong {
            display: block;
            font-size: 13px;
            font-family: 'DM Mono', monospace !important;
            margin: 2px 0;
        }

        .med-green {
            background: var(--green-bg);
            color: var(--green);
        }

        .med-blue-active {
            background: var(--blue-bg);
            color: var(--accent2);
            border: 2px solid var(--accent2);
        }

        .med-gray {
            background: var(--surface2);
            color: var(--text3);
            border: 1px dashed var(--border);
        }

        .etapa-row {
            display: grid;
            grid-template-columns: 28px 1fr 115px 100px 120px;
            gap: 10px;
            align-items: center;
            padding: 10px 0;
            border-bottom: 1px solid var(--border);
        }

        .etapa-row:last-child { border-bottom: none; }

        .etapa-row.header span {
            font-size: 10px;
            font-weight: 600;
            color: var(--text3);
            text-transform: uppercase;
            letter-spacing: 0.6px;
        }

        .etapa-num,
        .etapa-valor,
        .etapa-acum {
            font-size: 12px;
            font-family: 'DM Mono', monospace !important;
        }

        .etapa-num { color: var(--text3); }

        .etapa-nome {
            font-size: 13px;
            color: var(--text);
            font-weight: 500;
        }

        .etapa-valor {
            color: var(--text2);
            text-align: right;
        }

        .etapa-acum {
            color: var(--text);
            font-weight: 500;
            text-align: center;
        }

        .mini-bar-track {
            height: 5px;
            background: var(--bg);
            border-radius: 3px;
            overflow: hidden;
        }

        .mini-bar-fill {
            height: 100%;
            border-radius: 3px;
        }

        .med-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
        }

        .med-table th {
            text-align: left;
            font-size: 10px;
            font-weight: 600;
            color: var(--text3);
            text-transform: uppercase;
            letter-spacing: 0.6px;
            padding: 0 10px 10px;
            border-bottom: 1px solid var(--border);
        }

        .med-table td {
            padding: 10px;
            border-bottom: 1px solid var(--border);
            color: var(--text2);
        }

        .med-table tr:last-child td { border-bottom: none; }

        .med-table .mono { font-family: 'DM Mono', monospace !important; }

        .med-table .highlight {
            color: var(--text);
            font-weight: 600;
        }

        .med-table tr.current-row td { background: var(--blue-bg); }

        .med-table tr.current-row td:first-child {
            border-left: 3px solid var(--accent2);
            padding-left: 7px;
        }

        .info-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 8px;
        }

        .info-row {
            padding: 8px 0;
            border-bottom: 1px solid var(--border);
            display: flex;
            flex-direction: column;
            gap: 2px;
        }

        .info-label {
            font-size: 10px;
            color: var(--text3);
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }

        .info-val {
            font-size: 13px;
            color: var(--text);
            font-weight: 500;
        }

        .timeline-item {
            display: flex;
            gap: 14px;
            padding: 12px 0;
            border-bottom: 1px solid var(--border);
        }

        .timeline-item:last-child { border-bottom: none; }

        .tl-dot {
            width: 10px;
            height: 10px;
            border-radius: 50%;
            margin-top: 4px;
            flex-shrink: 0;
        }

        .tl-dot-green { background: #2D9B63; }
        .tl-dot-gray  { background: var(--text3); }
        .tl-dot-blue  { background: #2E6DA4; }

        .tl-content { flex: 1; }

        .tl-title {
            font-size: 13px;
            font-weight: 500;
            color: var(--text);
        }

        .tl-meta {
            font-size: 11px;
            color: var(--text3);
            margin-top: 2px;
        }

        .tl-val {
            font-size: 13px;
            font-family: 'DM Mono', monospace !important;
            font-weight: 500;
            color: var(--text2);
            text-align: right;
            flex-shrink: 0;
        }

        .text-green { color: var(--green) !important; }
        .text-blue  { color: var(--accent2) !important; }

        .status-box {
            background: var(--green-bg);
            border-radius: 8px;
            padding: 14px;
            text-align: center;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            min-height: 120px;
        }

        .status-box .label {
            font-size: 10px;
            color: var(--green);
            text-transform: uppercase;
            letter-spacing: 0.6px;
            margin-bottom: 6px;
        }

        .status-box .value {
            font-size: 36px;
            font-weight: 300;
            color: var(--green);
            font-family: 'DM Mono', monospace !important;
            line-height: 1;
        }

        .status-box .sub {
            font-size: 11px;
            color: var(--green);
            margin-top: 6px;
        }

        .small-box {
            background: var(--surface2);
            border-radius: 8px;
            padding: 10px 14px;
        }

        .small-box .label {
            font-size: 10px;
            color: var(--text3);
            text-transform: uppercase;
        }

        .small-box .value {
            font-size: 16px;
            font-weight: 600;
            color: var(--text);
            margin-top: 2px;
            font-family: 'DM Mono', monospace !important;
        }

        .small-box .sub {
            font-size: 10px;
            color: var(--text3);
        }

        [data-testid="stPlotlyChart"] {
            background: var(--surface) !important;
            border: 1px solid var(--border) !important;
            border-radius: var(--radius) !important;
            overflow: hidden !important;
        }

        /* Login: botão Entrar na cor do cabeçalho */
        [data-testid="stBaseButton-primary"] {
            background-color: var(--accent) !important;
            border-color: var(--accent) !important;
            color: white !important;
        }
        [data-testid="stBaseButton-primary"]:hover {
            background-color: #152e4a !important;
            border-color: #152e4a !important;
        }

        /* Inputs e selects com fundo branco */
        [data-testid="stSelectbox"] > div > div,
        [data-baseweb="select"] > div,
        [data-testid="stTextInput"] input,
        [data-testid="stNumberInput"] input,
        [data-testid="stTextArea"] textarea {
            background-color: white !important;
        }

        .watermark {
            text-align: center;
            margin-top: 32px;
            font-size: 11px;
            color: var(--text3);
            padding-top: 16px;
            border-top: 1px solid var(--border);
        }

        /* Data editor: alinhar todas as celulas a esquerda */
        [data-testid="stDataEditor"] .dvn-scroller .gdg-cell,
        [data-testid="stDataEditor"] .dvn-scroller [role="gridcell"],
        [data-testid="stDataEditor"] .dvn-scroller [role="columnheader"] {
            text-align: left !important;
            justify-content: flex-start !important;
        }
        [data-testid="stDataEditor"] span,
        [data-testid="stDataEditor"] div[style*="text-align: right"],
        [data-testid="stDataEditor"] div[style*="text-align:right"] {
            text-align: left !important;
            justify-content: flex-start !important;
        }
        </style>

        <script>
        (function() {
            var EQ_GROUPS = [];

            function equalizeGroup(cls) {
                var cards = Array.from(document.querySelectorAll('.' + cls));
                if (cards.length < 2) return;
                // Reset
                cards.forEach(function(c) { c.style.minHeight = ''; });
                // Reflow: force layout before measuring
                void cards[0].offsetHeight;
                var max = 0;
                cards.forEach(function(c) {
                    var h = c.getBoundingClientRect().height;
                    if (h > max) max = h;
                });
                if (max > 10) {
                    cards.forEach(function(c) { c.style.minHeight = max + 'px'; });
                }
            }

            function equalizeAll() {
                EQ_GROUPS.forEach(equalizeGroup);
            }

            // Run on load at multiple intervals to catch async Streamlit renders
            [100, 300, 700, 1500, 3000, 5000].forEach(function(t) {
                setTimeout(equalizeAll, t);
            });

            // Re-run on window resize/zoom
            var resizeTimer;
            window.addEventListener('resize', function() {
                clearTimeout(resizeTimer);
                resizeTimer = setTimeout(function() {
                    // Reset all first, then re-equalize
                    EQ_GROUPS.forEach(function(cls) {
                        Array.from(document.querySelectorAll('.' + cls))
                             .forEach(function(c) { c.style.minHeight = ''; });
                    });
                    equalizeAll();
                }, 150);
            });

            // Re-run whenever DOM changes (tab switches, re-renders)
            if (window.MutationObserver) {
                var mutTimer;
                var obs = new MutationObserver(function() {
                    clearTimeout(mutTimer);
                    mutTimer = setTimeout(equalizeAll, 200);
                });
                obs.observe(document.body, { childList: true, subtree: true });
            }
        })();
        </script>
        """,
        unsafe_allow_html=True,
    )


# =============================================================================
# CARREGAMENTO
# =============================================================================

@st.cache_data(show_spinner=False)
def load_from_path(path: str) -> dict:
    config, df_eap, df_totais = carregar_dados(path)
    return _build_data(config, df_eap, df_totais)


@st.cache_data(show_spinner=False)
def load_from_bytes(file_bytes: bytes) -> dict:
    import tempfile, os
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        config, df_eap, df_totais = carregar_dados(tmp_path)
        return _build_data(config, df_eap, df_totais)
    finally:
        os.unlink(tmp_path)


def load_current_data() -> dict:
    if "arquivo_eap_bytes" in st.session_state:
        return load_from_bytes(st.session_state["arquivo_eap_bytes"])
    if DEFAULT_EAP_FILE.exists():
        return load_from_path(str(DEFAULT_EAP_FILE))
    if ROOT_EAP_FILE.exists():
        return load_from_path(str(ROOT_EAP_FILE))
    raise FileNotFoundError("Nenhum arquivo EAP encontrado.")


# =============================================================================
# HTML HELPERS
# =============================================================================

def html(markup: str):
    st.html(markup)


def spacer(height: int = 22):
    html(f"<div style='height:{height}px'></div>")


def card(title: str, body: str, extra_class: str = "", min_height: int = 0):
    mh = f"min-height:{min_height}px;" if min_height else ""
    html(
        f"""
        <div class="card {extra_class}" style="{mh}">
            <div class="card-title">{title}</div>
            {body}
        </div>
        """
    )

def card_markup(title: str, body: str, extra_class: str = "") -> str:
    return f"""
    <div class="card {extra_class}">
        <div class="card-title">{title}</div>
        {body}
    </div>
    """


def card_pair(left: str, right: str):
    html(f"""
    <div class="equal-card-row">
        {left}
        {right}
    </div>
    """)


def render_header(data: dict):
    contrato = data["contrato"]
    med = data["medicoes"]["medicao_atual"]

    html(
        f"""
        <div class="dash-header">
            <div class="header-top">
                <div class="header-title">
                    <h1>Construção — Nova Sede da Justiça Federal</h1>
                    <p>Juazeiro do Norte / CE &nbsp;·&nbsp; {contrato["contratada"]} &nbsp;·&nbsp; CNPJ {contrato["cnpj"]}</p>
                </div>
                <div class="header-meta">
                    <div class="medicao-badge">{med}ª Medição</div>
                    <p>Dados extraídos das abas CONFIG e EAP DE MEDIÇÃO</p>
                </div>
            </div>
        </div>
        """
    )


def kpi(label: str, value: str, subtitle: str, badge: str, badge_type: str = "blue", small: bool = False):
    value_class = "kpi-value small" if small else "kpi-value"
    html(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="{value_class}">{value}</div>
            <div class="kpi-sub">{subtitle}</div>
            <div class="kpi-badge badge-{badge_type}">{badge}</div>
        </div>
        """
    )


def progress_item(label: str, pct: float, color: str) -> str:
    pct_num = as_percent(pct)
    return f"""
    <div>
        <div class="progress-labels">
            <span class="progress-name">{label}</span>
            <span class="progress-pct">{fmt_percent(pct)}</span>
        </div>
        <div class="progress-track">
            <div class="progress-bar bar-{color}" style="width:{pct_num}%"></div>
        </div>
    </div>
    """


def watermark():
    html(
        """
        <div class="watermark">
            Dashboard de Acompanhamento de Obra · Dados extraídos de CONFIG e EAP DE MEDIÇÃO
        </div>
        """
    )


def _calc_overview_height(data: dict) -> int:
    """
    Calcula altura mínima para igualar os dois cards da Visão Geral.
    Card esquerdo (Etapas Iniciadas): título(44) + itens iniciados * 38px
                                       + seção não iniciadas(50) + pills
    Card direito (Situação Geral): título(44) + IDP+boxes(160) + barra(60)
                                   + medições-grid(110) + separadores(20) = ~394px fixo
    """
    etapas_ini  = [e for e in data["etapas"] if as_percent(e["percentual_acumulado"]) > 0]
    etapas_zero = [e for e in data["etapas"] if as_percent(e["percentual_acumulado"]) == 0]
    # estimativa do card esquerdo
    altura_esq = 44 + len(etapas_ini) * 38 + 50 + max(1, (len(etapas_zero) // 3 + 1)) * 30 + 20
    # card direito fixo ~394px
    altura_dir = 394
    return max(altura_esq, altura_dir)


def _calc_fisico_height(data: dict) -> int:
    """
    Altura do card da esquerda na aba Avanço Físico.

    A referência correta é a soma visual da coluna direita:
    gráfico comparativo + espaçamento de seção + gráfico por etapa.
    Antes havia uma folga extra que deixava o card "Avanço por Etapa"
    mais alto que os dois gráficos da direita.
    """
    n = len(data["etapas"])
    altura_comparativo = 255
    altura_barras = max(260, n * 46 + 80)
    gap_entre_graficos = 22

    total_direita = altura_comparativo + gap_entre_graficos + altura_barras

    # Altura mínima natural da tabela, caso a tabela cresça em outra planilha.
    altura_tabela = 44 + 40 + n * 41 + 50 + 20

    return max(total_direita, altura_tabela)


def equalize_heights(class_name: str):
    """Injeta JS que iguala alturas de todos os cards com a classe dada."""
    import streamlit.components.v1 as components
    components.html(
        f"""
        <script>
        (function() {{
            function eq() {{
                // sobe até o shadow host do iframe
                const host = window.frameElement;
                if (!host) return;
                const root = host.closest('[data-testid="stAppViewBlockContainer"]') || document.body;
                const cards = Array.from((root || document).querySelectorAll('.{class_name}'));
                if (cards.length < 2) return;
                // reset
                cards.forEach(c => c.style.minHeight = '');
                const max = Math.max(...cards.map(c => c.getBoundingClientRect().height));
                cards.forEach(c => c.style.minHeight = max + 'px');
            }}
            // tenta algumas vezes pois o Streamlit renderiza em etapas
            [100, 400, 900, 1800].forEach(t => setTimeout(eq, t));
        }})();
        </script>
        """,
        height=0,
    )


# =============================================================================
# GRÁFICOS
# =============================================================================

def base_layout(title: str, height: int, extra: dict | None = None):
    cfg = dict(
        title=dict(
            text=title,
            font=dict(family="DM Sans", size=11, color="#6B6860"),
            x=0, xanchor="left",
            pad=dict(l=12, t=12),
        ),
        height=height,
        margin=dict(l=12, r=12, t=48, b=40),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="DM Sans", size=10, color="#9E9C96"),
    )
    if extra:
        cfg.update(extra)
    return cfg


def plot_curva_s(data: dict):
    schedule = data["cronograma"]
    x        = [r["mes_label"] for r in schedule]
    planned  = [as_percent(r["percentual_planejado_acumulado"]) for r in schedule]
    measured = [
        as_percent(r["percentual_medido_acumulado"])
        if r["percentual_medido_acumulado"] is not None
        else None
        for r in schedule
    ]

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=x, y=planned, name="Planejado",
        line=dict(color="#2E6DA4", width=2, dash="dot"),
        fill="tozeroy", fillcolor="rgba(46,109,164,0.08)",
    ))
    fig.add_trace(go.Scatter(
        x=x, y=measured, name="Realizado",
        line=dict(color="#2D9B63", width=2.5),
        fill="tozeroy", fillcolor="rgba(45,155,99,0.12)",
        mode="lines+markers", marker=dict(color="#2D9B63", size=5),
    ))
    fig.update_layout(**base_layout(
        "AVANÇO FÍSICO ACUMULADO (REALIZADO VS. PLANEJADO)", 300,
        dict(
            yaxis=dict(ticksuffix="%", range=[0, 100], dtick=20, gridcolor="rgba(0,0,0,0.05)"),
            xaxis=dict(showgrid=False, nticks=12),
            legend=dict(orientation="h", y=-0.15, x=0.5, xanchor="center",
                        font=dict(size=12, color="#6B6860")),
            showlegend=True,
        ),
    ))
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def plot_medicoes(data: dict):
    """Comparativo Realizado vs. Planejado — uma barra por medição."""
    df_totais = data["_df_totais"]
    df_cron   = data["_config"]["cronograma_mensal"]
    med_atual = data["medicoes"]["medicao_atual"]
    vc        = data["valor_contrato"]

    MESES_PT = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
    import datetime as _dt

    labels, realizado, planejado = [], [], []

    for i in range(1, med_atual + 1):
        row_t = df_totais[df_totais["medicao"] == i]
        pct_r = float(row_t["total_pct_mes"].iloc[0]) if (len(row_t) and pd.notna(row_t["total_pct_mes"].iloc[0])) else 0.0

        row_c = df_cron[df_cron["mes"] == i]
        pct_p = 0.0
        if len(row_c) and vc and pd.notna(row_c["valor_planejado"].iloc[0]):
            pct_p = float(row_c["valor_planejado"].iloc[0]) / vc

        # label com mês/ano
        if len(row_t) and row_t["data_ref"].iloc[0] and isinstance(row_t["data_ref"].iloc[0], _dt.date):
            d = row_t["data_ref"].iloc[0]
            lbl = f"Med. {i} {MESES_PT[d.month-1]}/{str(d.year)[2:]}"
        else:
            lbl = f"Med. {i}"

        labels.append(lbl)
        realizado.append(as_percent(pct_r))
        planejado.append(as_percent(pct_p))

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Planejado", x=labels, y=planejado,
        marker_color="rgba(46,109,164,0.55)",
        text=[f"{v:.2f}%".replace(".", ",") for v in planejado],
        textposition="outside", textfont=dict(size=9, color="#6B6860"),
    ))
    fig.add_trace(go.Bar(
        name="Realizado", x=labels, y=realizado,
        marker_color="rgba(45,155,99,0.80)",
        text=[f"{v:.2f}%".replace(".", ",") for v in realizado],
        textposition="outside", textfont=dict(size=9, color="#6B6860"),
    ))
    fig.update_layout(**base_layout(
        "COMPARATIVO REALIZADO VS. PLANEJADO", 255,
        dict(
            barmode="group",
            yaxis=dict(ticksuffix="%", gridcolor="rgba(0,0,0,0.05)"),
            xaxis=dict(showgrid=False),
            legend=dict(orientation="h", y=-0.18, x=0.5, xanchor="center",
                        font=dict(size=11, color="#6B6860")),
            showlegend=True,
        ),
    ))
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def plot_barras_etapas(data: dict):
    """Barras horizontais empilhadas 100%: % realizado (verde) vs % restante (vermelho claro) — ordenadas por % realizado."""
    etapas = data["etapas"]
    COR_REAL = "#2D9B63"          # mesmo verde do comparativo realizado vs. planejado
    COR_REST = "rgba(220,80,80,0.15)"   # vermelho claro

    if not etapas:
        return

    # Ordenar por % realizado crescente (menor no topo, maior na base)
    etapas_sorted = sorted(etapas, key=lambda e: as_percent(e["percentual_acumulado"]))

    nomes    = [e["nome"] for e in etapas_sorted]
    pct_real = [as_percent(e["percentual_acumulado"]) for e in etapas_sorted]
    pct_rest = [max(100.0 - as_percent(e["percentual_acumulado"]), 0) for e in etapas_sorted]

    # Quebrar nomes longos a cada 28 caracteres usando <br>
    def _wrap_name(n, max_chars=28):
        if len(n) <= max_chars:
            return n
        words = n.split()
        lines, cur = [], ""
        for w in words:
            if cur and len(cur) + 1 + len(w) > max_chars:
                lines.append(cur)
                cur = w
            else:
                cur = (cur + " " + w).strip()
        if cur:
            lines.append(cur)
        return "<br>".join(lines)

    nomes_curtos = [_wrap_name(n) for n in nomes]

    n      = len(etapas_sorted)
    # Altura por barra varia conforme numero de linhas do rotulo
    ALTURA_BASE_BAR = 40   # altura de uma barra de 1 linha
    ALTURA_EXTRA    = 14   # pixels extras por linha adicional
    altura_total = sum(
        ALTURA_BASE_BAR + max(0, nc.count("<br>")) * ALTURA_EXTRA
        for nc in nomes_curtos
    )
    altura = max(260, altura_total + 80)

    fig = go.Figure()

    # Threshold: barras com % realizado abaixo disso recebem anotacao externa
    THRESHOLD_INSIDE  = 8.0
    THRESHOLD_OUTSIDE = 0.1  # nao exibe se <= 0.1%

    # Separa quais barras recebem texto interno vs. anotacao externa
    text_real_inside = []   # texto dentro da barra (barras grandes)
    annots_real = []        # anotacoes para barras pequenas

    for i, (v, nome) in enumerate(zip(pct_real, nomes_curtos)):
        if v <= THRESHOLD_OUTSIDE:
            text_real_inside.append("")
        elif v < THRESHOLD_INSIDE:
            text_real_inside.append("")
            annots_real.append(dict(
                x=v + 1.0,
                y=nome,
                text=f"<b>{v:.2f}%</b>",
                xanchor="left",
                yanchor="middle",
                showarrow=False,
                font=dict(size=10, color="#1A6B40"),
                xref="x", yref="y",
            ))
        else:
            text_real_inside.append(f"{v:.2f}%")

    text_rest  = []
    tpos_rest  = []
    for v in pct_rest:
        if v <= THRESHOLD_OUTSIDE:
            text_rest.append("")
            tpos_rest.append("outside")
        elif v < THRESHOLD_INSIDE:
            text_rest.append(f"{v:.2f}%")
            tpos_rest.append("outside")
        else:
            text_rest.append(f"{v:.2f}%")
            tpos_rest.append("inside")

    fig.add_trace(go.Bar(
        name="Realizado",
        y=nomes_curtos,
        x=pct_real,
        orientation="h",
        marker_color=COR_REAL,
        text=text_real_inside,
        textposition="inside",
        insidetextanchor="start",
        textangle=0,
        textfont=dict(size=10, color="white"),
        hovertemplate="%{y}<br>Realizado: %{x:.2f}%<extra></extra>",
        cliponaxis=False,
    ))

    fig.add_trace(go.Bar(
        name="A realizar",
        y=nomes_curtos,
        x=pct_rest,
        orientation="h",
        marker_color=COR_REST,
        marker_line=dict(width=0),
        text=text_rest,
        textposition=tpos_rest,
        insidetextanchor="end",
        textangle=0,
        textfont=dict(size=9, color="#C04040"),
        outsidetextfont=dict(size=9, color="#C04040"),
        hovertemplate="%{y}<br>A realizar: %{x:.2f}%<extra></extra>",
        showlegend=True,
        cliponaxis=False,
    ))

    fig.update_layout(**base_layout(
        "AVANÇO FÍSICO POR ETAPA (% REALIZADO)", altura,
        dict(
            barmode="stack",
            xaxis=dict(
                range=[0, 112], ticksuffix="%", dtick=25,
                gridcolor="rgba(0,0,0,0.05)", showgrid=True,
                fixedrange=True,
            ),
            yaxis=dict(showgrid=False, automargin=True, tickfont=dict(size=11)),
            legend=dict(orientation="h", y=-0.08, x=0.5, xanchor="center",
                        font=dict(size=11, color="#6B6860")),
            showlegend=True,
            margin=dict(l=12, r=70, t=48, b=50),
            uniformtext=dict(mode=None),
            annotations=annots_real,
        ),
    ))
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def _plot_barras_etapas_valor_UNUSED(data: dict):
    """Barras horizontais: valor contratual por etapa (só etapas com valor > 0)."""
    etapas = data["etapas"]
    COLORS = ["#2D9B63","#2E6DA4","#D4910E","#B05CC0",
              "#B0ADA6","#888780","#B0ADA6","#888780","#B0ADA6","#B0ADA6"]

    # Filtrar só etapas com valor contratual
    etapas_com_valor = [e for e in etapas if e["valor_total"] > 0]
    if not etapas_com_valor:
        return

    nomes  = [e["nome"] for e in etapas_com_valor]
    totais = [e["valor_total"] for e in etapas_com_valor]
    medido = [e["valor_medido_acumulado"] for e in etapas_com_valor]
    cores  = [COLORS[etapas.index(e) % len(COLORS)] for e in etapas_com_valor]

    fig = go.Figure()

    # Faixa do valor total (background)
    fig.add_trace(go.Bar(
        name="Valor Contratual",
        y=nomes, x=totais,
        orientation="h",
        marker_color="rgba(0,0,0,0.07)",
        hovertemplate="%{y}: R$ %{x:,.0f}<extra>Contratual</extra>",
    ))
    # Valor medido acumulado (foreground)
    fig.add_trace(go.Bar(
        name="Medido Acumulado",
        y=nomes, x=medido,
        orientation="h",
        marker_color=cores,
        hovertemplate="%{y}: R$ %{x:,.0f}<extra>Medido</extra>",
        text=[
            f'{as_percent(e["percentual_acumulado"]):.1f}%'
            if e["percentual_acumulado"] > 0 else ""
            for e in etapas_com_valor
        ],
        textposition="inside",
        insidetextanchor="start",
        textfont=dict(size=10, color="white"),
    ))

    n = len(etapas_com_valor)
    altura = max(220, n * 44 + 80)

    fig.update_layout(**base_layout(
        "DISTRIBUIÇÃO POR ETAPA (VALOR CONTRATUAL)", altura,
        dict(
            barmode="overlay",
            xaxis=dict(tickprefix="R$ ", tickformat=".2s", gridcolor="rgba(0,0,0,0.05)", showgrid=True),
            yaxis=dict(showgrid=False, automargin=True, tickfont=dict(size=11)),
            legend=dict(orientation="h", y=-0.12, x=0.5, xanchor="center",
                        font=dict(size=11, color="#6B6860")),
            showlegend=True,
            margin=dict(l=12, r=12, t=48, b=60),
        ),
    ))
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def plot_financeiro(data: dict):
    """Evolução financeira acumulada — soma dos valores medidos por medição."""
    df_totais = data["_df_totais"]
    med_atual = data["medicoes"]["medicao_atual"]
    MESES_PT  = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
    import datetime as _dt

    labels, acumulado = [], []
    for i in range(1, med_atual + 1):
        row_t = df_totais[df_totais["medicao"] == i]
        v_ac  = float(row_t["total_medido_acum"].iloc[0]) if (len(row_t) and pd.notna(row_t["total_medido_acum"].iloc[0])) else 0.0
        if len(row_t) and row_t["data_ref"].iloc[0] and isinstance(row_t["data_ref"].iloc[0], _dt.date):
            d   = row_t["data_ref"].iloc[0]
            lbl = f"Med. {i}<br>{MESES_PT[d.month-1]}/{str(d.year)[2:]}"
        else:
            lbl = f"Med. {i}"
        labels.append(lbl)
        acumulado.append(v_ac)

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=labels, y=acumulado,
        name="Acumulado medido",
        line=dict(color="#2D9B63", width=2.5),
        fill="tozeroy", fillcolor="rgba(45,155,99,0.10)",
        mode="lines+markers", marker=dict(color="#2D9B63", size=7),
        text=[fmt_money(v) for v in acumulado],
        hovertemplate="%{x}: %{text}<extra></extra>",
    ))
    fig.update_layout(**base_layout(
        "EVOLUÇÃO FINANCEIRA ACUMULADA", 320,
        dict(
            yaxis=dict(tickprefix="R$ ", tickformat=".2s", gridcolor="rgba(0,0,0,0.05)"),
            xaxis=dict(showgrid=False),
            showlegend=False,
        ),
    ))
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


# =============================================================================
# COMPONENTES
# =============================================================================

def _get_proj_saldo_info() -> dict:
    """
    Recalcula dinamicamente montante_total, saldo_devedor e pct_restante
    a partir do montante salvo nos registros e dos pagamentos 'Pago' salvos.
    Nao depende do campo saldo_devedor gravado (que so e atualizado ao salvar Resumo).
    """
    df = _carregar_registros_projetos()
    if len(df) == 0:
        return {"montante": None, "saldo": None, "pct_restante": None}
    df_num = df.copy()
    df_num["montante_total"] = pd.to_numeric(df_num["montante_total"], errors="coerce")
    df_valid = df_num[df_num["montante_total"].notna() & (df_num["montante_total"] > 0)]
    if len(df_valid) == 0:
        return {"montante": None, "saldo": None, "pct_restante": None}
    montante = float(df_valid["montante_total"].max())

    # Recalcula total pago somando TODOS os pagamentos com status "Pago"
    df_pags = _carregar_pagamentos_projetos()
    total_pago = 0.0
    if len(df_pags) > 0 and "valor" in df_pags.columns:
        df_p = df_pags.copy()
        df_p["valor_num"] = pd.to_numeric(df_p["valor"], errors="coerce").fillna(0)
        total_pago = float(df_p.loc[df_p["status"] == "Pago", "valor_num"].sum())

    saldo    = max(montante - total_pago, 0.0)
    pct_rest = saldo / montante if montante else 0
    return {"montante": montante, "saldo": saldo, "pct_restante": pct_rest}


def render_kpis(data: dict):
    med = data["medicoes"]
    idp = data["idp"]
    cols = st.columns(3, gap="small")

    with cols[0]:
        kpi(
            "Avanço Físico Acumulado",
            fmt_percent(med["avanco_fisico_acumulado"]),
            f'Mês atual: <strong>{fmt_percent(med["mes_atual_percentual"])}</strong><br>'
            f'Medições anteriores: <strong>{fmt_percent(med["medicoes_anteriores_percentual"])}</strong>',
            f'✓ {idp["status"]} (IDP = {fmt_decimal(idp["valor"])})',
            "green",
        )

    with cols[1]:
        kpi(
            "Valor Medido Acumulado",
            fmt_money(med["valor_medido_acumulado"]),
            f'{fmt_percent(med["avanco_fisico_acumulado"])} do contrato executado',
            f'{med["medicao_atual"]}ª Medição',
            "blue", small=True,
        )

    with cols[2]:
        # ── KPI unificado: Saldo Contratual (Obras + Projetos) ──
        proj_info = _get_proj_saldo_info()

        # Obras
        v_obras      = data["valor_contrato"]
        saldo_obras  = data["saldo_contratual"]
        pct_obras    = data["saldo_percentual"]

        # Projetos
        if proj_info["saldo"] is not None:
            v_proj       = proj_info["montante"]
            saldo_proj   = proj_info["saldo"]
            pct_proj     = proj_info["pct_restante"]
            proj_ok      = True
        else:
            v_proj = saldo_proj = pct_proj = None
            proj_ok = False

        # Total (obras + projetos)
        if proj_ok:
            v_total     = v_obras + v_proj
            saldo_total = saldo_obras + saldo_proj
            pct_total   = saldo_total / v_total if v_total else 0
        else:
            v_total     = v_obras
            saldo_total = saldo_obras
            pct_total   = pct_obras

        proj_saldo_html = (
            f'<span style="font-family:\'DM Mono\',monospace;">{fmt_money(saldo_proj)}</span>'
            f'<span style="color:var(--text3);font-size:10px;margin-left:4px;">({fmt_percent(pct_proj)} restante)</span>'
        ) if proj_ok else (
            '<span style="color:var(--text3);font-size:11px;">Sem dados — preencha na aba Upload EAP</span>'
        )

        _proj_col = (
            f'<div style="font-size:13px;font-weight:600;color:var(--text);font-family:\'DM Mono\',monospace;line-height:1.2;">{fmt_money(saldo_proj)}</div>'
            f'<div style="font-size:9px;color:var(--text3);margin-top:2px;">de {fmt_money(v_proj)}</div>'
            f'<div style="display:inline-flex;font-size:9px;font-weight:500;padding:2px 6px;border-radius:10px;margin-top:4px;background:var(--blue-bg);color:var(--accent2);">{fmt_percent(pct_proj)} restante</div>'
        ) if proj_ok else (
            '<div style="font-size:10px;color:var(--text3);margin-top:4px;">Sem dados</div>'
            '<div style="font-size:9px;color:var(--text3);">Preencha na aba Upload</div>'
        )

        html(f"""
        <div class="kpi-card" style="height:150px;box-sizing:border-box;overflow:hidden;">
            <div class="kpi-label">Saldo Contratual</div>
            <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:0;width:100%;border-top:1px solid var(--border);padding-top:8px;margin-top:2px;">

                <div style="padding-right:10px;border-right:1px solid var(--border);">
                    <div style="font-size:8px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:3px;">Total (Obras+Proj.)</div>
                    <div style="font-size:13px;font-weight:600;color:var(--text);font-family:'DM Mono',monospace;line-height:1.2;">{fmt_money(saldo_total)}</div>
                    <div style="font-size:9px;color:var(--text3);margin-top:2px;">de {fmt_money(v_total)}</div>
                    <div style="display:inline-flex;font-size:9px;font-weight:500;padding:2px 6px;border-radius:10px;margin-top:4px;background:var(--blue-bg);color:var(--accent2);">{fmt_percent(pct_total)} restante</div>
                </div>

                <div style="padding:0 10px;border-right:1px solid var(--border);">
                    <div style="font-size:8px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:3px;">Obras</div>
                    <div style="font-size:13px;font-weight:600;color:var(--text);font-family:'DM Mono',monospace;line-height:1.2;">{fmt_money(saldo_obras)}</div>
                    <div style="font-size:9px;color:var(--text3);margin-top:2px;">de {fmt_money(v_obras)}</div>
                    <div style="display:inline-flex;font-size:9px;font-weight:500;padding:2px 6px;border-radius:10px;margin-top:4px;background:var(--blue-bg);color:var(--accent2);">{fmt_percent(pct_obras)} restante</div>
                </div>

                <div style="padding-left:10px;">
                    <div style="font-size:8px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:3px;">Projetos</div>
                    {_proj_col}
                </div>

            </div>
        </div>
        """)


def render_avanco_etapas(data: dict, return_markup: bool = False):
    etapas_ini  = [e for e in data["etapas"] if as_percent(e["percentual_acumulado"]) > 0]
    etapas_zero = [e for e in data["etapas"] if as_percent(e["percentual_acumulado"]) == 0]

    body = "".join(
        progress_item(e["nome"], e["percentual_acumulado"], e["cor"])
        for e in etapas_ini
    )
    body += """
    <div style="margin-top:16px;padding-top:14px;border-top:1px solid var(--border);">
        <div style="font-size:11px;color:var(--text3);margin-bottom:6px;text-transform:uppercase;letter-spacing:0.6px;">
            Etapas não iniciadas
        </div>
        <div class="pill-list">
    """
    body += "".join(f'<span class="pill">{e["nome"]}</span>' for e in etapas_zero)
    body += "</div></div>"

    markup = card_markup("Avanço Físico — Etapas Iniciadas", body, "overview-equal-card")
    if return_markup:
        return markup
    html(markup)


def render_situacao_obra(data: dict, return_markup: bool = False):
    prazos = data["prazos"]
    idp    = data["idp"]
    med    = data["medicoes"]

    dec   = prazos["execucao_decorrida_dias"]
    rest  = prazos["execucao_restante_dias"]
    total = prazos["execucao_total_dias"]
    pct   = dec / total if total else 0

    # Próximas medições: calcular mês a partir da data_ref da medição atual
    df_t      = data["_df_totais"]
    prox_num  = med["medicao_atual"] + 1
    fut_num   = med["medicao_atual"] + 2
    import datetime as _dt
    MESES_PT  = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
    prox_label = "A definir"
    fut_label  = "A definir"
    row_ref = df_t[df_t["medicao"] == med["medicao_atual"]]
    if len(row_ref) and row_ref["data_ref"].iloc[0]:
        ref = row_ref["data_ref"].iloc[0]
        if isinstance(ref, _dt.date):
            pm, py = (ref.month % 12) + 1, ref.year + (1 if ref.month == 12 else 0)
            fm, fy = ((ref.month + 1) % 12) + 1, ref.year + (1 if ref.month >= 11 else 0)
            prox_label = f"{MESES_PT[pm-1]}/{str(py)[2:]}"
            fut_label  = f"{MESES_PT[fm-1]}/{str(fy)[2:]}"

    body = f"""
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:16px;">
        <div class="status-box">
            <div class="label">IDP</div>
            <div class="value">{fmt_decimal(idp["valor"])}</div>
            <div class="sub">{idp["status"]}</div>
        </div>
        <div style="display:flex;flex-direction:column;gap:8px;">
            <div class="small-box">
                <div class="label">Execução decorrida</div>
                <div class="value">{dec} dias</div>
                <div class="sub">de {total} dias totais</div>
            </div>
            <div class="small-box">
                <div class="label">Execução restante</div>
                <div class="value">{rest} dias</div>
                <div class="sub">até {prazos["prazo_execucao_data"]}</div>
            </div>
        </div>
    </div>

    <div style="font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:0.6px;margin-bottom:8px;">
        Progresso temporal — execução
    </div>
    <div class="prazo-track">
        <div class="prazo-decorrido" style="width:{pct * 100:.1f}%">
            <span>{fmt_percent(pct)} ({dec} d)</span>
        </div>
        <span class="prazo-restante-label">{rest} dias restantes</span>
    </div>

    <div style="margin-top:14px;padding-top:14px;border-top:1px solid var(--border);">
        <div style="font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:0.6px;margin-bottom:8px;">
            Medições realizadas
        </div>
        <div class="medicoes-grid">
            <div class="med-card med-green">
                <div>Anterior</div>
                <strong>{fmt_percent(med["medicoes_anteriores_percentual"])}</strong>
                <div>Até med. {med["medicao_atual"] - 1}</div>
            </div>
            <div class="med-card med-blue-active">
                <div>Med. {med["medicao_atual"]} ▶</div>
                <strong>{fmt_percent(med["mes_atual_percentual"])}</strong>
                <div>Atual</div>
            </div>
            <div class="med-card med-gray">
                <div>Med. {prox_num}</div>
                <strong>—</strong>
                <div>{prox_label}</div>
            </div>
            <div class="med-card med-gray">
                <div>Med. {fut_num}</div>
                <strong>—</strong>
                <div>{fut_label}</div>
            </div>
        </div>
    </div>
    """
    markup = card_markup("Situação Geral da Obra", body, "overview-equal-card")
    if return_markup:
        return markup
    html(markup)


def render_etapas_table(data: dict, min_h: int = 0):
    """
    Tabela de etapas com colunas: Nº | Etapa | Valor Total | Medido Acum. | Saldo a Medir | Progresso | Risco
    """
    vc = data["valor_contrato"]
 
    def nivel_risco_etapa(pct_ac, saldo, valor_contrato):
        pct_saldo = saldo / valor_contrato if valor_contrato else 0
        if pct_ac < 0.05 and pct_saldo > 0.1:
            return "Alto", "var(--red)", "var(--red-bg)"
        elif pct_ac < 0.30 and pct_saldo > 0.05:
            return "Médio", "var(--amber)", "var(--amber-bg)"
        elif pct_ac >= 0.90:
            return "Baixo", "var(--green)", "var(--green-bg)"
        else:
            return "Normal", "var(--accent2)", "var(--blue-bg)"
 
    header = """
    <div style="display:grid;grid-template-columns:28px 1fr 115px 115px 115px 72px;
                gap:8px;padding:0 0 8px;border-bottom:2px solid var(--border);margin-bottom:2px;">
        <span style="font-size:10px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:0.6px;"></span>
        <span style="font-size:10px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:0.6px;">Etapa</span>
        <span style="font-size:10px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:0.6px;text-align:right;">Valor Total</span>
        <span style="font-size:10px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:0.6px;text-align:right;">Medido Acum.</span>
        <span style="font-size:10px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:0.6px;text-align:right;">Saldo a Medir</span>
        <span style="font-size:10px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:0.6px;text-align:center;">Risco</span>
    </div>
    """
 
    rows_html = ""
    for e in data["etapas"]:
        pct_ac  = float(e["percentual_acumulado"] or 0)
        vt      = float(e["valor_total"] or 0)
        vm_acum = float(e["valor_medido_acumulado"] or 0)
        saldo   = float(e["saldo_a_medir"] or 0)
        risco_txt, risco_cor, risco_bg = nivel_risco_etapa(pct_ac, saldo, vc)
        muted = "color:var(--text3);" if pct_ac == 0 else ""
 
        rows_html += f"""
        <div style="display:grid;grid-template-columns:28px 1fr 115px 115px 115px 72px;
                    gap:8px;align-items:center;padding:9px 0;border-bottom:1px solid var(--border);">
            <span style="font-size:11px;font-family:'DM Mono',monospace;color:var(--text3);">{e["item"]:02d}</span>
            <span style="font-size:12px;color:var(--text);font-weight:500;{muted}">{e["nome"]}</span>
            <span style="font-size:11px;font-family:'DM Mono',monospace;color:var(--text2);text-align:right;">{fmt_money(vt)}</span>
            <span style="font-size:11px;font-family:'DM Mono',monospace;color:var(--green);text-align:right;">{fmt_money(vm_acum)}</span>
            <span style="font-size:11px;font-family:'DM Mono',monospace;color:var(--accent);text-align:right;">{fmt_money(saldo)}</span>
            <div style="text-align:center;">
                <span style="display:inline-block;padding:2px 7px;border-radius:10px;
                             background:{risco_bg};color:{risco_cor};font-size:9px;font-weight:600;">
                    {risco_txt}
                </span>
            </div>
        </div>
        """
 
    legenda = """
    <div style="margin-top:10px;display:flex;gap:10px;flex-wrap:wrap;">
        <span style="font-size:10px;color:var(--text3);">Risco:</span>
        <span style="font-size:10px;color:var(--red);">● Alto — saldo &gt;10% e avanço &lt;5%</span>
        <span style="font-size:10px;color:var(--amber);">● Médio — saldo &gt;5% e avanço &lt;30%</span>
        <span style="font-size:10px;color:var(--green);">● Baixo — avanço ≥ 90%</span>
        <span style="font-size:10px;color:var(--accent2);">● Normal</span>
    </div>
    """
 
    card("Avanço por Etapa (EAP Orçamento)", header + rows_html + legenda, "fisico-equal-card", min_h)


def render_historico_medicoes(data: dict):
    """Histórico de medições — uma linha por medição."""
    med       = data["medicoes"]
    df_totais = data["_df_totais"]
    med_atual = med["medicao_atual"]
    MESES_PT  = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
    import datetime as _dt

    rows_html = ""
    for i in range(1, med_atual + 1):
        row_t = df_totais[df_totais["medicao"] == i]
        pct_m = float(row_t["total_pct_mes"].iloc[0])     if (len(row_t) and pd.notna(row_t["total_pct_mes"].iloc[0]))     else 0.0
        pct_a = float(row_t["total_pct_acum"].iloc[0])    if (len(row_t) and pd.notna(row_t["total_pct_acum"].iloc[0]))    else 0.0
        v_m   = float(row_t["total_medido"].iloc[0])      if (len(row_t) and pd.notna(row_t["total_medido"].iloc[0]))      else 0.0
        v_ac  = float(row_t["total_medido_acum"].iloc[0]) if (len(row_t) and pd.notna(row_t["total_medido_acum"].iloc[0])) else 0.0
        periodo = ""
        if len(row_t) and row_t["data_ref"].iloc[0] and isinstance(row_t["data_ref"].iloc[0], _dt.date):
            d = row_t["data_ref"].iloc[0]
            periodo = f"{MESES_PT[d.month-1]}/{d.year}"
        is_atual = (i == med_atual)
        tr_class = 'class="current-row"' if is_atual else ""
        hl   = "highlight" if is_atual else ""
        lbl  = f"{i:02d} ▶" if is_atual else f"{i:02d}"
        rows_html += f"""
        <tr {tr_class}>
            <td class="mono {hl}">{lbl}</td>
            <td class="mono {hl}">{periodo}</td>
            <td class="mono {hl}">{fmt_percent(pct_m)}</td>
            <td class="mono {hl}">{fmt_percent(pct_a)}</td>
            <td class="mono {hl}" style="text-align:right">{fmt_money(v_m)}</td>
            <td class="mono {hl}" style="text-align:right">{fmt_money(v_ac)}</td>
        </tr>
        """

    body = f"""
    <table class="med-table">
        <thead>
            <tr>
                <th>Med.</th><th>Período</th><th>% Mês</th><th>% Acum.</th>
                <th style="text-align:right">Valor Medido</th>
                <th style="text-align:right">Acumulado</th>
            </tr>
        </thead>
        <tbody>{rows_html}</tbody>
    </table>
    <div style="margin-top:14px;padding:10px;background:var(--blue-bg);border-radius:8px;display:flex;justify-content:space-between;align-items:center;">
        <span style="font-size:12px;color:var(--accent2);font-weight:500;">Saldo contratual</span>
        <span style="font-size:15px;font-weight:600;color:var(--accent);font-family:'DM Mono',monospace;">{fmt_money(data["saldo_contratual"])}</span>
    </div>
    """
    card("Histórico de Medições", body)


def render_fin_summary(data: dict):
    html(
        f"""
        <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:20px;">
            <div class="card" style="text-align:center;">
                <div class="kpi-label">Valor do Contrato</div>
                <div class="kpi-value small">{fmt_money(data["valor_contrato"])}</div>
                <div class="kpi-sub">EAP DE MEDIÇÃO — soma itens nível 1</div>
            </div>
            <div class="card" style="text-align:center;">
                <div class="kpi-label">Executado Acumulado</div>
                <div class="kpi-value small" style="color:var(--green);">{fmt_money(data["valor_medido_acumulado"])}</div>
                <div class="kpi-sub">{fmt_percent(data["medicoes"]["avanco_fisico_acumulado"])} do contrato</div>
            </div>
            <div class="card" style="text-align:center;">
                <div class="kpi-label">Saldo Contratual</div>
                <div class="kpi-value small" style="color:var(--accent2);">{fmt_money(data["saldo_contratual"])}</div>
                <div class="kpi-sub">{fmt_percent(data["saldo_percentual"])} disponível</div>
            </div>
        </div>
        """
    )


def render_fin_composition(data: dict):
    """Composição financeira por etapa — valores acumulados reais da EAP."""
    df_eap    = data["_df_eap"]
    med_atual = data["medicoes"]["medicao_atual"]
    df_orc    = data["_config"]["eap_orcamento"].copy()

    # Linhas da medição atual, nível 1 (sem ponto no item)
    df_at = df_eap[(df_eap["medicao"] == med_atual)].copy()
    df_at = df_at[df_at["item"].apply(lambda x: x is not None and "." not in str(x))]

    # Para o preço total real de cada etapa, usar df_eap (col preco_total_rs = EAP col E)
    # que tem o valor contratual correto por etapa
    df_preco = df_eap[df_eap["item"].apply(
        lambda x: x is not None and "." not in str(x)
    )].drop_duplicates(subset=["item"])[["item", "preco_total_rs"]]

    COLOR = {0: "#2D9B63", 1: "#2E6DA4", 2: "#D4910E"}
    blocks = ""
    count  = 0

    for i, row_orc in df_orc.iterrows():
        item_id = row_orc["item"]
        nome    = row_orc["descricao"]

        # Pegar preço total da EAP (correto) — fallback para CONFIG valor_rs
        r_preco = df_preco[df_preco["item"] == item_id]
        if len(r_preco) and pd.notna(r_preco["preco_total_rs"].iloc[0]):
            vt = float(r_preco["preco_total_rs"].iloc[0])
        else:
            vt = float(row_orc["valor_rs"] or 0) if pd.notna(row_orc["valor_rs"]) else 0.0

        r_at = df_at[df_at["item"] == item_id]
        if len(r_at) and pd.notna(r_at["pct_acumulado"].iloc[0]):
            pct_ac  = float(r_at["pct_acumulado"].iloc[0])
        else:
            pct_ac  = 0.0
        # Valor acumulado = % acumulado × preço total da etapa
        vm_acum = vt * pct_ac
        saldo   = max(vt - vm_acum, 0)
        color   = COLOR.get(i, "#B0ADA6")

        blocks += f"""
        <div style="padding:14px;background:var(--surface2);border-radius:8px;border-left:4px solid {color};">
            <div class="kpi-label">{nome}</div>
            <div class="kpi-value small">{fmt_money(vm_acum)}</div>
            <div class="kpi-sub">{fmt_percent(pct_ac)} acumulado · saldo {fmt_money(saldo)}</div>
        </div>
        """
        count += 1

    if count == 0:
        card("Composição Financeira por Etapa", "<p>Nenhuma etapa com dados disponíveis.</p>")
        return

    cols = min(count, 4)
    card(
        "Composição Financeira por Etapa",
        f'<div style="display:grid;grid-template-columns:repeat({cols},1fr);gap:12px;">{blocks}</div>',
    )


def _gerar_analise_idp(idp_val: float, prazos: dict) -> str:
    """Gera texto de análise de tendência com base no IDP. Independe do cache."""
    nova_data = prazos.get("nova_data_termino") or None
    nova_data_fmt = fmt_date(nova_data) if nova_data else prazos.get("prazo_execucao_data", "—")

    if idp_val >= 1.05:
        return (
            f"Com IDP = {idp_val:.2f}, a obra apresenta desempenho físico <strong>acima do planejado</strong>. "
            f"O ritmo atual indica conclusão antes da data contratual. "
            f"Tendência de término: <strong>{nova_data_fmt}</strong>, sem necessidade de aditivo de prazo."
        )
    elif idp_val >= 1.0:
        return (
            f"Com IDP = {idp_val:.2f}, a obra apresenta desempenho físico <strong>exatamente alinhado</strong> "
            f"ao cronograma contratual. A tendência de término mantém a data prevista de "
            f"<strong>{nova_data_fmt}</strong>, sem necessidade de aditivo de prazo."
        )
    elif idp_val >= 0.9:
        deficit = (1.0 - idp_val) * 100
        return (
            f"Com IDP = {idp_val:.2f}, a obra apresenta <strong>leve atraso físico</strong> em relação ao "
            f"cronograma ({deficit:.1f}% de déficit). Atenção redobrada é necessária. "
            f"Tendência de término: <strong>{nova_data_fmt}</strong>. Avalie a necessidade de aceleração."
        )
    elif idp_val >= 0.75:
        deficit = (1.0 - idp_val) * 100
        return (
            f"Com IDP = {idp_val:.2f}, a obra está em <strong>atraso significativo</strong> "
            f"({deficit:.1f}% abaixo do planejado). Risco elevado de descumprimento do prazo contratual. "
            f"Tendência de término: <strong>{nova_data_fmt}</strong>. Plano de recuperação urgente."
        )
    else:
        deficit = (1.0 - idp_val) * 100
        return (
            f"Com IDP = {idp_val:.2f}, a obra está em situação <strong>crítica de atraso</strong> "
            f"({deficit:.1f}% abaixo do planejado). Alta probabilidade de necessidade de aditivo de prazo. "
            f"Tendência de término: <strong>{nova_data_fmt}</strong>. Intervenção imediata necessária."
        )


def render_prazos_progress(data: dict, return_markup: bool = False):
    p = data["prazos"]

    vig_dec   = p["vigencia_decorrida_dias"]
    vig_total = p["vigencia_total_dias"]
    vig_rest  = p["vigencia_restante_dias"]
    vig_pct   = vig_dec / vig_total if vig_total else 0

    exe_dec   = p["execucao_decorrida_dias"]
    exe_total = p["execucao_total_dias"]
    exe_rest  = p["execucao_restante_dias"]
    exe_pct   = exe_dec / exe_total if exe_total else 0

    idp     = data["idp"]
    idp_val = idp["valor"]

    # Cor do bloco de análise baseada no IDP
    if idp_val >= 1.0:
        tend_bg    = "var(--green-bg)"
        tend_color = "var(--green)"
    elif idp_val >= 0.9:
        tend_bg    = "var(--amber-bg)"
        tend_color = "var(--amber)"
    else:
        tend_bg    = "var(--red-bg)"
        tend_color = "var(--red)"

    body = f"""
    <div style="margin-bottom:20px;">
        <div style="display:flex;justify-content:space-between;margin-bottom:6px;">
            <span style="font-size:12px;color:var(--text2);font-weight:500;">Vigência do Contrato</span>
            <span style="font-size:12px;font-family:'DM Mono',monospace;color:var(--text);">{vig_dec} / {vig_total} dias</span>
        </div>
        <div class="prazo-track">
            <div class="prazo-decorrido" style="width:{vig_pct * 100:.1f}%">
                <span>{fmt_percent(vig_pct)} decorrida</span>
            </div>
            <span class="prazo-restante-label">{vig_rest} dias restantes</span>
        </div>
        <div style="display:flex;justify-content:flex-end;font-size:10px;color:var(--text3);margin-top:4px;">
            <span>Fim: {p["prazo_vigencia_data"]}</span>
        </div>
    </div>

    <div>
        <div style="display:flex;justify-content:space-between;margin-bottom:6px;">
            <span style="font-size:12px;color:var(--text2);font-weight:500;">Prazo de Execução da Obra</span>
            <span style="font-size:12px;font-family:'DM Mono',monospace;color:var(--text);">{exe_dec} / {exe_total} dias</span>
        </div>
        <div class="prazo-track">
            <div class="prazo-decorrido" style="width:{exe_pct * 100:.1f}%;background:var(--accent)">
                <span>{fmt_percent(exe_pct)}</span>
            </div>
            <span class="prazo-restante-label">{exe_rest} dias restantes</span>
        </div>
        <div style="display:flex;justify-content:flex-end;font-size:10px;color:var(--text3);margin-top:4px;">
            <span>Fim: {p["prazo_execucao_data"]}</span>
        </div>
    </div>

    <div style="margin-top:20px;padding:14px;border-radius:8px;background:{tend_bg};">
        <div style="font-size:11px;font-weight:600;color:{tend_color};text-transform:uppercase;letter-spacing:0.6px;margin-bottom:6px;">
            Análise de Tendência · IDP = {fmt_decimal(idp_val)}
        </div>
        <div style="font-size:12px;color:{tend_color};line-height:1.65;">
            {idp.get("analise", "") or _gerar_analise_idp(idp_val, data["prazos"])}
        </div>
    </div>
    """
    markup = card_markup("Progresso Temporal", body, "prazos-equal-card")
    if return_markup:
        return markup
    html(markup)


def render_timeline(data: dict, return_markup: bool = False):
    p   = data["prazos"]
    idp = data["idp"]

    items = [
        ("green", "Prazo de Vigência",    p["prazo_vigencia_data"],  p["vigencia_decorrida_texto"],  "green"),
        ("green", "Vigência Restante",     "CONFIG · vigência",       p["vigencia_restante_texto"],   "green"),
        ("blue",  "Prazo de Execução",     p["prazo_execucao_data"],  p["execucao_decorrida_texto"],  "blue"),
        ("blue",  "Execução Restante",     "CONFIG · execução",       p["execucao_restante_texto"],   "blue"),
        ("green", "IDP",                   "CONFIG · M27",            fmt_decimal(idp["valor"]),       "green"),
        ("gray",  "Status",                "—",                       idp["status"],                   "default"),
    ]

    body = ""
    for dot, title, meta, value, color in items:
        cls = "text-green" if color == "green" else "text-blue" if color == "blue" else ""
        body += f"""
        <div class="timeline-item">
            <div class="tl-dot tl-dot-{dot}"></div>
            <div class="tl-content">
                <div class="tl-title">{title}</div>
                <div class="tl-meta">{meta}</div>
            </div>
            <div class="tl-val {cls}">{value}</div>
        </div>
        """
    markup = card_markup("Cronograma de Marcos", body, "prazos-equal-card")
    if return_markup:
        return markup
    html(markup)


def info_grid(items: list[tuple[str, str]]) -> str:
    rows = ""
    for label, value in items:
        rows += f"""
        <div class="info-row">
            <span class="info-label">{label}</span>
            <span class="info-val">{value}</span>
        </div>
        """
    return f'<div class="info-grid">{rows}</div>'


def _equalize_js(class_name: str):
    """
    No-op: a equalização de altura é feita pelo script global injetado em inject_css(),
    que observa o DOM e equaliza todos os grupos de cards com classes *-equal-card.
    """
    pass


# =============================================================================
# ABAS
# =============================================================================

def tab_visao(data: dict):
    render_kpis(data)
    spacer(22)
    card_pair(
        render_avanco_etapas(data, return_markup=True),
        render_situacao_obra(data, return_markup=True),
    )
    spacer(22)
    plot_curva_s(data)


def render_ultimos_servicos(data: dict):
    """Card: últimos serviços concluídos na medição atual em ordem cronológica."""
    df_eap    = data["_df_eap"]
    med_atual = data["medicoes"]["medicao_atual"]
    MESES_PT  = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
    import datetime as _dt

    # Itens com percentual de mês > 0 na medição atual (alguma evolução)
    df_med = df_eap[df_eap["medicao"] == med_atual].copy()
    df_med = df_med[df_med["pct_mes"].notna() & (df_med["pct_mes"] > 0)]

    # Ordenar por número de item hierarquicamente (1 < 1.3 < 1.4 < 2 < 3 < 3.1.1.1...)
    def _item_sort_key(item_val):
        if item_val is None:
            return (999,)
        parts = str(item_val).split(".")
        try:
            return tuple(int(p) for p in parts)
        except ValueError:
            return (999,)

    df_med["_sort_key"] = df_med["item"].apply(_item_sort_key)
    df_med_sorted = df_med.sort_values("_sort_key")

    # Pegar data de referência da medição atual
    row_ref = data["_df_totais"][data["_df_totais"]["medicao"] == med_atual]
    data_med = ""
    if len(row_ref) and row_ref["data_ref"].iloc[0] and isinstance(row_ref["data_ref"].iloc[0], _dt.date):
        d = row_ref["data_ref"].iloc[0]
        data_med = f"{MESES_PT[d.month-1]}/{d.year}"

    if len(df_med_sorted) == 0:
        return  # Nada a mostrar

    itens = df_med_sorted.head(20)  # limita a 20 itens

    rows_html = ""
    for _, row in itens.iterrows():
        item_id   = row["item"] if row["item"] is not None else "—"
        descricao = row["descricao"] or "—"
        pct_mes   = float(row["pct_mes"]) * 100 if pd.notna(row["pct_mes"]) else 0
        pct_acum  = float(row["pct_acumulado"]) * 100 if pd.notna(row["pct_acumulado"]) else 0
        concluido = pct_acum >= 100.0

        badge = ""
        if concluido:
            badge = '<span style="display:inline-block;padding:2px 7px;border-radius:10px;background:var(--green-bg);color:var(--green);font-size:10px;font-weight:600;">✓ Concluído</span>'
        else:
            badge = f'<span style="display:inline-block;padding:2px 7px;border-radius:10px;background:var(--blue-bg);color:var(--accent2);font-size:10px;font-weight:500;">Em andamento</span>'

        rows_html += f"""
        <div style="display:grid;grid-template-columns:48px 1fr 90px 95px 80px;gap:10px;align-items:center;
                    padding:9px 0;border-bottom:1px solid var(--border);">
            <span style="font-size:11px;color:var(--text3);font-family:'DM Mono',monospace;">{item_id}</span>
            <span style="font-size:12px;color:var(--text);font-weight:500;">{descricao}</span>
            {badge}
            <span style="font-size:12px;font-family:'DM Mono',monospace;color:var(--green);text-align:right;">+{pct_mes:.2f}%</span>
            <span style="font-size:12px;font-family:'DM Mono',monospace;color:var(--text);text-align:right;">{pct_acum:.2f}% acum.</span>
        </div>
        """

    header = f"""
    <div style="display:grid;grid-template-columns:48px 1fr 90px 95px 80px;gap:10px;
                padding:0 0 8px;border-bottom:2px solid var(--border);margin-bottom:2px;">
        <span style="font-size:10px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:0.6px;">Item</span>
        <span style="font-size:10px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:0.6px;">Serviço / Descrição</span>
        <span style="font-size:10px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:0.6px;">Status</span>
        <span style="font-size:10px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:0.6px;text-align:right;">Avanço Mês</span>
        <span style="font-size:10px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:0.6px;text-align:right;">% Acumulado</span>
    </div>
    """

    legenda = f'<div style="font-size:11px;color:var(--text3);margin-top:12px;">Medição {med_atual} · {data_med} · {len(itens)} serviços com evolução nesta medição</div>'

    card(
        f"Detalhamento — Serviços com Evolução na {med_atual}ª Medição",
        header + rows_html + legenda,
    )


def tab_fisico(data: dict):
    FISICO_H = _calc_fisico_height(data)
    c1, c2 = st.columns(2, gap="small")
    with c1:
        render_etapas_table(data, FISICO_H)
    with c2:
        plot_medicoes(data)
        spacer(22)
        plot_barras_etapas(data)
    spacer(22)
    render_ultimos_servicos(data)


def tab_financeiro(data: dict):
    render_fin_summary(data)
    c1, c2 = st.columns(2, gap="small")
    with c1:
        render_historico_medicoes(data)
    with c2:
        plot_financeiro(data)
    spacer(22)
    render_fin_composition(data)


def render_analise_ritmo(data: dict):
    """
    Análise de Ritmo de Execução:
    - Ritmo médio realizado (% por medição)
    - Ritmo necessário para concluir no prazo (% restante / medições restantes)
    - Projeção simples de conclusão ao ritmo atual
    """
    med       = data["medicoes"]
    prazos    = data["prazos"]
    df_totais = data["_df_totais"]
    med_atual = med["medicao_atual"]

    # Medições restantes estimadas (dias restantes / 30.44)
    dias_rest     = prazos["execucao_restante_dias"]
    medicoes_rest = max(round(dias_rest / 30.44), 1)
    medicoes_tot  = max(round(prazos["execucao_total_dias"] / 30.44), 1)

    # Avanço acumulado atual (0–1)
    pct_acum = float(med["avanco_fisico_acumulado"] or 0)
    pct_rest  = max(1.0 - pct_acum, 0)

    # Ritmo realizado médio por medição (excluindo medições com zero)
    ritmos = []
    for i in range(1, med_atual + 1):
        row = df_totais[df_totais["medicao"] == i]
        if len(row) and pd.notna(row["total_pct_mes"].iloc[0]):
            v = float(row["total_pct_mes"].iloc[0])
            if v > 0:
                ritmos.append(v)

    ritmo_medio_real = sum(ritmos) / len(ritmos) if ritmos else 0.0
    ritmo_necessario = pct_rest / medicoes_rest if medicoes_rest > 0 else 0.0

    # Projeção: com ritmo atual, quantas medições faltam
    if ritmo_medio_real > 0:
        med_proj = round(pct_rest / ritmo_medio_real)
    else:
        med_proj = None

    # Diferença entre ritmo necessário e realizado
    delta = ritmo_medio_real - ritmo_necessario

    if delta >= 0.005:
        delta_cor   = "var(--green)"
        delta_bg    = "var(--green-bg)"
        delta_sinal = "+"
        delta_msg   = "acima do necessário — tendência de conclusão antecipada"
    elif delta >= -0.005:
        delta_cor   = "var(--accent2)"
        delta_bg    = "var(--blue-bg)"
        delta_sinal = "≈"
        delta_msg   = "alinhado ao necessário — manter o ritmo atual"
    elif delta >= -0.02:
        delta_cor   = "var(--amber)"
        delta_bg    = "var(--amber-bg)"
        delta_sinal = ""
        delta_msg   = "abaixo do necessário — atenção ao ritmo de execução"
    else:
        delta_cor   = "var(--red)"
        delta_bg    = "var(--red-bg)"
        delta_sinal = ""
        delta_msg   = "significativamente abaixo — risco de não cumprimento do prazo"

    proj_html = (
        f"<strong>{med_proj}</strong> medições adicionais "
        f"(≈ {round(med_proj * 30.44 / 30.44)} meses ao ritmo atual)"
        if med_proj is not None else "Não calculável"
    )

    body = f"""
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:18px;">
        <div style="background:var(--surface2);border-radius:8px;padding:12px 14px;">
            <div style="font-size:10px;color:var(--text3);text-transform:uppercase;letter-spacing:0.6px;">
                Medições realizadas
            </div>
            <div style="font-size:22px;font-weight:600;color:var(--text);font-family:'DM Mono',monospace;margin-top:4px;">
                {med_atual}
            </div>
            <div style="font-size:10px;color:var(--text3);margin-top:2px;">de ~{medicoes_tot} estimadas</div>
        </div>
        <div style="background:var(--surface2);border-radius:8px;padding:12px 14px;">
            <div style="font-size:10px;color:var(--text3);text-transform:uppercase;letter-spacing:0.6px;">
                Ritmo médio realizado
            </div>
            <div style="font-size:22px;font-weight:600;color:var(--text);font-family:'DM Mono',monospace;margin-top:4px;">
                {fmt_percent(ritmo_medio_real)}/mês
            </div>
            <div style="font-size:10px;color:var(--text3);margin-top:2px;">média das medições com evolução</div>
        </div>
        <div style="background:var(--surface2);border-radius:8px;padding:12px 14px;">
            <div style="font-size:10px;color:var(--text3);text-transform:uppercase;letter-spacing:0.6px;">
                Ritmo necessário
            </div>
            <div style="font-size:22px;font-weight:600;color:var(--accent);font-family:'DM Mono',monospace;margin-top:4px;">
                {fmt_percent(ritmo_necessario)}/mês
            </div>
            <div style="font-size:10px;color:var(--text3);margin-top:2px;">para concluir em ~{medicoes_rest} medições restantes</div>
        </div>
        <div style="background:{delta_bg};border-radius:8px;padding:12px 14px;">
            <div style="font-size:10px;color:{delta_cor};text-transform:uppercase;letter-spacing:0.6px;">
                Δ Ritmo (real − necessário)
            </div>
            <div style="font-size:22px;font-weight:600;color:{delta_cor};font-family:'DM Mono',monospace;margin-top:4px;">
                {delta_sinal}{fmt_percent(abs(delta))}/mês
            </div>
            <div style="font-size:10px;color:{delta_cor};margin-top:2px;">{delta_msg}</div>
        </div>
    </div>
    <div style="background:var(--surface2);border-radius:8px;padding:12px 16px;display:flex;
                align-items:center;gap:14px;">
        <div style="font-size:20px;">📈</div>
        <div>
            <div style="font-size:11px;font-weight:600;color:var(--text2);text-transform:uppercase;
                        letter-spacing:0.6px;margin-bottom:2px;">Projeção ao ritmo atual</div>
            <div style="font-size:13px;color:var(--text);line-height:1.6;">
                Restam <strong>{fmt_percent(pct_rest)}</strong> de avanço físico a executar.
                Ao ritmo médio atual de <strong>{fmt_percent(ritmo_medio_real)}/mês</strong>,
                serão necessárias aproximadamente {proj_html} para conclusão.
            </div>
        </div>
    </div>
    """
    card("Análise de Ritmo de Execução", body)


def tab_prazos(data: dict):
    p   = data["prazos"]
    cfg = data["_config"]
    tend = cfg["tendencia"]

    # Prazo de execução: usar nova_duracao_meses da CONFIG, arredondado para baixo
    import math as _math
    nova_dur_meses = tend.get("nova_duracao_meses")
    if nova_dur_meses is not None and not (isinstance(nova_dur_meses, float) and pd.isna(nova_dur_meses)):
        exec_meses_label = f"{int(_math.floor(float(nova_dur_meses)))} meses"
    else:
        exe_meses = int(_math.floor(p["execucao_total_dias"] / 30.44))
        exec_meses_label = f"{exe_meses} meses"

    vig_meses = round(p["vigencia_total_dias"] / 30.44)

    cols = st.columns(3, gap="small")
    with cols[0]:
        kpi(
            "Prazo de Execução",
            exec_meses_label,
            f'OS: {fmt_date(cfg["execucao"]["inicio"])} → {p["prazo_execucao_data"]}',
            p["execucao_restante_texto"],
            "blue", small=True,
        )
    with cols[1]:
        kpi(
            "Prazo de Vigência",
            f"{int(vig_meses)} meses",
            f'{fmt_date(cfg["vigencia"]["inicio"])} → {p["prazo_vigencia_data"]}',
            p["vigencia_restante_texto"],
            "blue", small=True,
        )
    with cols[2]:
        kpi(
            "Tendência de Término",
            fmt_date(p["nova_data_termino"]) if p["nova_data_termino"] else p["prazo_execucao_data"],
            f'IDP = {fmt_decimal(data["idp"]["valor"])}',
            data["idp"]["status"],
            "green", small=True,
        )

    spacer(22)
    card_pair(
        render_prazos_progress(data, return_markup=True),
        render_timeline(data, return_markup=True),
    )
    spacer(22)
    render_analise_ritmo(data)


def tab_contrato(data: dict):
    m = data["medicoes"]
    p = data["prazos"]
    cfg = data["_config"]
    tend = cfg["tendencia"]

    # Prazo de execução dinâmico a partir da CONFIG, arredondado para baixo
    import math as _math
    nova_dur_meses = tend.get("nova_duracao_meses")
    if nova_dur_meses is not None and not (isinstance(nova_dur_meses, float) and pd.isna(nova_dur_meses)):
        exec_meses_str = str(int(_math.floor(float(nova_dur_meses))))
    else:
        exec_meses_str = str(int(_math.floor(p["execucao_total_dias"] / 30.44)))
    termino_exec = fmt_date(tend["nova_data_termino"]) if tend["nova_data_termino"] else p["prazo_execucao_data"]

    # Dados fixos do contrato
    CONTRATO = {
        "objeto":         "Construção da Nova Sede da Justiça Federal em Juazeiro do Norte / CE",
        "num_contrato":   "25/2025",
        "concorrencia":   "90001/2025",
        "cno":            "90.025.93927/72",
        "contratada":     "Consórcio Juazeiro do Norte",
        "cnpj":           "62.009.288/0001-12",
        "val_inicial":    28_566_749.33,
        "val_atual":      28_566_749.33,
        "aditivos_valor": "Nenhum",
        "assinatura":     "04/08/2025",
        "ordem_servico":  "10/12/2025",
        "prazo_exec":     f"{exec_meses_str} meses · Término em {termino_exec}",
        "prazo_vig":      f"40 meses · Término em 04/12/2028",
        "aditivos_prazo": "Nenhum",
        "situacao":       "Vigente",
        "paralisacoes":   "Nenhuma registrada",
        "endereco":       "Rua José Geraldo da Cruz (L) com a Rua Presidente Médici (O), na Rua Manoel Pires (N) com a Rua Frei Damião (S), Bairro Lagoa Seca, Juazeiro do Norte — CE.",
    }

    def fld(label, value, bold=False):
        style = "font-weight:600;" if bold else ""
        return f"""
        <div style="padding:10px 0;border-bottom:1px solid var(--border);">
            <div style="font-size:10px;color:var(--text3);text-transform:uppercase;letter-spacing:0.5px;">{label}</div>
            <div style="font-size:13px;color:var(--text);margin-top:3px;{style}">{value}</div>
        </div>"""

    badge_html = '<span style="display:inline-block;padding:3px 10px;border-radius:20px;background:#2D9B6322;color:#2D9B63;font-size:12px;font-weight:500;">Vigente</span>'

    # Identificação — dois campos por linha via grid
    id_html = f"""
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:0 20px;">
        {fld("Objeto", CONTRATO["objeto"])}
        {fld("Nº do Contrato", CONTRATO["num_contrato"])}
        {fld("Concorrência Pública", CONTRATO["concorrencia"])}
        {fld("CNO da Obra", CONTRATO["cno"])}
        {fld("Empresa Contratada", CONTRATO["contratada"])}
        {fld("CNPJ", CONTRATO["cnpj"])}
        <div style="padding:10px 0;border-bottom:1px solid var(--border);">
            <div style="font-size:10px;color:var(--text3);text-transform:uppercase;letter-spacing:0.5px;">Situação do Contrato</div>
            <div style="margin-top:4px;">{badge_html}</div>
        </div>
        {fld("Paralisações", CONTRATO["paralisacoes"])}
    </div>"""

    # Valores e Prazos — dois campos por linha via grid
    vp_html = f"""
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:0 20px;">
        {fld("Valor Inicial do Contrato", fmt_money(CONTRATO["val_inicial"]), bold=True)}
        {fld("Valor Atual do Contrato",   fmt_money(CONTRATO["val_atual"]),   bold=True)}
        {fld("Aditivos de Valor", CONTRATO["aditivos_valor"])}
        {fld("Assinatura do Contrato", CONTRATO["assinatura"])}
        {fld("Ordem de Serviço", CONTRATO["ordem_servico"])}
        {fld("Prazo de Execução", CONTRATO["prazo_exec"])}
        {fld("Prazo de Vigência", CONTRATO["prazo_vig"])}
        {fld("Aditivos de Prazo", CONTRATO["aditivos_prazo"])}
    </div>"""

    card_pair(
        card_markup("Identificação do Contrato", id_html, "contrato-equal-card"),
        card_markup("Valores e Prazos", vp_html, "contrato-equal-card"),
    )

    spacer(22)
    card("Endereço da Obra",
         f'<p style="font-size:13px;color:var(--text2);line-height:1.7;margin:0;">{CONTRATO["endereco"]}</p>')

    spacer(22)

    # Tabela de totais por medição
    df_totais = data["_df_totais"].copy()
    MESES_PT  = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]

def _projetos_excel_path() -> Path:
    return DATA_DIR / "registros_projetos.xlsx"


def _carregar_registros_projetos() -> pd.DataFrame:
    """Carrega histórico de registros (aba Resumo) do Excel persistente."""
    p = _projetos_excel_path()
    if p.exists():
        try:
            return pd.read_excel(p, sheet_name="Resumo", dtype=str)
        except Exception:
            try:
                return pd.read_excel(p, dtype=str)
            except Exception:
                pass
    return pd.DataFrame(columns=[
        "medicao", "data_registro", "usuario",
        "avanco_geral_pct", "montante_total", "saldo_devedor",
        "status_geral", "observacoes",
    ])


def _carregar_pagamentos_projetos() -> pd.DataFrame:
    """Carrega histórico de pagamentos por etapa (aba Pagamentos) do Excel persistente."""
    p = _projetos_excel_path()
    if p.exists():
        try:
            return pd.read_excel(p, sheet_name="Pagamentos", dtype=str)
        except Exception:
            pass
    return pd.DataFrame(columns=[
        "medicao", "etapa", "valor", "status",
    ])


def _salvar_somente_resumo(registro: dict, pagamentos: list):
    """Salva apenas a aba Resumo, preservando a aba Pagamentos existente."""
    df_resumo = _carregar_registros_projetos()
    df_pags   = _carregar_pagamentos_projetos()
    med_str   = str(registro["medicao"])

    if len(df_resumo) > 0 and "medicao" in df_resumo.columns:
        df_resumo = df_resumo[df_resumo["medicao"].astype(str) != med_str]

    pags_para_status = pagamentos or [
        {"status": r["status"]}
        for _, r in df_pags[df_pags["medicao"].astype(str) == med_str].iterrows()
        if "status" in df_pags.columns
    ] if len(df_pags) > 0 else []

    def _ss(s):
        return {"Pago": 1.0, "Pendente": 0.5, "Não Pago": 0.0}.get(s, 0.5)

    if pags_para_status:
        scores = [_ss(p["status"]) for p in pags_para_status]
        media  = sum(scores) / len(scores)
        registro["status_geral"] = "Pago" if media >= 0.9 else ("Pendente" if media >= 0.4 else "Não Pago")
    else:
        registro["status_geral"] = "Pendente"

    df_resumo_final = pd.concat([df_resumo, pd.DataFrame([registro])], ignore_index=True)

    with pd.ExcelWriter(_projetos_excel_path(), engine="openpyxl") as writer:
        df_resumo_final.to_excel(writer, sheet_name="Resumo",     index=False)
        df_pags.to_excel(writer,         sheet_name="Pagamentos", index=False)


def _salvar_somente_pagamentos(med_str: str, pagamentos: list):
    """Salva apenas a aba Pagamentos para a medição indicada, preservando a aba Resumo.
    Funciona mesmo que o arquivo ainda não exista (cria as duas abas do zero)."""
    df_resumo = _carregar_registros_projetos()
    df_pags   = _carregar_pagamentos_projetos()

    if len(df_pags) > 0 and "medicao" in df_pags.columns:
        df_pags = df_pags[df_pags["medicao"].astype(str) != med_str]

    if not pagamentos:
        raise ValueError("Nenhum pagamento para salvar. Adicione ao menos uma etapa.")

    novos_pags = pd.DataFrame([
        {
            "medicao": med_str,
            "etapa":   p["etapa"],
            "pct":     f"{p.get('pct', 0.0):.1f}",
            "valor":   f"{p['valor']:.2f}",
            "status":  p["status"],
        }
        for p in pagamentos
    ])

    df_pags_final = pd.concat([df_pags, novos_pags], ignore_index=True)

    def _ss(s):
        return {"Pago": 1.0, "Pendente": 0.5, "Não Pago": 0.0}.get(s, 0.5)

    scores = [_ss(p["status"]) for p in pagamentos]
    media  = sum(scores) / len(scores)
    status_geral = "Pago" if media >= 0.9 else ("Pendente" if media >= 0.4 else "Não Pago")
    if len(df_resumo) > 0 and "medicao" in df_resumo.columns:
        mask = df_resumo["medicao"].astype(str) == med_str
        df_resumo.loc[mask, "status_geral"] = status_geral

    with pd.ExcelWriter(_projetos_excel_path(), engine="openpyxl") as writer:
        df_resumo.to_excel(writer,       sheet_name="Resumo",     index=False)
        df_pags_final.to_excel(writer,   sheet_name="Pagamentos", index=False)


def _salvar_registro_projeto(registro: dict, pagamentos: list):
    """
    Salva registro de projetos em Excel com duas abas:
    - Resumo: uma linha por medição (montante, saldo, status geral, etc.)
    - Pagamentos: uma linha por etapa/pagamento de cada medição

    Se já existe um registro para a mesma medição, SUBSTITUI (não duplica).
    """
    import datetime as _dt
    from openpyxl import Workbook

    df_resumo = _carregar_registros_projetos()
    df_pags   = _carregar_pagamentos_projetos()

    med_str = str(registro["medicao"])

    # Remover linhas anteriores da mesma medição
    if len(df_resumo) > 0 and "medicao" in df_resumo.columns:
        df_resumo = df_resumo[df_resumo["medicao"].astype(str) != med_str]
    if len(df_pags) > 0 and "medicao" in df_pags.columns:
        df_pags = df_pags[df_pags["medicao"].astype(str) != med_str]

    # Calcular status geral como média dos status das etapas
    def _status_score(s):
        return {"Pago": 1.0, "Pendente": 0.5, "Não Pago": 0.0}.get(s, 0.5)

    if pagamentos:
        scores = [_status_score(p["status"]) for p in pagamentos]
        media  = sum(scores) / len(scores)
        if media >= 0.9:
            status_geral = "Pago"
        elif media >= 0.4:
            status_geral = "Pendente"
        else:
            status_geral = "Não Pago"
    else:
        status_geral = "Pendente"

    registro["status_geral"] = status_geral

    # Montar linhas de pagamento desta medição
    novos_pags = pd.DataFrame([
        {
            "medicao": med_str,
            "etapa":   p["etapa"],
            "pct":     f"{p.get('pct', 0.0):.1f}",
            "valor":   f"{p['valor']:.2f}",
            "status":  p["status"],
        }
        for p in pagamentos
    ]) if pagamentos else pd.DataFrame(columns=["medicao", "etapa", "pct", "valor", "status"])

    df_resumo_final = pd.concat([df_resumo, pd.DataFrame([registro])], ignore_index=True)
    df_pags_final   = pd.concat([df_pags, novos_pags], ignore_index=True)

    # Salvar as duas abas no mesmo arquivo
    with pd.ExcelWriter(_projetos_excel_path(), engine="openpyxl") as writer:
        df_resumo_final.to_excel(writer, sheet_name="Resumo",     index=False)
        df_pags_final.to_excel(writer,   sheet_name="Pagamentos", index=False)


def _medicao_ja_existe(medicao_num: int) -> bool:
    """Retorna True se já existe um registro para a medição informada."""
    df = _carregar_registros_projetos()
    if len(df) == 0:
        return False
    return str(medicao_num) in df["medicao"].astype(str).values


def _excluir_resumo(med_str: str):
    """Remove o resumo e todos os pagamentos de uma medição."""
    df_resumo = _carregar_registros_projetos()
    df_pags   = _carregar_pagamentos_projetos()

    if len(df_resumo) > 0:
        df_resumo = df_resumo[df_resumo["medicao"].astype(str) != med_str]
    if len(df_pags) > 0:
        df_pags = df_pags[df_pags["medicao"].astype(str) != med_str]

    with pd.ExcelWriter(_projetos_excel_path(), engine="openpyxl") as writer:
        df_resumo.to_excel(writer, sheet_name="Resumo",     index=False)
        df_pags.to_excel(writer,   sheet_name="Pagamentos", index=False)


def _excluir_pagamento(med_str: str, etapa_str: str):
    """Remove uma linha específica de pagamento (medição + etapa)."""
    df_resumo = _carregar_registros_projetos()
    df_pags   = _carregar_pagamentos_projetos()

    if len(df_pags) > 0:
        mask_del = (df_pags["medicao"].astype(str) == med_str) & (df_pags["etapa"].astype(str) == etapa_str)
        df_pags = df_pags[~mask_del]

    with pd.ExcelWriter(_projetos_excel_path(), engine="openpyxl") as writer:
        df_resumo.to_excel(writer, sheet_name="Resumo",     index=False)
        df_pags.to_excel(writer,   sheet_name="Pagamentos", index=False)


def tab_upload(data: dict):
    import datetime as _dt

    # ── Autenticação ──────────────────────────────────────────────────────
    USERS = {"César": "12", "Ian": "AtleticoMG"}

    if "logado_como" not in st.session_state:
        st.session_state["logado_como"] = None

    if st.session_state["logado_como"] is None:
        html("""
        <div style="max-width:400px;margin:60px auto 0;">
          <div style="text-align:center;margin-bottom:28px;">
            <div style="font-size:22px;font-weight:600;color:var(--text);">Acesso Restrito</div>
            <div style="font-size:13px;color:var(--text3);margin-top:6px;">
              Faça login para atualizar os dados
            </div>
          </div>
        </div>
        """)
        col_l, col_m, col_r = st.columns([1, 2, 1])
        with col_m:
            usuario = st.selectbox("Usuário", list(USERS.keys()), key="login_usuario")
            senha   = st.text_input("Senha", type="password", key="login_senha")
            st.markdown("""
            <style>
            div[data-testid="stForm"] button[kind="primary"],
            [data-testid="stBaseButton-primary"] {
                background-color: var(--accent) !important;
                border-color: var(--accent) !important;
            }
            </style>
            """, unsafe_allow_html=True)
            if st.button("Entrar", use_container_width=True, type="primary"):
                if USERS.get(usuario) == senha:
                    st.session_state["logado_como"] = usuario
                    st.rerun()
                else:
                    st.error("Senha incorreta.")
        return

    # ── Usuário autenticado ───────────────────────────────────────────────
    usuario_atual = st.session_state["logado_como"]
    col_info, col_sair = st.columns([5, 1])
    with col_info:
        html(f"""
        <div style="display:inline-flex;align-items:center;gap:10px;
                    background:var(--green-bg);border-radius:8px;padding:8px 14px;margin-bottom:4px;">
            <span style="font-size:13px;font-weight:600;color:var(--green);">✓ Logado como {usuario_atual}</span>
        </div>
        """)
    with col_sair:
        if st.button("Sair", use_container_width=True):
            st.session_state["logado_como"] = None
            st.rerun()

    spacer(16)

    # ── Sub-abas: Obras / Projetos ─────────────────────────────────────────
    st.markdown("""
    <style>
    div[data-testid="stTabs"] div[data-testid="stTabs"] div[data-baseweb="tab-list"] {
        background: var(--surface2) !important;
        padding: 6px 0 0 !important;
        gap: 4px !important;
        border-bottom: 1px solid var(--border) !important;
        border-radius: 0 !important;
        width: fit-content !important;
        min-width: unset !important;
    }
    div[data-testid="stTabs"] div[data-testid="stTabs"] button[data-baseweb="tab"] {
        background: transparent !important;
        color: var(--text2) !important;
        border-radius: 6px 6px 0 0 !important;
        font-size: 13px !important;
        font-weight: 500 !important;
        padding: 7px 18px !important;
        flex: unset !important;
        width: auto !important;
    }
    div[data-testid="stTabs"] div[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"] {
        background: var(--accent) !important;
        color: white !important;
        font-weight: 600 !important;
    }
    div[data-testid="stTabs"] div[data-testid="stTabs"] {
        background: var(--bg) !important;
        flex: unset !important;
    }
    div[data-testid="stTabs"] div[data-testid="stTabs"] div[data-baseweb="tab-list"] {
        background: var(--surface2) !important;
    }
    div[data-testid="stTabs"] div[data-testid="stTabs"] div[data-baseweb="tab-panel"] {
        background: var(--bg) !important;
        padding: 20px 0 0 !important;
        min-height: unset !important;
    }
    </style>
    """, unsafe_allow_html=True)

    aba_obras, aba_projetos = st.tabs(["🏗 Obras", "📋 Projetos"])

    # ══════════════════════════════════════════════════════════════════════
    # SUBABA OBRAS — Upload EAP
    # ══════════════════════════════════════════════════════════════════════
    with aba_obras:
        card(
            "Upload do Arquivo EAP — Obras",
            """
            <p style="font-size:13px;color:var(--text2);line-height:1.7;margin:0;">
                Envie o arquivo Excel de medição. O app lerá somente as abas
                <strong>CONFIG</strong> e <strong>EAP DE MEDIÇÃO</strong>.
                O arquivo será salvo em <strong>data/eap_atual.xlsx</strong>.
            </p>
            """
        )
        spacer(10)

        st.markdown("""
        <style>
        [data-testid="stFileUploader"] {
            max-width: 560px;
        }
        [data-testid="stFileUploader"] section {
            background: white !important;
            border-radius: 8px !important;
        }
        [data-testid="stFileUploader"] section > div {
            background: white !important;
        }
        </style>
        """, unsafe_allow_html=True)

        _uploader_key = f"uploader_eap_{st.session_state.get('_uploader_key', 0)}"
        uploaded = st.file_uploader("Selecione o arquivo Excel", type=["xlsx", "xlsm"], key=_uploader_key)

        # Exibe resultado de upload bem-sucedido (antes do uploader p/ não perder o estado)
        if st.session_state.get("_upload_sucesso_msg"):
            _smsg = st.session_state.pop("_upload_sucesso_msg")
            st.success(_smsg)

        if uploaded is not None:
            file_bytes = uploaded.getvalue()

            # Limpa caches antes de qualquer leitura para garantir dados frescos
            load_from_path.clear()
            load_from_bytes.clear()

            # Detecta a medição do arquivo enviado
            try:
                new_data_preview = load_from_bytes(file_bytes)
                med_novo = int(new_data_preview["medicoes"]["medicao_atual"])
            except Exception as exc:
                st.error(f"Não foi possível ler o arquivo: {exc}")
                med_novo = None

            if med_novo is not None:
                # Detecta medicao atual no arquivo salvo (leitura fresca, cache já limpo acima)
                med_existente = None
                if DEFAULT_EAP_FILE.exists():
                    try:
                        _d = load_from_path(str(DEFAULT_EAP_FILE))
                        med_existente = int(_d["medicoes"]["medicao_atual"])
                    except Exception:
                        pass

                # Determina destino
                dest_file = DATA_DIR / f"eap_medida{med_novo}.xlsx"
                # Sobrescrita: medicao ja existe no sistema OU arquivo de destino ja existe
                eh_sobrescrita  = (
                    (med_existente is not None and med_novo <= med_existente)
                    or dest_file.exists()
                )
                eh_nova_medicao = not eh_sobrescrita

                # Exibe info sobre o arquivo detectado
                html(f"""
                <div style="max-width:560px;background:#EBF0F8;border:1px solid #C8D8EE;border-radius:8px;
                            padding:10px 14px;font-size:13px;color:#1B3A5C;margin-top:4px;margin-bottom:12px;">
                    📊 Arquivo detectado: <strong>{med_novo}ª Medição</strong>
                    {"&nbsp;·&nbsp;Medição atual no sistema: <strong>" + str(med_existente) + "ª</strong>" if med_existente else ""}
                </div>
                """)

                if eh_sobrescrita:
                    # Pede confirmação antes de sobrescrever
                    if not st.session_state.get("_confirmar_upload_eap"):
                        if med_novo == med_existente:
                            _aviso_msg = (
                                f"⚠ O sistema já possui dados da **{med_existente}ª Medição**. "
                                f"O arquivo enviado é da mesma medição. "
                                f"Deseja **substituir** os dados existentes?"
                            )
                        else:
                            _aviso_msg = (
                                f"⚠ O sistema já possui dados até a **{med_existente}ª Medição**. "
                                f"O arquivo enviado é de uma medição anterior (**{med_novo}ª**). "
                                f"Deseja **substituir** os dados existentes por uma versão mais antiga?"
                            )
                        st.warning(_aviso_msg)
                        cu1, cu2, _ = st.columns([1, 1, 3])
                        with cu1:
                            if st.button("✓ Sim, substituir", type="primary", use_container_width=True, key="btn_confirm_upload"):
                                st.session_state["_confirmar_upload_eap"] = True
                                st.session_state["_upload_eap_bytes_pending"] = file_bytes
                                st.session_state["_upload_eap_med_novo"] = med_novo
                                st.rerun()
                        with cu2:
                            if st.button("✕ Cancelar", use_container_width=True, key="btn_cancel_upload"):
                                st.session_state.pop("_confirmar_upload_eap", None)
                                st.session_state.pop("_upload_eap_bytes_pending", None)
                                st.session_state.pop("_upload_eap_med_novo", None)
                                # Reset the uploader by cycling its key
                                _cur_k = st.session_state.get("_uploader_key", 0)
                                st.session_state["_uploader_key"] = _cur_k + 1
                                st.rerun()
                    else:
                        # Confirmado — executa a substituição
                        _bytes  = st.session_state.pop("_upload_eap_bytes_pending", file_bytes)
                        _med    = st.session_state.pop("_upload_eap_med_novo", med_novo)
                        st.session_state.pop("_confirmar_upload_eap", None)
                        _dest   = DATA_DIR / f"eap_medida{_med}.xlsx"
                        try:
                            with open(_dest, "wb") as fh:
                                fh.write(_bytes)
                            with open(DEFAULT_EAP_FILE, "wb") as fh:
                                fh.write(_bytes)
                            load_from_path.clear()
                            load_from_bytes.clear()
                            st.session_state["arquivo_eap_bytes"] = _bytes
                            _cur_k2 = st.session_state.get("_uploader_key", 0)
                            st.session_state["_uploader_key"] = _cur_k2 + 1
                            st.session_state["_upload_sucesso_msg"] = f"✓ {_med}ª Medição substituída com sucesso por {usuario_atual}. Arquivo: data/eap_medida{_med}.xlsx"
                            st.rerun()
                        except Exception as exc:
                            st.error(f"Erro ao salvar arquivo: {exc}")

                else:
                    # Nova medição — botão direto
                    if st.button(f"⬆ Salvar {med_novo}ª Medição", type="primary", use_container_width=False, key="btn_upload_nova"):
                        try:
                            with open(dest_file, "wb") as fh:
                                fh.write(file_bytes)
                            with open(DEFAULT_EAP_FILE, "wb") as fh:
                                fh.write(file_bytes)
                            load_from_path.clear()
                            load_from_bytes.clear()
                            st.session_state["arquivo_eap_bytes"] = file_bytes
                            _cur_k3 = st.session_state.get("_uploader_key", 0)
                            st.session_state["_uploader_key"] = _cur_k3 + 1
                            st.session_state["_upload_sucesso_msg"] = f"✓ {med_novo}ª Medição salva com sucesso por {usuario_atual}. Arquivo: data/eap_medida{med_novo}.xlsx"
                            st.rerun()
                        except Exception as exc:
                            st.error(f"Erro ao salvar arquivo: {exc}")
        else:
            # Mostra arquivo atual em uso com info da medição
            _med_info = ""
            if DEFAULT_EAP_FILE.exists():
                _nome_arq = DEFAULT_EAP_FILE.name
                try:
                    _d_cur = load_from_path(str(DEFAULT_EAP_FILE))
                    _med_cur = int(_d_cur["medicoes"]["medicao_atual"])
                    _med_info = f" &nbsp;·&nbsp; <strong>{_med_cur}ª Medição</strong>"
                except Exception:
                    pass
            elif ROOT_EAP_FILE.exists():
                _nome_arq = ROOT_EAP_FILE.name
            else:
                _nome_arq = None

            if _nome_arq:
                html(f"""
                <div style="max-width:560px;background:#EBF0F8;border:1px solid #C8D8EE;border-radius:8px;
                            padding:10px 14px;font-size:13px;color:#1B3A5C;margin-top:4px;">
                    📄 Arquivo atual em uso: <strong>{_nome_arq}</strong>{_med_info}
                </div>
                """)
            else:
                html("""
                <div style="max-width:560px;background:#FEF3DC;border:1px solid #F0D08A;border-radius:8px;
                            padding:10px 14px;font-size:13px;color:#8A5300;margin-top:4px;">
                    ⚠ Nenhum arquivo EAP encontrado.
                </div>
                """)

    # ══════════════════════════════════════════════════════════════════════
    # SUBABA PROJETOS — Registro e Histórico (data editors)
    # ══════════════════════════════════════════════════════════════════════
    with aba_projetos:
        med_atual_num = data["medicoes"]["medicao_atual"]

        card(
            "Registro por Medição — Avanço dos Projetos",
            f"""
            <p style="font-size:13px;color:var(--text2);line-height:1.7;margin:0;">
                Edite diretamente as tabelas abaixo. Use os editores para adicionar, editar ou excluir registros.
                Os dados são salvos em <strong>data/registros_projetos.xlsx</strong>.
                Medição atual carregada: <strong>{med_atual_num}ª</strong>.
            </p>
            """
        )
        spacer(16)

        # ══════════════════════════════════════════════════════════════════
        # DATA EDITOR — RESUMO
        # ══════════════════════════════════════════════════════════════════
        import io as _io

        df_hist = _carregar_registros_projetos()

        # Download button
        col_dl, col_clr, _esp = st.columns([1, 1, 4])
        with col_dl:
            _buf_dl = _io.BytesIO()
            _df_pags_dl = _carregar_pagamentos_projetos()
            with pd.ExcelWriter(_buf_dl, engine="openpyxl") as _w:
                df_hist.to_excel(_w, sheet_name="Resumo", index=False)
                _df_pags_dl.to_excel(_w, sheet_name="Pagamentos", index=False)
            st.download_button(
                "⬇ Baixar Excel",
                data=_buf_dl.getvalue(),
                file_name="registros_projetos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_proj_excel",
            )
        with col_clr:
            if st.button("🗑 Limpar Histórico", key="btn_clear_hist2", use_container_width=True):
                st.session_state["_confirm_clear_hist"] = True

        if st.session_state.get("_confirm_clear_hist"):
            st.warning("⚠ Tem certeza? Isso apagará **todo** o histórico de registros e pagamentos.")
            _cc1, _cc2, _ = st.columns([1, 1, 4])
            with _cc1:
                if st.button("✓ Sim, apagar tudo", type="primary", use_container_width=True, key="conf_clear_hist"):
                    try:
                        _projetos_excel_path().unlink(missing_ok=True)
                        st.session_state.pop("_confirm_clear_hist", None)
                        st.success("Histórico apagado.")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Erro: {exc}")
            with _cc2:
                if st.button("✕ Cancelar", use_container_width=True, key="canc_clear_hist"):
                    st.session_state.pop("_confirm_clear_hist", None)
                    st.rerun()

        spacer(12)
        st.caption("RESUMO POR MEDIÇÃO")

        # Prepare Resumo dataframe for editor
        _RESUMO_COLS = ["medicao", "data_registro", "avanco_geral_pct",
                        "montante_total", "saldo_devedor", "status_geral", "observacoes"]
        import datetime as _dt2
        _RESUMO_EMPTY = {
            "medicao": 1, "data_registro": _dt2.date.today(),
            "avanco_geral_pct": 0.0, "montante_total": 0.0,
            "saldo_devedor": 0.0, "status_geral": "Pendente", "observacoes": "",
        }

        if len(df_hist) > 0:
            df_resumo_edit = df_hist[_RESUMO_COLS].copy()
        else:
            df_resumo_edit = pd.DataFrame([_RESUMO_EMPTY])

        # Converte tipos para compatibilidade com os column_config
        def _to_float_br(v):
            """Converte string BR (vírgula decimal) ou ponto decimal para float."""
            if v is None or (isinstance(v, float) and pd.isna(v)):
                return 0.0
            if isinstance(v, (int, float)):
                return float(v)
            s = str(v).strip()
            # Formato BR: 1.041.272,80 -> remover pontos de milhar, trocar vírgula por ponto
            if "," in s:
                s = s.replace(".", "").replace(",", ".")
            try:
                return float(s)
            except Exception:
                return 0.0
        df_resumo_edit["medicao"] = pd.to_numeric(df_resumo_edit["medicao"], errors="coerce").fillna(1).astype(int)
        df_resumo_edit["avanco_geral_pct"] = df_resumo_edit["avanco_geral_pct"].apply(_to_float_br)
        df_resumo_edit["montante_total"]   = df_resumo_edit["montante_total"].apply(_to_float_br)
        df_resumo_edit["saldo_devedor"]    = df_resumo_edit["saldo_devedor"].apply(_to_float_br)
        # Recalcula avanco e saldo a partir dos pagamentos salvos
        _df_pags_calc = _carregar_pagamentos_projetos()
        for _idx, _row in df_resumo_edit.iterrows():
            _m = str(int(_row["medicao"]))
            _mont = float(_row["montante_total"])
            if len(_df_pags_calc) > 0 and "medicao" in _df_pags_calc.columns:
                _mp = _df_pags_calc[_df_pags_calc["medicao"].astype(str) == _m]
                _pcts = pd.to_numeric(_mp["pct"], errors="coerce").dropna().tolist()
                _vals = pd.to_numeric(_mp["valor"], errors="coerce").fillna(0).tolist()
                _avanco = round(sum(_pcts) / len(_pcts), 1) if _pcts else 0.0
                _pago = sum(_vals)
            else:
                _avanco = 0.0
                _pago = 0.0
            df_resumo_edit.at[_idx, "avanco_geral_pct"] = _avanco
            df_resumo_edit.at[_idx, "saldo_devedor"] = max(_mont - _pago, 0.0)
        # Converte data_registro para datetime.date para DateColumn
        import datetime as _dt2
        def _parse_date(v):
            if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "":
                return _dt2.date.today()
            for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d"):
                try:
                    return _dt2.datetime.strptime(str(v)[:len(fmt.replace("%d","00").replace("%m","00").replace("%Y","0000").replace("%H","00").replace("%M","00"))], fmt).date()
                except Exception:
                    pass
            try:
                return pd.to_datetime(str(v)).date()
            except Exception:
                return _dt2.date.today()
        df_resumo_edit["data_registro"] = df_resumo_edit["data_registro"].apply(_parse_date)
        df_resumo_edit["status_geral"] = df_resumo_edit["status_geral"].fillna("Pendente").replace("", "Pendente").astype(str)
        df_resumo_edit["observacoes"]  = df_resumo_edit["observacoes"].fillna("").astype(str)

        # Formata colunas numericas como texto para alinhar a esquerda (canvas nao aceita CSS)
        df_resumo_disp = df_resumo_edit.copy()
        df_resumo_disp["medicao"]          = df_resumo_disp["medicao"].astype(str)
        df_resumo_disp["avanco_geral_pct"] = df_resumo_disp["avanco_geral_pct"].apply(lambda v: f"{v:.1f}")
        df_resumo_disp["montante_total"]   = df_resumo_disp["montante_total"].apply(lambda v: f"R$ {v:,.2f}".replace(",","X").replace(".",",").replace("X","."))
        df_resumo_disp["saldo_devedor"]    = df_resumo_disp["saldo_devedor"].apply(lambda v: f"R$ {v:,.2f}".replace(",","X").replace(".",",").replace("X","."))
        df_resumo_disp["data_registro"]    = df_resumo_disp["data_registro"].apply(lambda v: v.strftime("%d/%m/%Y") if hasattr(v,"strftime") else str(v))

        edited_resumo = st.data_editor(
            df_resumo_disp,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            key="de_resumo2",
            column_config={
                "medicao":          st.column_config.TextColumn("Med.", width="small"),
                "data_registro":    st.column_config.TextColumn("Data", width="medium"),
                "avanco_geral_pct": st.column_config.TextColumn("Avanco %", width="small"),
                "montante_total":   st.column_config.TextColumn("Montante (R$)", width="medium"),
                "saldo_devedor":    st.column_config.TextColumn("Saldo (R$)", width="medium"),
                "status_geral":     st.column_config.SelectboxColumn("Status", width="small", options=["Pago", "Pendente", "Nao Pago"], required=True),
                "observacoes":      st.column_config.TextColumn("Observacoes", width="large"),
            },
        )

        _sr1, _sr2, _ = st.columns([1, 1, 4])
        with _sr1:
            if st.button("💾 Salvar Resumo", type="primary", use_container_width=True, key="btn_save_resumo_de"):
                st.session_state["_confirm_save_resumo"] = True
        if st.session_state.get("_confirm_save_resumo"):
            st.warning("⚠ Tem certeza que deseja salvar as alterações no Resumo?")
            _cs1, _cs2, _ = st.columns([1, 1, 4])
            with _cs1:
                if st.button("✓ Confirmar", type="primary", use_container_width=True, key="conf_save_resumo"):
                    try:
                        _df_pags_cur = _carregar_pagamentos_projetos()
                        _df_r_save = edited_resumo.copy()
                        for col in _RESUMO_COLS:
                            if col not in _df_r_save.columns:
                                _df_r_save[col] = ""
                        # Converte data de volta para string
                        _df_r_save["data_registro"] = _df_r_save["data_registro"].apply(
                            lambda v: v.strftime("%d/%m/%Y") if hasattr(v, "strftime") else str(v)
                        )
                        # Garante float nos campos numericos (converte strings BR: "R$ 1.234,56" -> 1234.56)
                        def _br_to_float(v):
                            if v is None or (isinstance(v, float) and pd.isna(v)): return 0.0
                            s = str(v).strip().replace("R$","").strip()
                            if "," in s: s = s.replace(".","").replace(",",".")
                            try: return float(s)
                            except: return 0.0
                        _df_r_save["montante_total"]   = _df_r_save["montante_total"].apply(_br_to_float)
                        _df_r_save["saldo_devedor"]    = _df_r_save["saldo_devedor"].apply(_br_to_float)
                        _df_r_save["avanco_geral_pct"] = _df_r_save["avanco_geral_pct"].apply(_br_to_float)
                        # Preenche status vazio
                        _df_r_save["status_geral"] = _df_r_save["status_geral"].fillna("Pendente").replace("", "Pendente")
                        with pd.ExcelWriter(_projetos_excel_path(), engine="openpyxl") as _w:
                            _df_r_save[_RESUMO_COLS].to_excel(_w, sheet_name="Resumo", index=False)
                            _df_pags_cur.to_excel(_w, sheet_name="Pagamentos", index=False)
                        st.session_state.pop("_confirm_save_resumo", None)
                        st.success("✓ Resumo salvo.")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Erro: {exc}")
            with _cs2:
                if st.button("✕ Cancelar", use_container_width=True, key="canc_save_resumo"):
                    st.session_state.pop("_confirm_save_resumo", None)
                    st.rerun()

        st.markdown(
            f"<div style='font-size:11px;color:var(--text3);margin-top:6px;'>"
            f"{len(df_hist)} registro(s) · data/registros_projetos.xlsx</div>",
            unsafe_allow_html=True,
        )

        spacer(24)

        # ══════════════════════════════════════════════════════════════════
        # DATA EDITOR — PAGAMENTOS
        # ══════════════════════════════════════════════════════════════════
        st.caption("PAGAMENTOS POR ETAPA")

        _PAGS_COLS = ["medicao", "etapa", "pct", "valor", "status"]
        _PAGS_EMPTY = {"medicao": 1, "etapa": "", "pct": 0.0, "valor": 0.0, "status": "Pendente"}

        df_pags_hist = _carregar_pagamentos_projetos()
        if len(df_pags_hist) > 0:
            df_pags_edit = df_pags_hist[_PAGS_COLS].copy()
        else:
            df_pags_edit = pd.DataFrame([_PAGS_EMPTY])

        def _to_float_br2(v):
            if v is None or (isinstance(v, float) and pd.isna(v)):
                return 0.0
            if isinstance(v, (int, float)):
                return float(v)
            s = str(v).strip()
            if "," in s:
                s = s.replace(".", "").replace(",", ".")
            try:
                return float(s)
            except Exception:
                return 0.0
        df_pags_edit["medicao"] = pd.to_numeric(df_pags_edit["medicao"], errors="coerce").fillna(1).astype(int)
        df_pags_edit["pct"]     = df_pags_edit["pct"].apply(_to_float_br2)
        df_pags_edit["valor"]   = df_pags_edit["valor"].apply(_to_float_br2)
        df_pags_edit["status"]  = df_pags_edit["status"].astype(str)
        df_pags_edit["etapa"]   = df_pags_edit["etapa"].astype(str)

        # Formata colunas numericas como texto para alinhar a esquerda
        df_pags_disp = df_pags_edit.copy()
        df_pags_disp["medicao"] = df_pags_disp["medicao"].astype(str)
        df_pags_disp["pct"]     = df_pags_disp["pct"].apply(lambda v: f"{v:.1f}")
        df_pags_disp["valor"]   = df_pags_disp["valor"].apply(lambda v: f"R$ {v:,.2f}".replace(",","X").replace(".",",").replace("X","."))

        edited_pags = st.data_editor(
            df_pags_disp,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            key="de_pags2",
            column_config={
                "medicao": st.column_config.TextColumn("Med.", width="small"),
                "etapa":   st.column_config.TextColumn("Etapa / Projeto", width="large"),
                "pct":     st.column_config.TextColumn("% Avanço", width="small"),
                "valor":   st.column_config.TextColumn("Valor (R$)", width="medium"),
                "status":  st.column_config.SelectboxColumn("Status", width="small", options=["Pago", "Pendente", "Nao Pago"], required=True),
            },
        )

        _pp1, _pp2, _ = st.columns([1, 1, 4])
        with _pp1:
            if st.button("💾 Salvar Pagamentos", type="primary", use_container_width=True, key="btn_save_pags_de"):
                st.session_state["_confirm_save_pags"] = True
        if st.session_state.get("_confirm_save_pags"):
            st.warning("⚠ Tem certeza que deseja salvar as alterações nos Pagamentos?")
            _cp1, _cp2, _ = st.columns([1, 1, 4])
            with _cp1:
                if st.button("✓ Confirmar", type="primary", use_container_width=True, key="conf_save_pags"):
                    try:
                        _df_ps = edited_pags.copy()
                        _df_ps["medicao"] = _df_ps["medicao"].astype(str)
                        def _br2f(v):
                            if v is None or (isinstance(v, float) and pd.isna(v)): return 0.0
                            s = str(v).strip().replace("R$","").strip()
                            if "," in s: s = s.replace(".","").replace(",",".")
                            try: return float(s)
                            except: return 0.0
                        _df_ps["pct"]   = _df_ps["pct"].apply(_br2f)
                        _df_ps["valor"] = _df_ps["valor"].apply(_br2f)
                        def _ss3(s): return {"Pago":1.0,"Pendente":0.5,"Nao Pago":0.0,"Não Pago":0.0}.get(str(s),0.5)
                        _df_rs3 = _carregar_registros_projetos()
                        for _ms in _df_ps["medicao"].unique():
                            _sc = [_ss3(r) for r in _df_ps[_df_ps["medicao"]==_ms]["status"].tolist()]
                            _sg = "Pago" if (sum(_sc)/len(_sc) if _sc else 0)>=0.9 else ("Pendente" if (sum(_sc)/len(_sc) if _sc else 0)>=0.4 else "Não Pago")
                            if len(_df_rs3)>0 and "medicao" in _df_rs3.columns:
                                _df_rs3.loc[_df_rs3["medicao"].astype(str)==_ms,"status_geral"] = _sg
                        with pd.ExcelWriter(_projetos_excel_path(), engine="openpyxl") as _w:
                            _df_rs3.to_excel(_w, sheet_name="Resumo", index=False)
                            _df_ps.to_excel(_w, sheet_name="Pagamentos", index=False)
                        st.session_state.pop("_confirm_save_pags", None)
                        st.success("✓ Pagamentos salvos.")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Erro: {exc}")
            with _cp2:
                if st.button("✕ Cancelar", use_container_width=True, key="canc_save_pags"):
                    st.session_state.pop("_confirm_save_pags", None)
                    st.rerun()




# =============================================================================
# APP
# =============================================================================
inject_css()

try:
    data = load_current_data()
except Exception as error:
    st.error(f"Não foi possível carregar os dados do Excel: {error}")
    st.stop()

render_header(data)

aba_visao, aba_fisico, aba_financeiro, aba_prazos, aba_contrato, aba_upload = st.tabs(ABAS)

with aba_visao:
    tab_visao(data)
    watermark()

with aba_fisico:
    tab_fisico(data)
    watermark()

with aba_financeiro:
    tab_financeiro(data)
    watermark()

with aba_prazos:
    tab_prazos(data)
    watermark()

with aba_contrato:
    tab_contrato(data)
    watermark()

with aba_upload:
    tab_upload(data)
    watermark()