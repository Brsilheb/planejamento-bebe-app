# app_v2.py
# Streamlit interface – Planejamento

import os
import io
import math
from datetime import datetime, date, timedelta
from typing import Dict

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Planejamento", layout="wide")

DEFAULT_FILE   = "Planejamento_Bebe_Heber.xlsx"
SHEET_ITENS    = "01_Itens"
SHEET_PLANO    = "02_Plano_Semanal"
SHEET_ORC      = "03_Orcamento"
SHEET_CONFIG   = "04_Config"

DPP_DEFAULT       = date(2026, 9, 25)
SEMANA_DEFAULT    = 22
DATA_BASE_DEFAULT = date(2026, 5, 22)

PRIORIDADES = ["Alta", "Média", "Baixa"]
STATUS      = ["A comprar", "Cotar", "Comprado", "Recebido"]

ITENS_COLUMNS = [
    "ID", "Categoria", "Item", "Prioridade", "Qtd", "Unidade",
    "Marca/Modelo", "Loja/Link", "Preço Unitário (R$)", "Custo Estimado (R$)",
    "Total (R$)", "Status", "Data da Compra", "Garantia (meses)", "Notas/Observações"
]

# -----------------------------
# Dados: desenvolvimento semanal
# -----------------------------

DESENVOLVIMENTO_BEBE = {
    22: {
        "fruta": "mamão papaia",
        "comprimento": "27 cm",
        "peso": "430 g",
        "desenvolvimento": (
            "Audição bem desenvolvida — o bebê já ouve a sua voz! "
            "Cérebro em crescimento acelerado. Lábios, olhos e sobrancelhas bem formados."
        ),
    },
    23: {
        "fruta": "manga",
        "comprimento": "28 cm",
        "peso": "500 g",
        "desenvolvimento": (
            "Pulmões em desenvolvimento — produz surfactante. "
            "Movimentos mais coordenados. Pele ainda translúcida e enrugada."
        ),
    },
    24: {
        "fruta": "espiga de milho",
        "comprimento": "30 cm",
        "peso": "600 g",
        "desenvolvimento": (
            "Marco de viabilidade fetal! Rosto quase completamente formado. "
            "Papilas gustativas desenvolvidas."
        ),
    },
    25: {
        "fruta": "couve-nabo",
        "comprimento": "34 cm",
        "peso": "700 g",
        "desenvolvimento": (
            "Responde melhor a sons externos. Capilares se formando sob a pele. "
            "Narinas começam a funcionar."
        ),
    },
    26: {
        "fruta": "alface americana",
        "comprimento": "35 cm",
        "peso": "760 g",
        "desenvolvimento": (
            "Olhos começam a se abrir. Atividade cerebral intensa. "
            "Gordura subcutânea em formação."
        ),
    },
    27: {
        "fruta": "couve-flor",
        "comprimento": "36 cm",
        "peso": "875 g",
        "desenvolvimento": (
            "Último dia do 2° trimestre em breve! Padrões de sono e vigília estabelecidos. "
            "Cérebro com mais sulcos e dobras."
        ),
    },
    28: {
        "fruta": "berinjela",
        "comprimento": "37 cm",
        "peso": "1,0 kg",
        "desenvolvimento": (
            "Início do 3° trimestre! Olhos totalmente abertos. "
            "Fase de engorda intensa. Sonhos possíveis (fase REM ativa)."
        ),
    },
    29: {
        "fruta": "abóbora japonesa pequena",
        "comprimento": "38 cm",
        "peso": "1,2 kg",
        "desenvolvimento": (
            "Ossos endurecendo. Músculos e gordura crescendo rapidamente. "
            "Chutes mais fortes e visíveis."
        ),
    },
    30: {
        "fruta": "repolho",
        "comprimento": "40 cm",
        "peso": "1,3 kg",
        "desenvolvimento": (
            "Cérebro controlando temperatura corporal. "
            "Volume de líquido amniótico no pico. Unhas dos pés crescendo."
        ),
    },
    31: {
        "fruta": "coco",
        "comprimento": "41 cm",
        "peso": "1,5 kg",
        "desenvolvimento": (
            "Todos os sentidos operantes. Pulmões quase maduros. "
            "Processamento de informações do cérebro acelerado."
        ),
    },
    32: {
        "fruta": "nabo grande",
        "comprimento": "42 cm",
        "peso": "1,7 kg",
        "desenvolvimento": (
            "Pele mais lisa com gordura acumulada. "
            "Ossos do crânio ainda maleáveis (essencial para o parto). Unhas completas."
        ),
    },
    33: {
        "fruta": "abacaxi",
        "comprimento": "43 cm",
        "peso": "1,9 kg",
        "desenvolvimento": (
            "Imunidade passando da mãe para o bebê via placenta. "
            "Ossos endurecendo — exceto crânio. Pulmões quase prontos."
        ),
    },
    34: {
        "fruta": "melão cantaloupe",
        "comprimento": "45 cm",
        "peso": "2,1 kg",
        "desenvolvimento": (
            "Sistema nervoso central maduro. "
            "Maioria dos bebês já em posição cefálica. Gordura aumentando rapidamente."
        ),
    },
    35: {
        "fruta": "melão médio",
        "comprimento": "46 cm",
        "peso": "2,4 kg",
        "desenvolvimento": (
            "Rins totalmente funcionais. Fígado processa resíduos. "
            "Bebê ocupa quase todo o útero — movimentos mais limitados."
        ),
    },
    36: {
        "fruta": "alface romana",
        "comprimento": "47 cm",
        "peso": "2,6 kg",
        "desenvolvimento": (
            "Perda do lanugo (pelos finos). Gordura corporal ~15%. "
            "A partir da semana 37 é considerado a termo."
        ),
    },
    37: {
        "fruta": "acelga grande",
        "comprimento": "48 cm",
        "peso": "2,9 kg",
        "desenvolvimento": (
            "Bebê a termo! Pulmões maduros. "
            "Reflexos de sucção e deglutição prontos para amamentação."
        ),
    },
    38: {
        "fruta": "alho-poró grande",
        "comprimento": "49 cm",
        "peso": "3,1 kg",
        "desenvolvimento": (
            "Órgãos totalmente desenvolvidos. "
            "Vérnix (camada protetora) diminuindo. Pronto para nascer."
        ),
    },
    39: {
        "fruta": "melancia pequena",
        "comprimento": "50 cm",
        "peso": "3,3 kg",
        "desenvolvimento": (
            "Em posição de parto. Hormônios maternos preparando para o trabalho de parto. "
            "Placenta ainda nutrida."
        ),
    },
    40: {
        "fruta": "abóbora pequena",
        "comprimento": "51 cm",
        "peso": "3,5 kg",
        "desenvolvimento": (
            "DPP chegou! Bebê completamente desenvolvido. "
            "Pronto para vir ao mundo — cada dia é de espera ansiosa."
        ),
    },
}

# -----------------------------
# Dados: enxoval padrão (70 itens)
# (Categoria, Item, Prioridade, Qtd, Unidade, Preço Unitário, Notas)
# -----------------------------

DEFAULT_ENXOVAL = [
    # ── ROUPAS (10) ──────────────────────────────────────────────────────────
    ("Roupas", "Saída de maternidade (kit)",        "Alta",  1, "kit",  120.0, "Kit com manta, body e touquinha"),
    ("Roupas", "Body manga curta RN",               "Baixa", 5, "un",    35.0, "Tamanho RN — comprar no 3° tri"),
    ("Roupas", "Body manga longa RN",               "Baixa", 4, "un",    40.0, "Para noites mais frias"),
    ("Roupas", "Macacão com pé (tamanho P)",        "Baixa", 3, "un",    55.0, ""),
    ("Roupas", "Macacão sem pé (tamanho P)",        "Baixa", 3, "un",    45.0, ""),
    ("Roupas", "Meias antiderrapantes bebê",        "Baixa", 6, "par",   15.0, ""),
    ("Roupas", "Touquinha de algodão",              "Baixa", 3, "un",    18.0, ""),
    ("Roupas", "Luvas anti-arranhão",               "Baixa", 3, "par",   12.0, "Evita arranhões no rosto"),
    ("Roupas", "Mijão (pijama com pé)",             "Baixa", 3, "un",    60.0, ""),
    ("Roupas", "Casaco de plush",                   "Baixa", 2, "un",    80.0, "Para dias frios"),

    # ── HIGIENE E CUIDADOS (12) ───────────────────────────────────────────────
    ("Higiene e Cuidados", "Fraldas descartáveis RN",        "Alta",  4, "pct",   45.0, "Pacote c/ ~70 fraldas"),
    ("Higiene e Cuidados", "Fraldas descartáveis P",         "Alta",  4, "pct",   50.0, "Pacote c/ ~70 fraldas"),
    ("Higiene e Cuidados", "Lenços umedecidos s/ álcool",    "Alta",  6, "pct",   22.0, "Sem perfume e álcool"),
    ("Higiene e Cuidados", "Pomada para assaduras",          "Alta",  3, "un",    30.0, "Bepantol Baby ou similar"),
    ("Higiene e Cuidados", "Termômetro digital",             "Alta",  1, "un",    60.0, "Axilar ou auricular"),
    ("Higiene e Cuidados", "Aspirador nasal",                "Alta",  1, "un",    45.0, "Manual tipo NoseFrida ou elétrico"),
    ("Higiene e Cuidados", "Sabonete líquido bebê",          "Média", 2, "un",    25.0, "Hipoalergênico"),
    ("Higiene e Cuidados", "Shampoo bebê",                   "Média", 2, "un",    25.0, "Sem lágrimas"),
    ("Higiene e Cuidados", "Óleo de amêndoas / hidratante",  "Média", 1, "un",    35.0, ""),
    ("Higiene e Cuidados", "Algodão hidrófilo",              "Média", 2, "pct",   12.0, ""),
    ("Higiene e Cuidados", "Tesoura/cortador de unhas",      "Média", 1, "un",    30.0, "Específico para bebê"),
    ("Higiene e Cuidados", "Cotonetes para bebê",            "Baixa", 1, "caixa", 15.0, "Com proteção na ponta"),

    # ── ALIMENTAÇÃO (10) ──────────────────────────────────────────────────────
    ("Alimentação", "Almofada de amamentação",       "Alta",  1, "un",   120.0, "Formato C ou U"),
    ("Alimentação", "Sutiã para amamentação",        "Alta",  3, "un",    80.0, "Comprar perto do parto — tamanho muda"),
    ("Alimentação", "Protetor de seios (disco)",     "Alta",  1, "caixa", 40.0, "Descartável ou lavável"),
    ("Alimentação", "Bomba tira-leite elétrica",     "Média", 1, "un",   450.0, "Medela Harmony / Philips Avent"),
    ("Alimentação", "Mamadeira 150 ml anti-cólica",  "Média", 2, "un",    65.0, "NUK ou MAM"),
    ("Alimentação", "Mamadeira 240 ml anti-cólica",  "Média", 2, "un",    70.0, "NUK ou MAM"),
    ("Alimentação", "Esterilizador elétrico",        "Média", 1, "un",   250.0, "Capacidade 6 mamadeiras"),
    ("Alimentação", "Babador impermeável",            "Baixa", 5, "un",    20.0, ""),
    ("Alimentação", "Copo de treinamento",            "Baixa", 1, "un",    40.0, "Para introdução alimentar ~6m"),
    ("Alimentação", "Colheres de silicone",           "Baixa", 2, "un",    18.0, "Para papinhas — a partir dos 6 meses"),

    # ── MÓVEIS E QUARTO (10) ──────────────────────────────────────────────────
    ("Móveis e Quarto", "Berço grade fixa",           "Alta",  1, "un",  700.0, "Com certificação INMETRO"),
    ("Móveis e Quarto", "Colchão para berço",         "Alta",  1, "un",  250.0, "Espuma D18 ou superior"),
    ("Móveis e Quarto", "Trocador de fraldas",        "Alta",  1, "un",  200.0, "Com bordas de segurança"),
    ("Móveis e Quarto", "Cômoda para quarto bebê",    "Alta",  1, "un",  900.0, "Com ou sem trocador acoplado"),
    ("Móveis e Quarto", "Kit berço (6 peças)",        "Alta",  1, "kit", 280.0, "Fronha, protetor, saia, lençol — sem itens soltos no berço"),
    ("Móveis e Quarto", "Luminária noturna",          "Média", 1, "un",   80.0, "Luz âmbar — não perturba sono"),
    ("Móveis e Quarto", "Monitor de bebê com câmera", "Média", 1, "un",  350.0, ""),
    ("Móveis e Quarto", "Manta para berço",           "Média", 2, "un",   60.0, "Fina, algodão — não colocar no berço RN"),
    ("Móveis e Quarto", "Mobile musical",             "Baixa", 1, "un",  120.0, "Estimulação visual e auditiva"),
    ("Móveis e Quarto", "Umidificador de ar",         "Baixa", 1, "un",  150.0, "Qualidade do ar no quarto"),

    # ── TRANSPORTE (3) ────────────────────────────────────────────────────────
    ("Transporte", "Bebê conforto (cadeirinha 0+)", "Alta",  1, "un", 800.0,  "Grupo 0+, até 13 kg — INMETRO obrigatório"),
    ("Transporte", "Carrinho de bebê",              "Alta",  1, "un", 1200.0, "Reclinável plano para RN"),
    ("Transporte", "Mochila canguru ergonômica",    "Média", 1, "un",  180.0, "Frente e costas, até 20 kg"),

    # ── BANHO (5) ─────────────────────────────────────────────────────────────
    ("Banho", "Banheira com suporte/redutor",   "Alta",  1, "un",  180.0, "Com redutor para RN"),
    ("Banho", "Termômetro de banho",            "Alta",  1, "un",   35.0, "Manter água a ~37 °C"),
    ("Banho", "Toalha com capuz",               "Média", 3, "un",   70.0, "100% algodão"),
    ("Banho", "Esponja de banho bebê",          "Média", 2, "un",   20.0, "Macia"),
    ("Banho", "Tina estilo balde (Puj/Stokke)", "Baixa", 1, "un",  220.0, "Opcional — banho de balde"),

    # ── SAÚDE E SEGURANÇA (5) ────────────────────────────────────────────────
    ("Saúde e Segurança", "Protetor de tomadas",              "Alta", 20, "un",    3.0, "Instalar antes de engatinhar"),
    ("Saúde e Segurança", "Tapete antiderrapante banheiro",   "Alta",  1, "un",   45.0, ""),
    ("Saúde e Segurança", "Termômetro de ambiente",           "Alta",  1, "un",   50.0, "Manter quarto a 20-22 °C"),
    ("Saúde e Segurança", "Kit primeiros socorros",           "Alta",  1, "kit",  80.0, "Gazes, soro fisiológico, esparadrapo"),
    ("Saúde e Segurança", "Cinto de segurança trocador",      "Alta",  1, "un",   25.0, "Nunca deixar bebê sem cinto"),

    # ── MATERNIDADE E MÃE (7) ────────────────────────────────────────────────
    ("Maternidade e Mãe", "Bolsa maternidade",               "Alta",  1, "un",  200.0, "Grande, com muitos compartimentos"),
    ("Maternidade e Mãe", "Absorvente pós-parto noturno",    "Alta",  2, "pct",  25.0, "Extra-absorção"),
    ("Maternidade e Mãe", "Calcinha descartável pós-parto",  "Alta",  2, "pct",  30.0, "Kit 5 un."),
    ("Maternidade e Mãe", "Cinta pós-parto",                "Média", 1, "un",   90.0, "Usar após orientação médica"),
    ("Maternidade e Mãe", "Sutiã amamentação para dormir",   "Média", 2, "un",   60.0, "Confortável, sem arame"),
    ("Maternidade e Mãe", "Travesseiro de amamentação",      "Média", 1, "un",  120.0, "Pode ser a almofada de amamentação"),
    ("Maternidade e Mãe", "Creme preventivo de estrias",     "Baixa", 2, "un",   50.0, "Usar já agora durante a gestação"),

    # ── DOCUMENTOS E SERVIÇOS (3) ────────────────────────────────────────────
    ("Documentos e Serviços", "Plano de saúde infantil (cotação)", "Alta", 1, "un", 400.0, "Incluir bebê em até 30 dias após o nascimento"),
    ("Documentos e Serviços", "Certidão de nascimento",            "Alta", 1, "un",   0.0, "Registrar em cartório em até 15 dias — gratuito"),
    ("Documentos e Serviços", "Caderneta de vacinação",            "Alta", 1, "un",   0.0, "Gratuita no SUS — disponível na maternidade"),

    # ── BRINQUEDOS E ESTÍMULOS (5) ───────────────────────────────────────────
    ("Brinquedos e Estímulos", "Chupeta ortodôntica 0-6 m",  "Média", 3, "un",   30.0, "INMETRO, ortodôntica"),
    ("Brinquedos e Estímulos", "Chocalho colorido",          "Baixa", 2, "un",   30.0, "Estimulação visual e auditiva"),
    ("Brinquedos e Estímulos", "Tapete de atividades",       "Baixa", 1, "un",  150.0, "Com arcos e penduricalhos"),
    ("Brinquedos e Estímulos", "Livros de pano",             "Baixa", 3, "un",   35.0, "Estimulação sensorial"),
    ("Brinquedos e Estímulos", "Mordedor de silicone",       "Baixa", 2, "un",   25.0, "Para fase de dentição"),
]

# Orçamento padrão (Junho a Outubro 2026, R$ 1.000/mês)
DEFAULT_ORCAMENTO = [
    {"Mês": "2026-06", "Orçado (R$)": 1000.0, "Previsto (R$)": 1000.0, "Realizado (R$)": 0.0, "Diferença (R$)": 1000.0},
    {"Mês": "2026-07", "Orçado (R$)": 1000.0, "Previsto (R$)": 1000.0, "Realizado (R$)": 0.0, "Diferença (R$)": 1000.0},
    {"Mês": "2026-08", "Orçado (R$)": 1000.0, "Previsto (R$)": 1000.0, "Realizado (R$)": 0.0, "Diferença (R$)": 1000.0},
    {"Mês": "2026-09", "Orçado (R$)": 1000.0, "Previsto (R$)": 1000.0, "Realizado (R$)": 0.0, "Diferença (R$)": 1000.0},
    {"Mês": "2026-10", "Orçado (R$)": 1000.0, "Previsto (R$)": 1000.0, "Realizado (R$)": 0.0, "Diferença (R$)": 1000.0},
]

# -----------------------------
# Helpers
# -----------------------------

def _ensure_numeric(df: pd.DataFrame, cols):
    for c in cols:
        df[c] = pd.to_numeric(df.get(c, 0), errors="coerce").fillna(0.0)
    return df


def _normalize_id_column(df: pd.DataFrame) -> pd.DataFrame:
    if "ID" not in df.columns:
        df.insert(0, "ID", range(1, len(df) + 1))
    else:
        df["ID"] = pd.to_numeric(df["ID"], errors="coerce").fillna(0).astype(int)
        if (df["ID"] == 0).any() or df["ID"].duplicated().any():
            df = df.drop(columns=["ID"]).reset_index(drop=True)
            df.insert(0, "ID", range(1, len(df) + 1))
    return df


def _criar_df_enxoval() -> pd.DataFrame:
    rows = []
    for i, (cat, item, pri, qtd, und, preco, notas) in enumerate(DEFAULT_ENXOVAL, start=1):
        total = round(qtd * preco, 2)
        rows.append({
            "ID": i,
            "Categoria": cat,
            "Item": item,
            "Prioridade": pri,
            "Qtd": qtd,
            "Unidade": und,
            "Marca/Modelo": "",
            "Loja/Link": "",
            "Preço Unitário (R$)": preco,
            "Custo Estimado (R$)": total,
            "Total (R$)": total,
            "Status": "A comprar",
            "Data da Compra": None,
            "Garantia (meses)": 0,
            "Notas/Observações": notas,
        })
    return pd.DataFrame(rows, columns=ITENS_COLUMNS)


def gerar_sugestoes_faseado(df_sorted: pd.DataFrame, n_semanas: int) -> list:
    if n_semanas == 0:
        return []
    alta  = df_sorted[df_sorted["Prioridade"] == "Alta"].reset_index(drop=True)
    media = df_sorted[df_sorted["Prioridade"] == "Média"].reset_index(drop=True)
    baixa = df_sorted[df_sorted["Prioridade"] == "Baixa"].reset_index(drop=True)
    ordered = pd.concat([alta, media, baixa], ignore_index=True)
    items_per_week = max(3, math.ceil(len(ordered) / n_semanas))
    sugestoes = []
    idx = 0
    for _ in range(n_semanas):
        bloc = []
        for _ in range(items_per_week):
            if idx < len(ordered):
                r = ordered.iloc[idx]
                bloc.append(f"{r['Categoria']} – {r['Item']} ({r['Prioridade']})")
                idx += 1
        sugestoes.append("\n".join(bloc))
    return sugestoes


def carregar_config(dfs: Dict[str, pd.DataFrame]) -> dict:
    defaults = {
        "dpp": DPP_DEFAULT,
        "semana_base": SEMANA_DEFAULT,
        "data_base": DATA_BASE_DEFAULT,
    }
    if SHEET_CONFIG not in dfs or dfs[SHEET_CONFIG].empty:
        return defaults
    try:
        cfg_series = dfs[SHEET_CONFIG].set_index("Chave")["Valor"]
        if "dpp" in cfg_series.index:
            defaults["dpp"] = date.fromisoformat(str(cfg_series["dpp"]))
        if "semana_base" in cfg_series.index:
            defaults["semana_base"] = int(cfg_series["semana_base"])
        if "data_base" in cfg_series.index:
            defaults["data_base"] = date.fromisoformat(str(cfg_series["data_base"]))
    except Exception:
        pass
    return defaults


def salvar_config(dfs: Dict[str, pd.DataFrame], data_base: date, semana_base: int, dpp: date) -> None:
    dfs[SHEET_CONFIG] = pd.DataFrame({
        "Chave": ["dpp", "semana_base", "data_base"],
        "Valor": [dpp.isoformat(), str(semana_base), data_base.isoformat()],
    })


@st.cache_data(show_spinner=False)
def carregar_planilha(caminho: str) -> Dict[str, pd.DataFrame]:
    df_config_vazio = pd.DataFrame(columns=["Chave", "Valor"])

    if not os.path.exists(caminho):
        return {
            SHEET_ITENS:  pd.DataFrame(columns=ITENS_COLUMNS),
            SHEET_PLANO:  pd.DataFrame(columns=["Semana #", "Terça", "Itens sugeridos para cotar/comprar", "Status da semana", "Notas"]),
            SHEET_ORC:    pd.DataFrame(columns=["Mês", "Orçado (R$)", "Previsto (R$)", "Realizado (R$)", "Diferença (R$)"]),
            SHEET_CONFIG: df_config_vazio,
        }

    xls = pd.read_excel(caminho, sheet_name=None, engine="openpyxl")

    # --- ITENS ---
    df_itens = xls.get(SHEET_ITENS, pd.DataFrame())
    if df_itens is None:
        df_itens = pd.DataFrame()
    for c in ITENS_COLUMNS:
        if c not in df_itens.columns:
            df_itens[c] = '' if c not in ["Qtd", "Preço Unitário (R$)", "Custo Estimado (R$)", "Garantia (meses)", "Total (R$)"] else 0
    df_itens = _normalize_id_column(df_itens)
    df_itens["Prioridade"] = df_itens["Prioridade"].replace({None: ""})
    df_itens["Status"] = df_itens["Status"].replace({None: "A comprar"})
    for col in ["Qtd", "Preço Unitário (R$)", "Custo Estimado (R$)", "Garantia (meses)"]:
        df_itens[col] = pd.to_numeric(df_itens[col], errors="coerce").fillna(0)
    df_itens["Total (R$)"] = (df_itens["Qtd"].astype(float) * df_itens["Preço Unitário (R$)"].astype(float)).round(2)

    # --- PLANO ---
    df_plano = xls.get(SHEET_PLANO, pd.DataFrame())
    if df_plano is None:
        df_plano = pd.DataFrame()
    for c in ["Semana #", "Terça", "Itens sugeridos para cotar/comprar", "Status da semana", "Notas"]:
        if c not in df_plano.columns:
            df_plano[c] = "" if c != "Semana #" else 0
    if not df_plano.empty and "Terça" in df_plano.columns:
        try:
            df_plano["Terça"] = pd.to_datetime(df_plano["Terça"]).dt.date
        except Exception:
            pass

    # --- ORCAMENTO ---
    df_orc = xls.get(SHEET_ORC, pd.DataFrame())
    if df_orc is None:
        df_orc = pd.DataFrame()
    for c in ["Mês", "Orçado (R$)", "Previsto (R$)", "Realizado (R$)", "Diferença (R$)"]:
        if c not in df_orc.columns:
            df_orc[c] = 0 if c != "Mês" else ""
    df_orc = _ensure_numeric(df_orc, ["Orçado (R$)", "Previsto (R$)", "Realizado (R$)"])
    df_orc["Diferença (R$)"] = (df_orc["Orçado (R$)"] - df_orc["Realizado (R$)"]).round(2)

    # --- CONFIG ---
    df_config = xls.get(SHEET_CONFIG, pd.DataFrame())
    if df_config is None:
        df_config = pd.DataFrame()
    for c in ["Chave", "Valor"]:
        if c not in df_config.columns:
            df_config[c] = ""

    return {
        SHEET_ITENS:  df_itens[ITENS_COLUMNS],
        SHEET_PLANO:  df_plano,
        SHEET_ORC:    df_orc,
        SHEET_CONFIG: df_config,
    }


def salvar_planilha(dfs: Dict[str, pd.DataFrame], caminho: str):
    with pd.ExcelWriter(caminho, engine="openpyxl", mode="w") as writer:
        dfs[SHEET_ITENS][ITENS_COLUMNS].to_excel(writer, sheet_name=SHEET_ITENS, index=False)
        dfs[SHEET_PLANO].to_excel(writer, sheet_name=SHEET_PLANO, index=False)
        dfs[SHEET_ORC].to_excel(writer, sheet_name=SHEET_ORC, index=False)
        if SHEET_CONFIG in dfs and not dfs[SHEET_CONFIG].empty:
            dfs[SHEET_CONFIG].to_excel(writer, sheet_name=SHEET_CONFIG, index=False)


def gerar_tercas(data_inicio: date, data_fim: date) -> pd.DataFrame:
    dias_ate_terca = (1 - data_inicio.weekday()) % 7
    proxima_terca = data_inicio + timedelta(days=dias_ate_terca)
    atual = proxima_terca
    linhas = []
    semana = 1
    while atual <= data_fim:
        linhas.append({
            "Semana #": semana,
            "Terça": atual,
            "Itens sugeridos para cotar/comprar": "",
            "Status da semana": "",
            "Notas": ""
        })
        semana += 1
        atual += timedelta(weeks=1)
    return pd.DataFrame(linhas)


def _precisa_inicializar(caminho: str) -> bool:
    if not os.path.exists(caminho):
        return True
    try:
        df_check = pd.read_excel(caminho, sheet_name=SHEET_ITENS, engine="openpyxl")
        return len(df_check) == 0
    except Exception:
        return True


def _inicializar_se_necessario(caminho: str) -> None:
    if not _precisa_inicializar(caminho):
        return

    df_itens_seed = _criar_df_enxoval()
    df_orc_seed   = pd.DataFrame(DEFAULT_ORCAMENTO)

    dfs_seed: Dict[str, pd.DataFrame] = {
        SHEET_ITENS: df_itens_seed,
        SHEET_ORC:   df_orc_seed,
    }
    salvar_config(dfs_seed, DATA_BASE_DEFAULT, SEMANA_DEFAULT, DPP_DEFAULT)

    df_plan = gerar_tercas(DATA_BASE_DEFAULT, DPP_DEFAULT)
    sugestoes = gerar_sugestoes_faseado(df_itens_seed, len(df_plan))
    df_plan["Itens sugeridos para cotar/comprar"] = sugestoes
    dfs_seed[SHEET_PLANO] = df_plan

    salvar_planilha(dfs_seed, caminho)
    st.cache_data.clear()


# -----------------------------
# Sidebar
# -----------------------------
st.sidebar.title("Configuração de dados")

caminho_xlsx = st.sidebar.text_input("Caminho do arquivo Excel", value=DEFAULT_FILE)
arquivo_up = st.sidebar.file_uploader("Ou carregue um Excel (.xlsx)", type=["xlsx"], help="Se enviar, ele substituirá o arquivo atual enquanto a sessão estiver ativa.")

if arquivo_up is not None:
    with open(caminho_xlsx, "wb") as f:
        f.write(arquivo_up.read())
    st.sidebar.success("Arquivo carregado e salvo com sucesso.")

col_sb_a, col_sb_b = st.sidebar.columns(2)
if col_sb_a.button("Carregar", use_container_width=True):
    st.cache_data.clear()
if col_sb_b.button("Salvar agora", use_container_width=True):
    st.session_state["_force_save_"] = True

# -----------------------------
# Inicialização + carregar dados
# -----------------------------
_inicializar_se_necessario(caminho_xlsx)

base = carregar_planilha(caminho_xlsx)
cfg  = carregar_config(base)

df_itens = base[SHEET_ITENS].copy()
df_plano = base[SHEET_PLANO].copy()
df_orc   = base[SHEET_ORC].copy()

# KPIs
total_itens = len(df_itens)
comprados   = df_itens[df_itens["Status"].isin(["Comprado", "Recebido"])].shape[0]
porc_compra = (comprados / total_itens * 100) if total_itens else 0

valor_estimado  = float(df_itens.get("Custo Estimado (R$)", pd.Series(dtype=float)).sum())
valor_realizado = float(df_itens[df_itens["Status"].isin(["Comprado", "Recebido"])]["Total (R$)"].sum())

st.markdown("## 🍼 Héber e Ju")
col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)
col_kpi1.metric("Itens (total)", f"{total_itens}")
col_kpi2.metric("Itens comprados/recebidos", f"{comprados}", f"{porc_compra:.0f}%")
col_kpi3.metric("Estimado (itens)", f"R$ {valor_estimado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
col_kpi4.metric("Realizado (itens)", f"R$ {valor_realizado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))

abas = st.tabs(["Gestação", "Itens", "Plano semanal (terça)", "Orçamento", "Relatórios", "Configurações"])

# -----------------------------
# TAB 0 — GESTAÇÃO
# -----------------------------
with abas[0]:
    st.subheader("🤰 Dashboard da Gestação")

    hoje_gest = date.today()
    dpp_cfg   = cfg["dpp"]
    db_cfg    = cfg["data_base"]
    sb_cfg    = cfg["semana_base"]

    dias_desde_base = (hoje_gest - db_cfg).days
    semana_atual    = sb_cfg + (dias_desde_base // 7)
    semana_atual    = max(1, min(semana_atual, 42))

    if semana_atual <= 12:
        trimestre_str = "1° Trimestre"
    elif semana_atual <= 27:
        trimestre_str = "2° Trimestre"
    else:
        trimestre_str = "3° Trimestre"

    dias_restantes = (dpp_cfg - hoje_gest).days

    col_g1, col_g2, col_g3, col_g4 = st.columns(4)
    col_g1.metric("Semana atual", f"{semana_atual}ª semana")
    col_g2.metric("Trimestre", trimestre_str)
    col_g3.metric("Dias até o parto", f"{max(0, dias_restantes)} dias")
    col_g4.metric("DPP estimada", dpp_cfg.strftime("%d/%m/%Y"))

    progresso = min(1.0, semana_atual / 40)
    st.progress(progresso, text=f"Progresso da gestação: **{semana_atual}/40 semanas ({progresso:.0%})**")

    st.markdown("---")

    semana_dev = min(max(semana_atual, 22), 40)
    dev = DESENVOLVIMENTO_BEBE.get(semana_dev, {})

    if dev:
        col_dev1, col_dev2 = st.columns([1, 2])
        with col_dev1:
            st.markdown(f"### Bebê na semana {semana_dev}")
            st.markdown(f"**Comprimento:** {dev['comprimento']}")
            st.markdown(f"**Peso:** {dev['peso']}")
            st.markdown(f"**Tamanho de:** {dev['fruta'].capitalize()}")
        with col_dev2:
            st.markdown("**O que está acontecendo esta semana:**")
            st.info(dev["desenvolvimento"])
    else:
        st.info("Informações de desenvolvimento não disponíveis para esta semana.")

    st.markdown("---")

    if not df_itens.empty:
        itens_alta_pendentes = df_itens[
            (df_itens["Prioridade"] == "Alta") &
            (~df_itens["Status"].isin(["Comprado", "Recebido"]))
        ].shape[0]
        custo_restante = float(
            df_itens[~df_itens["Status"].isin(["Comprado", "Recebido"])]["Custo Estimado (R$)"].sum()
        )
        custo_restante_fmt = (
            f"R$ {custo_restante:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )
        col_s1, col_s2 = st.columns(2)
        col_s1.metric("Itens Alta ainda a comprar", f"{itens_alta_pendentes}")
        col_s2.metric("Custo estimado restante", custo_restante_fmt)
    else:
        st.info("Nenhum item cadastrado ainda.")

    if dias_restantes < 0:
        st.success(
            f"Parabéns, Héber e Ju! O bebê já deve ter nascido "
            f"(DPP foi há {abs(dias_restantes)} dias). Bem-vindo ao mundo! 🎉"
        )
    elif dias_restantes <= 14:
        st.warning(
            f"Faltam apenas **{dias_restantes} dias** para a DPP! "
            "Verifique se a mala da maternidade está pronta. 🏥"
        )

# -----------------------------
# TAB 1 — ITENS
# -----------------------------
with abas[1]:
    st.subheader("📦 Itens do enxoval e cuidados do nosso bebê")

    col_f1, col_f2, col_f3, col_f4 = st.columns([2, 2, 2, 2])
    categorias   = sorted(df_itens["Categoria"].dropna().unique().tolist())
    filtro_cat   = col_f1.multiselect("Categoria", options=categorias, default=[])
    filtro_pri   = col_f2.multiselect("Prioridade", options=PRIORIDADES, default=[])
    filtro_sta   = col_f3.multiselect("Status", options=STATUS, default=[])
    termo_busca  = col_f4.text_input("Buscar por item ou nota")

    df_view = df_itens.copy()
    if filtro_cat:
        df_view = df_view[df_view["Categoria"].isin(filtro_cat)]
    if filtro_pri:
        df_view = df_view[df_view["Prioridade"].isin(filtro_pri)]
    if filtro_sta:
        df_view = df_view[df_view["Status"].isin(filtro_sta)]
    if termo_busca:
        mask = (
            df_view["Item"].str.contains(termo_busca, case=False, na=False) |
            df_view["Notas/Observações"].str.contains(termo_busca, case=False, na=False)
        )
        df_view = df_view[mask]

    st.caption("Edite direto na tabela. 'Total (R$)' é recalculado ao salvar.")

    cfg_cols = {
        "Prioridade": st.column_config.SelectboxColumn("Prioridade", options=PRIORIDADES, width=120),
        "Status": st.column_config.SelectboxColumn("Status", options=STATUS, width=140),
        "Qtd": st.column_config.NumberColumn("Qtd", min_value=0, step=1),
        "Preço Unitário (R$)": st.column_config.NumberColumn("Preço Unitário (R$)", min_value=0.0, step=1.0, format="R$ %.2f"),
        "Custo Estimado (R$)": st.column_config.NumberColumn("Custo Estimado (R$)", min_value=0.0, step=1.0, format="R$ %.2f"),
        "Garantia (meses)": st.column_config.NumberColumn("Garantia (meses)", min_value=0, step=1),
        "Total (R$)": st.column_config.NumberColumn("Total (R$)", disabled=True, format="R$ %.2f"),
        "Data da Compra": st.column_config.DateColumn("Data da Compra", format="DD/MM/YYYY"),
        "Loja/Link": st.column_config.LinkColumn("Loja/Link"),
    }

    num_cols = ["Qtd", "Preço Unitário (R$)", "Custo Estimado (R$)", "Garantia (meses)", "Total (R$)"]
    for c in num_cols:
        if c in df_view.columns:
            df_view[c] = pd.to_numeric(df_view[c], errors="coerce").fillna(0.0)

    if "ID" in df_view.columns:
        df_view["ID"] = pd.to_numeric(df_view["ID"], errors="coerce").fillna(0).astype(int)
        mask_zero = df_view["ID"] <= 0
        if mask_zero.any():
            df_view.loc[mask_zero, "ID"] = range(df_view["ID"].max()+1, df_view["ID"].max()+1+mask_zero.sum())

    cat_cols = ["Prioridade", "Status", "Categoria", "Unidade", "Marca/Modelo", "Loja/Link", "Notas/Observações", "Item"]
    for c in cat_cols:
        if c in df_view.columns:
            df_view[c] = df_view[c].astype(str).replace({"nan": ""}).fillna("")

    if "Data da Compra" in df_view.columns:
        def _to_date(x):
            if pd.isna(x) or x in ("", "nan", "NaT"):
                return None
            try:
                return pd.to_datetime(x, dayfirst=True).date()
            except Exception:
                return None
        df_view["Data da Compra"] = df_view["Data da Compra"].apply(_to_date)

    if "Loja/Link" in df_view.columns:
        def _fix_url(s):
            s = (s or "").strip()
            if not s:
                return ""
            if s.startswith(("http://", "https://")):
                return s
            return "http://" + s
        df_view["Loja/Link"] = df_view["Loja/Link"].apply(_fix_url)

    edited_df = st.data_editor(
        df_view,
        column_config=cfg_cols,
        width="stretch",
        height=450,
        num_rows="dynamic",
        key="editor_itens"
    )

    col_a, col_b, col_c = st.columns([1, 1, 1])

    with col_a:
        if st.button("💾 Salvar alterações", type="primary") or st.session_state.get("_force_save_"):
            df_merged = df_itens.set_index("ID").copy()
            for _, row in edited_df.iterrows():
                rid = int(row["ID"]) if not pd.isna(row["ID"]) else None
                if rid in df_merged.index:
                    df_merged.loc[rid, edited_df.columns] = row
                else:
                    if str(row.get("Item", "")).strip():
                        new_id = (df_merged.index.max() or 0) + 1
                        row_cp = row.to_dict()
                        row_cp["ID"] = new_id
                        df_merged = pd.concat([df_merged, pd.DataFrame([row_cp]).set_index("ID")])

            df_merged["Qtd"] = pd.to_numeric(df_merged["Qtd"], errors="coerce").fillna(0)
            df_merged["Preço Unitário (R$)"] = pd.to_numeric(df_merged["Preço Unitário (R$)"], errors="coerce").fillna(0.0)
            df_merged["Total (R$)"] = (df_merged["Qtd"] * df_merged["Preço Unitário (R$)"]).round(2)

            df_final = df_merged.reset_index()
            df_final = _normalize_id_column(df_final)
            base[SHEET_ITENS] = df_final[ITENS_COLUMNS]
            try:
                salvar_planilha(base, caminho_xlsx)
                st.success("Itens salvos com sucesso!")
                st.session_state["_force_save_"] = False
                st.cache_data.clear()
            except Exception as e:
                st.error(f"Erro ao salvar: {e}")

    with col_b:
        with st.popover("➕ Adicionar novo item"):
            st.markdown("**Preencha os campos do novo item**")
            c1, c2 = st.columns(2)
            cat  = c1.text_input("Categoria")
            item = c2.text_input("Item")
            c3, c4, c5 = st.columns(3)
            pri  = c3.selectbox("Prioridade", PRIORIDADES, index=1)
            qtd  = c4.number_input("Qtd", min_value=0, step=1, value=1)
            und  = c5.text_input("Unidade", value="un")
            c6, c7 = st.columns(2)
            marca = c6.text_input("Marca/Modelo")
            loja  = c7.text_input("Loja/Link")
            c8, c9 = st.columns(2)
            preco     = c8.number_input("Preço Unitário (R$)", min_value=0.0, step=1.0, format="%.2f")
            custo_est = c9.number_input("Custo Estimado (R$)", min_value=0.0, step=1.0, format="%.2f")
            c10, c11 = st.columns(2)
            status   = c10.selectbox("Status", STATUS, index=0)
            garantia = c11.number_input("Garantia (meses)", min_value=0, step=1, value=0)
            notas    = st.text_area("Notas/Observações")
            if st.button("Adicionar"):
                new_id = (df_itens["ID"].max() if len(df_itens) else 0) + 1
                novo = {
                    "ID": new_id, "Categoria": cat, "Item": item,
                    "Prioridade": pri, "Qtd": qtd, "Unidade": und,
                    "Marca/Modelo": marca, "Loja/Link": loja,
                    "Preço Unitário (R$)": preco, "Custo Estimado (R$)": custo_est,
                    "Total (R$)": round(qtd * preco, 2), "Status": status,
                    "Data da Compra": "", "Garantia (meses)": garantia, "Notas/Observações": notas
                }
                df_new = pd.concat([df_itens, pd.DataFrame([novo])], ignore_index=True)
                df_new = _normalize_id_column(df_new)
                base[SHEET_ITENS] = df_new[ITENS_COLUMNS]
                try:
                    salvar_planilha(base, caminho_xlsx)
                    st.success("Item adicionado!")
                    st.cache_data.clear()
                except Exception as e:
                    st.error(f"Erro ao salvar: {e}")

    with col_c:
        with st.popover("🗑️ Excluir itens"):
            ids = st.multiselect("Selecione IDs para excluir", options=df_itens["ID"].tolist())
            if st.button("Excluir selecionados", type="secondary"):
                df_new = df_itens[~df_itens["ID"].isin(ids)].copy()
                df_new = _normalize_id_column(df_new)
                base[SHEET_ITENS] = df_new[ITENS_COLUMNS]
                try:
                    salvar_planilha(base, caminho_xlsx)
                    st.success("Itens excluídos.")
                    st.cache_data.clear()
                except Exception as e:
                    st.error(f"Erro ao salvar: {e}")

# -----------------------------
# TAB 2 — PLANO SEMANAL
# -----------------------------
with abas[2]:
    st.subheader("🗓️ Plano semanal — toda terça")
    st.caption("Acompanhe as terças e acrescente notas/status da semana.")

    if df_plano.empty:
        st.info("Ainda não há plano semanal. Gere nas Configurações.")
    else:
        cfg_plano = {
            "Status da semana": st.column_config.SelectboxColumn("Status da semana", options=["", "A fazer", "Concluída"], width=120),
            "Terça": st.column_config.DateColumn("Terça", format="DD/MM/YYYY"),
        }
        edited_plano = st.data_editor(df_plano, width="stretch", height=450, key="editor_plano", column_config=cfg_plano)
        if st.button("💾 Salvar plano", type="primary"):
            base[SHEET_PLANO] = edited_plano
            try:
                salvar_planilha(base, caminho_xlsx)
                st.success("Plano semanal salvo!")
                st.cache_data.clear()
            except Exception as e:
                st.error(f"Erro ao salvar: {e}")

# -----------------------------
# TAB 3 — ORÇAMENTO
# -----------------------------
with abas[3]:
    st.subheader("💰 Orçamento mensal")
    st.caption("Preencha Orçado/Previsto/Realizado. A Diferença é calculada automaticamente.")

    df_orc = _ensure_numeric(df_orc, ["Orçado (R$)", "Previsto (R$)", "Realizado (R$)"])
    df_orc["Diferença (R$)"] = (df_orc["Orçado (R$)"] - df_orc["Realizado (R$)"]).round(2)

    cfg_orc = {
        "Mês": st.column_config.TextColumn("Mês", help="Formato AAAA-MM"),
        "Orçado (R$)": st.column_config.NumberColumn("Orçado (R$)", format="R$ %.2f"),
        "Previsto (R$)": st.column_config.NumberColumn("Previsto (R$)", format="R$ %.2f"),
        "Realizado (R$)": st.column_config.NumberColumn("Realizado (R$)", format="R$ %.2f"),
        "Diferença (R$)": st.column_config.NumberColumn("Diferença (R$)", disabled=True, format="R$ %.2f"),
    }

    edited_orc = st.data_editor(df_orc, width="stretch", height=350, key="editor_orc", column_config=cfg_orc)

    col_o1, col_o2 = st.columns([1, 1])
    if col_o1.button("💾 Salvar orçamento", type="primary"):
        edited_orc = _ensure_numeric(edited_orc, ["Orçado (R$)", "Previsto (R$)", "Realizado (R$)"])
        edited_orc["Diferença (R$)"] = (edited_orc["Orçado (R$)"] - edited_orc["Realizado (R$)"]).round(2)
        base[SHEET_ORC] = edited_orc
        try:
            salvar_planilha(base, caminho_xlsx)
            st.success("Orçamento salvo!")
            st.cache_data.clear()
        except Exception as e:
            st.error(f"Erro ao salvar: {e}")

    st.markdown("---")
    st.markdown("#### Visão mensal consolidada")
    if not edited_orc.empty:
        total_orcado         = float(edited_orc["Orçado (R$)"].sum())
        total_previsto       = float(edited_orc["Previsto (R$)"].sum())
        total_realizado_orc  = float(edited_orc["Realizado (R$)"].sum())
        col_oa, col_ob, col_oc = st.columns(3)
        col_oa.metric("Total orçado", f"R$ {total_orcado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        col_ob.metric("Total previsto", f"R$ {total_previsto:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        col_oc.metric("Total realizado (orçamento)", f"R$ {total_realizado_orc:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))

# -----------------------------
# TAB 4 — RELATÓRIOS
# -----------------------------
with abas[4]:
    st.subheader("📊 Relatórios e acompanhamento")

    if not df_itens.empty:
        st.markdown("**Gastos por categoria (realizado vs estimado)**")
        df_cat = df_itens.copy()
        df_cat["Realizado (R$)"] = df_cat.apply(
            lambda r: r["Total (R$)"] if r["Status"] in ["Comprado", "Recebido"] else 0.0, axis=1
        )
        grp = df_cat.groupby("Categoria").agg({
            "Custo Estimado (R$)": "sum",
            "Realizado (R$)": "sum"
        }).sort_values("Realizado (R$)", ascending=False)
        st.bar_chart(grp, width="stretch")

        st.markdown("**Progresso por status**")
        prog = df_itens["Status"].value_counts().reindex(STATUS, fill_value=0)
        st.bar_chart(prog, width="stretch")

        st.markdown("**Top 10 itens por custo estimado**")
        df_top = df_cat.sort_values("Custo Estimado (R$)", ascending=False).head(10)[["Categoria", "Item", "Custo Estimado (R$)", "Prioridade", "Status"]]
        st.dataframe(df_top, width="stretch", use_container_width=True)
    else:
        st.info("Sem itens para gerar relatórios.")

# -----------------------------
# TAB 5 — CONFIGURAÇÕES
# -----------------------------
with abas[5]:
    st.subheader("⚙️ Configurações e geração do plano de terças")

    col_c1, col_c2, col_c3 = st.columns(3)
    data_base         = col_c1.date_input("Data base (dia de hoje na gestação)", value=cfg["data_base"])
    idade_gestacional = col_c2.number_input("Idade gestacional (semanas)", min_value=0, max_value=42, value=cfg["semana_base"])
    dpp_auto          = data_base + timedelta(weeks=(40 - idade_gestacional)) if idade_gestacional <= 40 else data_base
    dpp               = col_c3.date_input("DPP (estimada)", value=cfg["dpp"])

    st.caption("Gerar ou recalcular as terças até a DPP. Você pode editar manualmente depois.")
    if st.button("📅 Gerar/atualizar plano de terças", type="primary"):
        df_new_plan = gerar_tercas(data_base, dpp)

        if not df_itens.empty:
            prioridade_ordem = {"Alta": 0, "Média": 1, "Baixa": 2}
            df_sorted = df_itens.sort_values(
                by=["Prioridade", "Categoria", "Item"],
                key=lambda s: s.map(prioridade_ordem).fillna(3) if s.name == "Prioridade" else s
            )
            sugestoes = gerar_sugestoes_faseado(df_sorted, len(df_new_plan))
            df_new_plan["Itens sugeridos para cotar/comprar"] = sugestoes

        salvar_config(base, data_base, idade_gestacional, dpp)
        base[SHEET_PLANO] = df_new_plan
        try:
            salvar_planilha(base, caminho_xlsx)
            st.success("Plano de terças atualizado! A aba Gestação também foi atualizada.")
            st.cache_data.clear()
        except Exception as e:
            st.error(f"Erro ao salvar: {e}")

    st.markdown("---")
    st.markdown("### Exportar uma cópia")
    buffer = io.BytesIO()
    try:
        salvar_planilha(base, caminho_xlsx)
        dfs_copy = carregar_planilha(caminho_xlsx)
        with pd.ExcelWriter(buffer, engine="openpyxl", mode="w") as writer:
            dfs_copy[SHEET_ITENS][ITENS_COLUMNS].to_excel(writer, sheet_name=SHEET_ITENS, index=False)
            dfs_copy[SHEET_PLANO].to_excel(writer, sheet_name=SHEET_PLANO, index=False)
            dfs_copy[SHEET_ORC].to_excel(writer, sheet_name=SHEET_ORC, index=False)
            if SHEET_CONFIG in dfs_copy and not dfs_copy[SHEET_CONFIG].empty:
                dfs_copy[SHEET_CONFIG].to_excel(writer, sheet_name=SHEET_CONFIG, index=False)
        st.download_button(
            label="Baixar cópia do Excel",
            data=buffer.getvalue(),
            file_name="Planejamento_Bebe_Heber_copia.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Não foi possível preparar a cópia: {e}")
