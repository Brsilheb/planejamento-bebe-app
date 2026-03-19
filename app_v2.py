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

DEFAULT_FILE = "Planejamento_Bebe_Heber.xlsx"
SHEET_ITENS = "01_Itens"
SHEET_PLANO = "02_Plano_Semanal"
SHEET_ORC = "03_Orcamento"

PRIORIDADES = ["Alta", "Média", "Baixa"]
STATUS = ["A comprar", "Cotar", "Comprado", "Recebido"]

ITENS_COLUMNS = [
    "ID", "Categoria", "Item", "Prioridade", "Qtd", "Unidade",
    "Marca/Modelo", "Loja/Link", "Preço Unitário (R$)", "Custo Estimado (R$)",
    "Total (R$)", "Status", "Data da Compra", "Garantia (meses)", "Notas/Observações"
]

# -----------------------------
# Helpers
# -----------------------------

def _ensure_numeric(df: pd.DataFrame, cols):
    for c in cols:
        df[c] = pd.to_numeric(df.get(c, 0), errors="coerce").fillna(0.0)
    return df


def _normalize_id_column(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure 'ID' exists, is unique, integer, sequential starting at 1."""
    if "ID" not in df.columns:
        df.insert(0, "ID", range(1, len(df) + 1))
    else:
        # Coerce to int, replace NaN with 0
        df["ID"] = pd.to_numeric(df["ID"], errors="coerce").fillna(0).astype(int)
        # If any duplicated or zeros -> regenerate sequential IDs
        if (df["ID"] == 0).any() or df["ID"].duplicated().any():
            df = df.drop(columns=["ID"]).reset_index(drop=True)
            df.insert(0, "ID", range(1, len(df) + 1))
    return df


@st.cache_data(show_spinner=False)
def carregar_planilha(caminho: str) -> Dict[str, pd.DataFrame]:
    if not os.path.exists(caminho):
        df_itens = pd.DataFrame(columns=ITENS_COLUMNS)
        df_plano = pd.DataFrame(columns=["Semana #", "Terça", "Itens sugeridos para cotar/comprar", "Status da semana", "Notas"])
        df_orc = pd.DataFrame(columns=["Mês", "Orçado (R$)", "Previsto (R$)", "Realizado (R$)", "Diferença (R$)"])
        return {SHEET_ITENS: df_itens, SHEET_PLANO: df_plano, SHEET_ORC: df_orc}

    xls = pd.read_excel(caminho, sheet_name=None, engine="openpyxl")

    # --- ITENS ---
    df_itens = xls.get(SHEET_ITENS, pd.DataFrame())
    if df_itens is None:
        df_itens = pd.DataFrame()
    # Garante colunas
    for c in ITENS_COLUMNS:
        if c not in df_itens.columns:
            df_itens[c] = '' if c not in ["Qtd", "Preço Unitário (R$)", "Custo Estimado (R$)", "Garantia (meses)", "Total (R$)"] else 0
    # Normaliza ID
    df_itens = _normalize_id_column(df_itens)
    # Tipos
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

    return {SHEET_ITENS: df_itens[ITENS_COLUMNS], SHEET_PLANO: df_plano, SHEET_ORC: df_orc}


def salvar_planilha(dfs: Dict[str, pd.DataFrame], caminho: str):
    with pd.ExcelWriter(caminho, engine="openpyxl", mode="w") as writer:
        dfs[SHEET_ITENS][ITENS_COLUMNS].to_excel(writer, sheet_name=SHEET_ITENS, index=False)
        dfs[SHEET_PLANO].to_excel(writer, sheet_name=SHEET_PLANO, index=False)
        dfs[SHEET_ORC].to_excel(writer, sheet_name=SHEET_ORC, index=False)


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
# Carregar dados
# -----------------------------
base = carregar_planilha(caminho_xlsx)

df_itens = base[SHEET_ITENS].copy()
df_plano = base[SHEET_PLANO].copy()
df_orc = base[SHEET_ORC].copy()

# KPIs
total_itens = len(df_itens)
comprados = df_itens[df_itens["Status"].isin(["Comprado", "Recebido"])].shape[0]
porc_compra = (comprados / total_itens * 100) if total_itens else 0

valor_estimado = float(df_itens.get("Custo Estimado (R$)", pd.Series(dtype=float)).sum())
valor_realizado = float(df_itens[df_itens["Status"].isin(["Comprado", "Recebido"])]["Total (R$)"].sum())

st.markdown("## 🍼 Planejamento")
col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)
col_kpi1.metric("Itens (total)", f"{total_itens}")
col_kpi2.metric("Itens comprados/recebidos", f"{comprados}", f"{porc_compra:.0f}%")
col_kpi3.metric("Estimado (itens)", f"R$ {valor_estimado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
col_kpi4.metric("Realizado (itens)", f"R$ {valor_realizado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))

abas = st.tabs(["Itens", "Plano semanal (terça)", "Orçamento", "Relatórios", "Configurações"])

# -----------------------------
# ITENS
# -----------------------------
with abas[0]:
    st.subheader("📦 Itens do enxoval e cuidados")

    col_f1, col_f2, col_f3, col_f4 = st.columns([2, 2, 2, 2])
    categorias = sorted(df_itens["Categoria"].dropna().unique().tolist())
    filtro_cat = col_f1.multiselect("Categoria", options=categorias, default=[])
    filtro_pri = col_f2.multiselect("Prioridade", options=PRIORIDADES, default=[])
    filtro_sta = col_f3.multiselect("Status", options=STATUS, default=[])
    termo_busca = col_f4.text_input("Buscar por item ou nota")

    df_view = df_itens.copy()
    if filtro_cat:
        df_view = df_view[df_view["Categoria"].isin(filtro_cat)]
    if filtro_pri:
        df_view = df_view[df_view["Prioridade"].isin(filtro_pri)]
    if filtro_sta:
        df_view = df_view[df_view["Status"].isin(filtro_sta)]
    if termo_busca:
        mask = df_view["Item"].str.contains(termo_busca, case=False, na=False) | df_view["Notas/Observações"].str.contains(termo_busca, case=False, na=False)
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

    edited_df = st.data_editor(
        df_view,
        column_config=cfg_cols,
        width="stretch",
        height=450,
        num_rows="dynamic",
        key="editor_itens"
    )

    col_a, col_b, col_c = st.columns([1,1,1])

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

            # Recalcula total e normaliza IDs
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
            cat = c1.text_input("Categoria")
            item = c2.text_input("Item")
            c3, c4, c5 = st.columns(3)
            pri = c3.selectbox("Prioridade", PRIORIDADES, index=1)
            qtd = c4.number_input("Qtd", min_value=0, step=1, value=1)
            und = c5.text_input("Unidade", value="un")
            c6, c7 = st.columns(2)
            marca = c6.text_input("Marca/Modelo")
            loja = c7.text_input("Loja/Link")
            c8, c9 = st.columns(2)
            preco = c8.number_input("Preço Unitário (R$)", min_value=0.0, step=1.0, format="%.2f")
            custo_est = c9.number_input("Custo Estimado (R$)", min_value=0.0, step=1.0, format="%.2f")
            c10, c11 = st.columns(2)
            status = c10.selectbox("Status", STATUS, index=0)
            garantia = c11.number_input("Garantia (meses)", min_value=0, step=1, value=0)
            notas = st.text_area("Notas/Observações")
            if st.button("Adicionar"):
                new_id = (df_itens["ID"].max() if len(df_itens) else 0) + 1
                novo = {
                    "ID": new_id,
                    "Categoria": cat,
                    "Item": item,
                    "Prioridade": pri,
                    "Qtd": qtd,
                    "Unidade": und,
                    "Marca/Modelo": marca,
                    "Loja/Link": loja,
                    "Preço Unitário (R$)": preco,
                    "Custo Estimado (R$)": custo_est,
                    "Total (R$)": round(qtd * preco, 2),
                    "Status": status,
                    "Data da Compra": "",
                    "Garantia (meses)": garantia,
                    "Notas/Observações": notas
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
# PLANO SEMANAL
# -----------------------------
with abas[1]:
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
# ORÇAMENTO
# -----------------------------
with abas[2]:
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

    col_o1, col_o2 = st.columns([1,1])
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
        total_orcado = float(edited_orc["Orçado (R$)"].sum())
        total_previsto = float(edited_orc["Previsto (R$)"].sum())
        total_realizado_orc = float(edited_orc["Realizado (R$)"].sum())
        col_oa, col_ob, col_oc = st.columns(3)
        col_oa.metric("Total orçado", f"R$ {total_orcado:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        col_ob.metric("Total previsto", f"R$ {total_previsto:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        col_oc.metric("Total realizado (orçamento)", f"R$ {total_realizado_orc:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))

# -----------------------------
# RELATÓRIOS
# -----------------------------
with abas[3]:
    st.subheader("📊 Relatórios e acompanhamento")

    if not df_itens.empty:
        st.markdown("**Gastos por categoria (realizado vs estimado)**")
        df_cat = df_itens.copy()
        df_cat["Realizado (R$)"] = df_cat.apply(lambda r: r["Total (R$)"] if r["Status"] in ["Comprado", "Recebido"] else 0.0, axis=1)
        grp = df_cat.groupby("Categoria").agg({
            "Custo Estimado (R$)": "sum",
            "Realizado (R$)": "sum"
        }).sort_values("Realizado (R$)", ascending=False)
        st.bar_chart(grp, width="stretch")

        st.markdown("**Progresso por status**")
        prog = df_itens["Status"].value_counts().reindex(STATUS, fill_value=0)
        st.bar_chart(prog, width="stretch")

        st.markdown("**Top 10 itens por gasto realizado**")
        df_top = df_cat.sort_values("Realizado (R$)", ascending=False).head(10)[["Categoria", "Item", "Realizado (R$)"]]
        st.dataframe(df_top, width="stretch", use_container_width=False)
    else:
        st.info("Sem itens para gerar relatórios.")

# -----------------------------
# CONFIGURAÇÕES
# -----------------------------
with abas[4]:
    st.subheader("⚙️ Configurações e geração do plano de terças")

    hoje = date.today()
    col_c1, col_c2, col_c3 = st.columns(3)
    data_base = col_c1.date_input("Data base", value=hoje)
    idade_gestacional = col_c2.number_input("Idade gestacional (semanas)", min_value=0, max_value=42, value=13)
    dpp_auto = data_base + timedelta(weeks=(40 - idade_gestacional)) if idade_gestacional <= 40 else data_base
    dpp = col_c3.date_input("DPP (estimada)", value=dpp_auto)

    st.caption("Gerar ou recalcular as terças até a DPP. Você pode editar manualmente depois.")
    if st.button("📅 Gerar/atualizar plano de terças", type="primary"):
        df_new_plan = gerar_tercas(data_base, dpp)

        if not df_itens.empty:
            prioridade_ordem = {"Alta": 0, "Média": 1, "Baixa": 2}
            df_sorted = df_itens.sort_values(by=["Prioridade", "Categoria", "Item"], key=lambda s: s.map(prioridade_ordem).fillna(3) if s.name=="Prioridade" else s)
            n_semanas = max(len(df_new_plan), 1)
            N = max(3, math.ceil(len(df_sorted) / n_semanas))
            idx = 0
            sugestoes = []
            for _ in range(len(df_new_plan)):
                bloc = []
                for _ in range(N):
                    if idx < len(df_sorted):
                        r = df_sorted.iloc[idx]
                        bloc.append(f"{r['Categoria']} - {r['Item']} (Prioridade: {r['Prioridade']})")
                        idx += 1
                sugestoes.append("\n".join(bloc))
            df_new_plan["Itens sugeridos para cotar/comprar"] = sugestoes

        base[SHEET_PLANO] = df_new_plan
        try:
            salvar_planilha(base, caminho_xlsx)
            st.success("Plano de terças atualizado!")
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
        st.download_button(
            label="Baixar cópia do Excel",
            data=buffer.getvalue(),
            file_name="Planejamento_Bebe_Heber_copia.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Não foi possível preparar a cópia: {e}")
