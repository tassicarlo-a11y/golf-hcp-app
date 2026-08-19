import os
import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="Golf Handicap Tracker", page_icon="⛳", layout="wide"
)




# --- FUNZIONE CALCOLO SD DA STABLEFORD ---
def calcola_sd_da_stableford(
    stbl_points, playing_hcp, par, cr, sr, buche=18, pcc=0.0
):
  if buche == 9:
    par_eff = par * 2 if par < 50 else par
    cr_eff = cr * 2 if cr < 50 else cr
    hcp_eff = playing_hcp * 2 if playing_hcp < 15 else playing_hcp
    stbl_eff = stbl_points + 17 if stbl_points <= 25 else stbl_points
  else:
    par_eff, cr_eff, hcp_eff, stbl_eff = par, cr, playing_hcp, stbl_points


  ags = par_eff + hcp_eff + 36 - stbl_eff
  sd = (113 / sr) * (ags - cr_eff - pcc)
  return round(sd, 1)




# --- CARICAMENTO DATI ---
@st.cache_data(ttl=1)
def load_data():
  if os.path.exists("Handicap_2026.xlsx"):
    df_raw = pd.read_excel("Handicap_2026.xlsx", sheet_name="Foglio2")
    return df_raw
  return pd.DataFrame()




df = load_data()




# --- CALCOLO HCP CORRENTE E STORICO ---
def calcola_hcp_corrente(df_in):
  validi = df_in[
      (df_in["Valida"].isin(["V", "S"])) & (df_in["SD"].notna())
  ].copy()
  ultimi_20 = validi.head(20)
  if len(ultimi_20) == 0:
    return 0.0, ultimi_20, []
  n_best = min(8, len(ultimi_20))
  migliori_8 = ultimi_20.nsmallest(n_best, "SD")
  return round(migliori_8["SD"].mean(), 1), ultimi_20, migliori_8.index.tolist()




def genera_trend_storico(df_in):
  validi = df_in[
      (df_in["Valida"].isin(["V", "S"])) & (df_in["SD"].notna())
  ].copy()
  if validi.empty:
    return pd.DataFrame()


  validi["Data_DT"] = pd.to_datetime(validi["Data"])
  # Ordiniamo cronologicamente (dal più vecchio al più recente)
  validi_cron = validi.sort_values("Data_DT", ascending=True).reset_index(
      drop=True
  )


  hcp_progressivo = []
  for i in range(len(validi_cron)):
    # Finestra degli ultimi max 20 risultati fino alla gara i-esima
    window = validi_cron.iloc[max(0, i - 19) : i + 1]
    n_best = min(8, len(window))
    avg_hcp = round(window["SD"].nsmallest(n_best).mean(), 1)
    hcp_progressivo.append(avg_hcp)


  validi_cron["HCP_Storico"] = hcp_progressivo
  return validi_cron




# --- HEADER PRINCIPALE ---
st.title("⛳ Golf Handicap Tracker & Simulator")


if not df.empty:
  hcp_attuale, ultimi_20, id_migliori = calcola_hcp_corrente(df)
  df_trend = genera_trend_storico(df)


  # Indicatori principali in alto
  col_m1, col_m2, col_m3 = st.columns(3)
  with col_m1:
    st.metric(
        "Handicap Index Attuale", value=hcp_attuale, help="Media dei migliori 8 ultimi 20"
    )
  with col_m2:
    st.metric("Gare Valide Registrate", value=len(df_trend))
  with col_m3:
    if not df_trend.empty:
      miglior_hcp = df_trend["HCP_Storico"].min()
      st.metric("Miglior Handicap Storico", value=miglior_hcp)


  # --- STRUTTURA A SCHEDE ---
  tab_dash, tab_sim, tab_reg = st.tabs([
      "📊 Dashboard & Trend HCP",
      "🔮 Simulazione Gara",
      "📋 Registro Gare Ufficiale",
  ])


  # ==========================================
  # TAB 1: DASHBOARD & TREND
  # ==========================================
  with tab_dash:
    st.subheader("📈 Trend Storico Handicap Index")


    if not df_trend.empty:
      # Grafico interattivo Plotly
      fig = px.line(
          df_trend,
          x="Data_DT",
          y="HCP_Storico",
          markers=True,
          hover_data={"Gara": True, "SD": True, "Data_DT": False},
          labels={"Data_DT": "Data", "HCP_Storico": "Handicap Index"},
          title="Evoluzione Handicap nel Tempo",
      )
      fig.update_traces(
          line_color="#2E7D32",
          line_width=3,
          marker=dict(size=8, color="#1B5E20"),
      )
      fig.update_layout(
          hovermode="x unified", yaxis=dict(autorange="reversed")
      )  # Inverte Y (più basso è l'HCP, più sta in alto)
      st.plotly_chart(fig, use_container_width=True)


    st.subheader("🟢 Ultimi 20 Risultati Validi (Evidenziati i Migliori 8)")


    def highlight_b8(row):
      return (
          ["background-color: #d4edda; font-weight: bold;"] * len(row)
          if row.name in id_migliori
          else [""] * len(row)
      )


    st.dataframe(
        ultimi_20[["Data", "Gara", "Valida", "Buche", "SD"]].style.apply(
            highlight_b8, axis=1
        ),
        use_container_width=True,
    )


  # ==========================================
  # TAB 2: SIMULAZIONE GARA
  # ==========================================
  with tab_sim:
    st.subheader("🔮 Simula Impatto Prossima Gara")


    modalita = st.radio(
        "Modalità Inserimento:",
        ["Calcola da Punti Stableford", "Score Differential (SD) Manuale"],
        horizontal=True,
    )


    with st.form("simulazione_form"):
      nome_gara = st.text_input("Nome Gara Simulata", "Gara Prossima")
      col1, col2 = st.columns(2)


      with col1:
        buche = st.selectbox("Buche", [18, 9])
        par = st.number_input("Par Campo", value=72 if buche == 18 else 32)
        cr = st.number_input(
            "Course Rating (CR)", value=71.2 if buche == 18 else 31.7
        )


      with col2:
        sr = st.number_input("Slope Rating (SR)", value=124 if buche == 18 else 114)
        pcc = st.number_input("PCC (-1.0 a +3.0)", value=0.0, step=0.5)


      if modalita == "Calcola da Punti Stableford":
        col3, col4 = st.columns(2)
        with col3:
          stbl = st.number_input(
              "Punti Stableford Realizzati", value=36 if buche == 18 else 18
          )
        with col4:
          playing_hcp = st.number_input(
              "Playing HCP (Gara)", value=10.0 if buche == 18 else 5.0
          )
        sd_simulato = calcola_sd_da_stableford(
            stbl, playing_hcp, par, cr, sr, buche, pcc
        )
        st.info(f"💡 **Score Differential (SD) Calcolato:** `{sd_simulato}`")
      else:
        sd_simulato = st.number_input("SD Manuale", value=10.0, step=0.1)


      submit_sim = st.form_submit_button("🧪 Esegui Simulazione")


    if submit_sim:
      riga_sim = pd.DataFrame([{
          "Data": pd.Timestamp.now().strftime("%Y-%m-%d"),
          "Gara": f"[SIMULATA] {nome_gara}",
          "Valida": "V",
          "Buche": buche,
          "Par": par,
          "CR": cr,
          "SR": sr,
          "SD": sd_simulato,
      }])


      df_sim = pd.concat([riga_sim, df], ignore_index=True)
      hcp_sim, _, _ = calcola_hcp_corrente(df_sim)
      diff = round(hcp_sim - hcp_attuale, 1)


      if diff < 0:
        st.success(
            f"🎉 **Nuovo HCP Simulato: {hcp_sim}** (Miglioramento di"
            f" {abs(diff)} colpi!)"
        )
      elif diff > 0:
        st.warning(
            f"⚠️ **Nuovo HCP Simulato: {hcp_sim}** (Peggiore di +{diff} colpi)"
        )
      else:
        st.info(f"➡️ **Nuovo HCP Simulato: {hcp_sim}** (Nessuna variazione)")


      if len(ultimi_20) == 20:
        scartato = ultimi_20.iloc[-1]
        st.caption(
            f"📌 Uscirà dal calcolo dei 20 risultati la gara: **{scartato['Gara']}**"
            f" del {scartato['Data']} (SD: {scartato['SD']})"
        )


  # ==========================================
  # TAB 3: REGISTRO & INSERIMENTO DEFINITIVO
  # ==========================================
  with tab_reg:
    st.subheader("💾 Inserisci Risultato Definitivo in Excel")


    with st.form("inserimento_definitivo"):
      g_data = st.date_input("Data Gara")
      g_nome = st.text_input("Nome Gara Ufficiale")
      col_a, col_b = st.columns(2)
      with col_a:
        g_buche = st.selectbox("Buche ", [18, 9])
        g_sd = st.number_input("SD Definitivo", value=10.0, step=0.1)
      with col_b:
        g_valida = st.selectbox("Valida", ["V", "S", "N"])


      submit_save = st.form_submit_button("💾 Salva in Excel")


      if submit_save and g_nome:
        nuova_riga = pd.DataFrame([{
            "Data": pd.to_datetime(g_data).strftime("%Y-%m-%d"),
            "Gara": g_nome,
            "Valida": g_valida,
            "Buche": g_buche,
            "SD": g_sd if g_valida in ["V", "S"] else None,
        }])
        df_finale = pd.concat([nuova_riga, df], ignore_index=True)
        df_finale.to_excel(
            "Handicap_2026.xlsx", sheet_name="Foglio2", index=False
        )
        st.success("Risultato salvato permanentemente nel file Excel!")
        st.rerun()


else:
  st.error(
      "File 'Handicap_2026.xlsx' non trovato. Assicurati che sia presente nella"
      " cartella principale."
  )