import json
import os
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Golf Handicap Tracker", page_icon="⛳", layout="wide"
)


# --- CARICAMENTO DATI INTELLIGENTE ---
@st.cache_data(ttl=1)
def load_data():
  filename = None
  for fname in ["Handicap_2026.xlsx", "Handicap 2026.xlsx"]:
    if os.path.exists(fname):
      filename = fname
      break

  if not filename:
    return pd.DataFrame()

  # Prova prima con la riga 1 come intestazione (header=0)
  df = pd.read_excel(filename, sheet_name="Foglio2")
  if "Data" in df.columns and "SD" in df.columns:
    df["Data"] = pd.to_datetime(df["Data"])
    return df

  # Se non trova le colonne, prova con la riga 2 (header=1)
  df_h1 = pd.read_excel(filename, sheet_name="Foglio2", header=1)
  if "Data" in df_h1.columns and "SD" in df_h1.columns:
    df_h1["Data"] = pd.to_datetime(df_h1["Data"])
    return df_h1

  return df


def save_data(df_to_save):
  filename = "Handicap_2026.xlsx"
  if os.path.exists("Handicap 2026.xlsx") and not os.path.exists(
      "Handicap_2026.xlsx"
  ):
    filename = "Handicap 2026.xlsx"

  df_clean = df_to_save.copy()
  if "Data" in df_clean.columns:
    df_clean["Data"] = pd.to_datetime(df_clean["Data"]).dt.strftime("%Y-%m-%d")

  df_clean.to_excel(filename, sheet_name="Foglio2", index=False)
  st.cache_data.clear()


# --- CARICAMENTO E SALVATAGGIO ANAGRAFICA CAMPI ---
def load_campi():
  json_path = "campi.json"
  if os.path.exists(json_path):
    try:
      with open(json_path, "r", encoding="utf-8") as f:
        return json.load(f)
    except Exception:
      pass

  campi = []
  df_ex = load_data()
  if not df_ex.empty and "Esecutore" in df_ex.columns:
    subset = (
        df_ex[["Esecutore", "Buche", "Par", "CR", "SR"]]
        .dropna(subset=["Esecutore"])
        .drop_duplicates()
    )
    for _, row in subset.iterrows():
      campi.append({
          "Nome": str(row["Esecutore"]),
          "Buche": int(row["Buche"]),
          "Par": float(row["Par"]),
          "CR": float(row["CR"]),
          "SR": float(row["SR"]),
      })

  if not campi:
    campi = [
        {
            "Nome": "MONTEVEGLIO ASD",
            "Buche": 9,
            "Par": 32.0,
            "CR": 31.7,
            "SR": 114.0,
        },
        {
            "Nome": "MONTEVEGLIO ASD",
            "Buche": 18,
            "Par": 64.0,
            "CR": 63.4,
            "SR": 114.0,
        },
        {"Nome": "MONTICELLO", "Buche": 18, "Par": 72.0, "CR": 72.1, "SR": 133.0},
    ]
  save_campi(campi)
  return campi


def save_campi(campi_list):
  with open("campi.json", "w", encoding="utf-8") as f:
    json.dump(campi_list, f, indent=4, ensure_ascii=False)


# --- CALCOLO SD DA STABLEFORD ---
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


# --- CALCOLO HANDICAP ---
def calcola_hcp_corrente(df_in):
  if df_in.empty or "SD" not in df_in.columns:
    return 0.0, pd.DataFrame(), []

  validi = df_in[df_in["SD"].notna()].copy()
  validi = validi.sort_values("Data", ascending=False).reset_index(drop=True)
  ultimi_20 = validi.head(20)

  if len(ultimi_20) == 0:
    return 0.0, ultimi_20, []

  n_best = min(8, len(ultimi_20))
  migliori_8 = ultimi_20.nsmallest(n_best, "SD")
  return round(migliori_8["SD"].mean(), 1), ultimi_20, migliori_8.index.tolist()


def genera_trend_storico(df_in):
  if df_in.empty or "SD" not in df_in.columns:
    return pd.DataFrame()

  validi = df_in[df_in["SD"].notna()].copy()
  validi["Data_DT"] = pd.to_datetime(validi["Data"])
  validi_cron = validi.sort_values("Data_DT", ascending=True).reset_index(
      drop=True
  )

  hcp_progressivo = []
  for i in range(len(validi_cron)):
    window = validi_cron.iloc[max(0, i - 19) : i + 1]
    n_best = min(8, len(window))
    avg_hcp = round(window["SD"].nsmallest(n_best).mean(), 1)
    hcp_progressivo.append(avg_hcp)

  validi_cron["HCP_Storico"] = hcp_progressivo
  validi_cron["Data_Formatted"] = validi_cron["Data_DT"].dt.strftime("%d/%m/%Y")
  return validi_cron


# --- HEADER PRINCIPALE ---
st.title("⛳ Golf Handicap Tracker & Simulator")

df = load_data()
lista_campi = load_campi()

if not df.empty and "SD" in df.columns:
  hcp_attuale, ultimi_20, id_migliori = calcola_hcp_corrente(df)
  df_trend = genera_trend_storico(df)

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

  # --- TABS NAVIGAZIONE ---
  tab_dash, tab_sim, tab_campi, tab_reg = st.tabs([
      "📊 Dashboard & Trend HCP",
      "🔮 Simulazione Gara",
      "⛳ Campi di Gioco",
      "📋 Registro Gare Ufficiale",
  ])

  # ==========================================
  # TAB 1: DASHBOARD & TREND
  # ==========================================
  with tab_dash:
    st.subheader("📈 Trend Storico Handicap Index")

    if not df_trend.empty:
      fig = px.line(
          df_trend,
          x="Data_DT",
          y="HCP_Storico",
          markers=True,
          hover_data={
              "Gara": True,
              "SD": True,
              "Data_Formatted": True,
              "Data_DT": False,
          },
          labels={
              "Data_DT": "Data",
              "HCP_Storico": "Handicap Index",
              "Data_Formatted": "Data",
          },
          title="Evoluzione Handicap nel Tempo",
      )
      fig.update_traces(
          line_color="#2E7D32",
          line_width=3,
          marker=dict(size=8, color="#1B5E20"),
      )
      fig.update_layout(hovermode="x unified")
      st.plotly_chart(fig, use_container_width=True)

    st.subheader("🟢 Ultimi 20 Risultati (Evidenziati i Migliori 8)")

    # Colonne dalla B alla H dell'Excel (Data, Gara, Esecutore, Buche, Playing HCP, Stbl, SD)
    cols_target = [
        "Data",
        "Gara",
        "Esecutore",
        "Buche",
        "Playing HCP",
        "Stbl",
        "SD",
    ]
    cols_display = [c for c in cols_target if c in ultimi_20.columns]

    ultimi_20_show = ultimi_20.copy()
    if "Data" in ultimi_20_show.columns:
      ultimi_20_show["Data"] = pd.to_datetime(
          ultimi_20_show["Data"]
      ).dt.strftime("%d/%m/%Y")

    def highlight_b8(row):
      return (
          ["background-color: #d4edda; font-weight: bold;"] * len(row)
          if row.name in id_migliori
          else [""] * len(row)
      )

    st.dataframe(
        ultimi_20_show[cols_display].style.apply(highlight_b8, axis=1),
        use_container_width=True,
    )

  # ==========================================
  # TAB 2: SIMULAZIONE GARA
  # ==========================================
  with tab_sim:
    st.subheader("🔮 Simula Impatto Prossima Gara")

    nomi_campi_disponibili = (
        sorted(list(set([c["Nome"] for c in lista_campi])))
        if lista_campi
        else ["MONTEVEGLIO ASD"]
    )

    col_c1, col_c2 = st.columns(2)
    with col_c1:
      campo_selezionato = st.selectbox(
          "Seleziona Campo di Gioco", nomi_campi_disponibili
      )
    with col_c2:
      buche_selezionate = st.radio("Buche", [18, 9], horizontal=True)

    campo_info = next(
        (
            c
            for c in lista_campi
            if c["Nome"] == campo_selezionato
            and c["Buche"] == buche_selezionate
        ),
        None,
    )

    default_par = (
        campo_info["Par"]
        if campo_info
        else (72.0 if buche_selezionate == 18 else 32.0)
    )
    default_cr = (
        campo_info["CR"]
        if campo_info
        else (71.2 if buche_selezionate == 18 else 31.7)
    )
    default_sr = (
        campo_info["SR"]
        if campo_info
        else (124.0 if buche_selezionate == 18 else 114.0)
    )

    modalita = st.radio(
        "Modalità Inserimento:",
        ["Calcola da Punti Stableford", "Score Differential (SD) Manuale"],
        horizontal=True,
    )

    with st.form("simulazione_form"):
      nome_gara = st.text_input("Nome Gara Simulata", "Gara Prossima")
      col1, col2, col3 = st.columns(3)

      with col1:
        par = st.number_input("Par Campo", value=float(default_par))
      with col2:
        cr = st.number_input("Course Rating (CR)", value=float(default_cr))
      with col3:
        sr = st.number_input("Slope Rating (SR)", value=float(default_sr))

      pcc = st.number_input("PCC (-1.0 a +3.0)", value=0.0, step=0.5)

      if modalita == "Calcola da Punti Stableford":
        col3_s, col4_s = st.columns(2)
        with col3_s:
          stbl = st.number_input(
              "Punti Stableford Realizzati",
              value=36 if buche_selezionate == 18 else 18,
          )
        with col4_s:
          playing_hcp = st.number_input(
              "Playing HCP (Gara)", value=10.0 if buche_selezionate == 18 else 5.0
          )
        sd_simulato = calcola_sd_da_stableford(
            stbl, playing_hcp, par, cr, sr, buche_selezionate, pcc
        )
        st.info(f"💡 **Score Differential (SD) Calcolato:** `{sd_simulato}`")
      else:
        sd_simulato = st.number_input("SD Manuale", value=10.0, step=0.1)

      submit_sim = st.form_submit_button("🧪 Esegui Simulazione")

    if submit_sim:
      riga_sim = pd.DataFrame([{
          "Data": pd.Timestamp.now(),
          "Gara": f"[SIMULATA] {nome_gara}",
          "Esecutore": campo_selezionato,
          "Buche": buche_selezionate,
          "Playing HCP": (
              playing_hcp
              if modalita == "Calcola da Punti Stableford"
              else None
          ),
          "Stbl": stbl if modalita == "Calcola da Punti Stableford" else None,
          "SD": sd_simulato,
          "CR": cr,
          "Par": par,
          "SR": sr,
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
        data_scartato = pd.to_datetime(scartato["Data"]).strftime("%d/%m/%Y")
        st.caption(
            f"📌 Uscirà dal calcolo dei 20 risultati la gara: **{scartato['Gara']}**"
            f" del {data_scartato} (SD: {scartato['SD']})"
        )

  # ==========================================
  # TAB 3: ANAGRAFICA CAMPI DI GIOCO
  # ==========================================
  with tab_campi:
    st.subheader("⛳ Gestione Anagrafica Campi di Gioco")
    st.write(
        "Aggiungi o modifica i campi da gioco con i relativi valori di Par, CR"
        " e SR per 9 e 18 buche."
    )

    df_campi_curr = pd.DataFrame(lista_campi)

    df_campi_edited = st.data_editor(
        df_campi_curr,
        num_rows="dynamic",
        use_container_width=True,
        key="editor_campi",
    )

    if st.button("💾 Salva Modifiche Anagrafica Campi"):
      nuova_lista_campi = df_campi_edited.to_dict(orient="records")
      save_campi(nuova_lista_campi)
      st.success("Anagrafica campi aggiornata con successo!")
      st.rerun()

    st.markdown("---")
    st.markdown("##### ➕ Aggiungi un Nuovo Campo Veloce")
    with st.form("form_nuovo_campo"):
      c_nome = st.text_input("Nome Campo", "MONTEVEGLIO ASD")
      col_f1, col_f2, col_f3, col_f4 = st.columns(4)
      with col_f1:
        c_buche = st.selectbox("Buche Campo", [18, 9])
      with col_f2:
        c_par = st.number_input(
            "Par Campo", value=72.0 if c_buche == 18 else 32.0
        )
      with col_f3:
        c_cr = st.number_input(
            "Course Rating (CR)", value=71.2 if c_buche == 18 else 31.7
        )
      with col_f4:
        c_sr = st.number_input(
            "Slope Rating (SR)", value=124.0 if c_buche == 18 else 114.0
        )

      submit_campo = st.form_submit_button("➕ Salva Nuovo Campo")

      if submit_campo and c_nome:
        nuovo_c = {
            "Nome": c_nome.upper().strip(),
            "Buche": int(c_buche),
            "Par": float(c_par),
            "CR": float(c_cr),
            "SR": float(c_sr),
        }
        lista_campi.append(nuovo_c)
        save_campi(lista_campi)
        st.success(f"Campo '{c_nome}' aggiunto con successo!")
        st.rerun()

  # ==========================================
  # TAB 4: REGISTRO GARE & EDITOR EXCEL
  # ==========================================
  with tab_reg:
    st.subheader("📋 Registro Gare Ufficiale (Modifica Dati)")
    st.write(
        "Modifica direttamente le celle della tabella per correggere eventuali"
        " dati e clicca sul pulsante sottostante per salvare le modifiche nell'Excel."
    )

    df_editable = df.copy()
    if "Data" in df_editable.columns:
      df_editable["Data"] = pd.to_datetime(df_editable["Data"]).dt.strftime(
          "%Y-%m-%d"
      )

    edited_df = st.data_editor(
        df_editable,
        num_rows="dynamic",
        use_container_width=True,
        key="editor_excel",
    )

    if st.button("💾 Salva Tutte le Modifiche nel File Excel"):
      save_data(edited_df)
      st.success("File Excel aggiornato con successo!")
      st.rerun()

    st.markdown("---")
    st.markdown("##### ➕ Registra Nuova Gara Ufficiale")

    nomi_campi = sorted(list(set([c["Nome"] for c in lista_campi])))
    with st.form("inserimento_definitivo"):
      col_r1, col_r2 = st.columns(2)
      with col_r1:
        g_data = st.date_input("Data Gara")
        g_nome = st.text_input("Nome Gara Ufficiale")
        g_campo = st.selectbox("Campo di Gioco", nomi_campi)
      with col_r2:
        g_buche = st.selectbox("Buche Gara", [18, 9])
        g_stbl = st.number_input(
            "Punti Stableford", value=36 if g_buche == 18 else 18
        )
        g_playing = st.number_input(
            "Playing HCP", value=10.0 if g_buche == 18 else 5.0
        )

      c_match = next(
          (
              c
              for c in lista_campi
              if c["Nome"] == g_campo and c["Buche"] == g_buche
          ),
          None,
      )
      p_par = (
          c_match["Par"] if c_match else (72.0 if g_buche == 18 else 32.0)
      )
      p_cr = c_match["CR"] if c_match else (71.2 if g_buche == 18 else 31.7)
      p_sr = c_match["SR"] if c_match else (124.0 if g_buche == 18 else 114.0)

      sd_calcolato = calcola_sd_da_stableford(
          g_stbl, g_playing, p_par, p_cr, p_sr, g_buche
      )

      st.info(f"SD Calcolato per la nuova gara: `{sd_calcolato}`")
      submit_save = st.form_submit_button("💾 Registra Gara in Excel")

      if submit_save and g_nome:
        nuova_riga = pd.DataFrame([{
            "Data": pd.to_datetime(g_data).strftime("%Y-%m-%d"),
            "Gara": g_nome,
            "Esecutore": g_campo,
            "Buche": g_buche,
            "Playing HCP": g_playing,
            "Stbl": g_stbl,
            "SD": sd_calcolato,
            "CR": p_cr,
            "Par": p_par,
            "SR": p_sr,
        }])
        df_finale = pd.concat([nuova_riga, df], ignore_index=True)
        save_data(df_finale)
        st.success("Nuova gara salvata con successo!")
        st.rerun()

else:
  st.error(
      "File Excel non trovato o privo della struttura corretta. Assicurati che"
      " 'Handicap_2026.xlsx' sia caricato correttamente."
  )
