import json
import os
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Golf Handicap Tracker", page_icon="⛳", layout="wide"
)

NOME_FILE_EXCEL = "Handicap_2026.xlsx"
NOME_FILE_CAMPI = "campi.json"


# --- CARICAMENTO E SALVATAGGIO DATI EXCEL ---
@st.cache_data(ttl=1)
def load_data():
  filename = None
  for fname in [NOME_FILE_EXCEL, "Handicap 2026.xlsx"]:
    if os.path.exists(fname):
      filename = fname
      break

  if not filename:
    return pd.DataFrame()

  df = pd.read_excel(filename, sheet_name="Foglio2")
  if "Data" in df.columns and "SD" in df.columns:
    df["Data"] = pd.to_datetime(df["Data"])
    return df

  df_h1 = pd.read_excel(filename, sheet_name="Foglio2", header=1)
  if "Data" in df_h1.columns and "SD" in df_h1.columns:
    df_h1["Data"] = pd.to_datetime(df_h1["Data"])
    return df_h1

  return df


def save_data(df_to_save):
  df_clean = df_to_save.copy()
  if "Data" in df_clean.columns:
    df_clean["Data"] = pd.to_datetime(df_clean["Data"]).dt.strftime("%Y-%m-%d")

  df_clean.to_excel(NOME_FILE_EXCEL, sheet_name="Foglio2", index=False)
  st.cache_data.clear()


# --- ANAGRAFICA CAMPI DI GIOCO ---
def load_campi():
  if os.path.exists(NOME_FILE_CAMPI):
    try:
      with open(NOME_FILE_CAMPI, "r", encoding="utf-8") as f:
        return json.load(f)
    except Exception:
      pass

  default_db = [{
      "Circolo": "MONTEVEGLIO ASD",
      "Percorso": "18 Buche",
      "Buche": 18,
      "Par": 64.0,
      "Tees": {
          "Giallo": {"CR": 63.4, "SR": 114.0},
          "Verde": {"CR": 60.6, "SR": 108.0},
          "Rosso": {"CR": 63.6, "SR": 108.0},
      },
  }]
  save_campi(default_db)
  return default_db


def save_campi(campi_list):
  with open(NOME_FILE_CAMPI, "w", encoding="utf-8") as f:
    json.dump(campi_list, f, indent=2, ensure_ascii=False)


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


# --- CALCOLO HANDICAP DI GIOCO SPETTANTE ---
def calcola_playing_hcp_automatico(hcp_index, cr, sr, par, buche=18):
  if buche == 9:
    cr_eff = cr * 2 if cr < 50 else cr
    par_eff = par * 2 if par < 50 else par
    ch = (hcp_index * (sr / 113)) + (cr_eff - par_eff)
    return max(0, round(ch / 2))
  else:
    ch = (hcp_index * (sr / 113)) + (cr - par)
    return max(0, round(ch))


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
db_campi = load_campi()

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

  # --- SCHEDE NAVIGAZIONE ---
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

    cols_b_h = [
        "Data",
        "Gara",
        "Esecutore",
        "Buche",
        "Playing HCP",
        "Stbl",
        "SD",
    ]
    cols_display = [c for c in cols_b_h if c in ultimi_20.columns]

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

    circoli_unici = sorted(list(set(c["Circolo"] for c in db_campi)))

    col_sel1, col_sel2, col_sel3 = st.columns(3)

    with col_sel1:
      idx_def_circolo = (
          circoli_unici.index("MONTEVEGLIO ASD")
          if "MONTEVEGLIO ASD" in circoli_unici
          else 0
      )
      sel_circolo = st.selectbox(
          "1. Seleziona Circolo", circoli_unici, index=idx_def_circolo
      )

    percorsi_disponibili = [c for c in db_campi if c["Circolo"] == sel_circolo]
    nomi_percorsi = [p["Percorso"] for p in percorsi_disponibili]

    with col_sel2:
      sel_percorso_nome = st.selectbox("2. Seleziona Percorso", nomi_percorsi)

    obj_percorso = next(
        p for p in percorsi_disponibili if p["Percorso"] == sel_percorso_nome
    )
    tees_disponibili = list(obj_percorso["Tees"].keys())

    with col_sel3:
      idx_def_tee = (
          tees_disponibili.index("Giallo") if "Giallo" in tees_disponibili else 0
      )
      sel_tee = st.selectbox(
          "3. Seleziona Tee di Partenza",
          tees_disponibili,
          index=idx_def_tee,
      )

    auto_par = float(obj_percorso["Par"])
    auto_buche = int(obj_percorso["Buche"])
    auto_cr = float(obj_percorso["Tees"][sel_tee]["CR"])
    auto_sr = float(obj_percorso["Tees"][sel_tee]["SR"])

    # Calcolo automatico del Playing HCP per il campo selezionato
    playing_hcp_suggerito = calcola_playing_hcp_automatico(
        hcp_attuale, auto_cr, auto_sr, auto_par, auto_buche
    )

    modalita = st.radio(
        "Modalità Inserimento Score:",
        ["Calcola da Punti Stableford", "Score Differential (SD) Manuale"],
        horizontal=True,
    )

    nome_gara = st.text_input("Nome Gara Simulata", "Sunday Cup")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
      st.text_input("Buche Campo", value=f"{auto_buche} buche", disabled=True)
    with col2:
      par = st.number_input("Par Campo", value=auto_par)
    with col3:
      cr = st.number_input("Course Rating (CR)", value=auto_cr)
    with col4:
      sr = st.number_input("Slope Rating (SR)", value=auto_sr)

    pcc = st.number_input("PCC (-1.0 a +3.0)", value=0.0, step=0.5)

    if modalita == "Calcola da Punti Stableford":
      col3_s, col4_s = st.columns(2)
      with col3_s:
        stbl = st.number_input(
            "Punti Stableford Realizzati",
            value=36 if auto_buche == 18 else 18,
        )
      with col4_s:
        playing_hcp = st.number_input(
            "Playing HCP (Calcolato in base al tuo HCP Index)",
            value=float(playing_hcp_suggerito),
        )

      sd_simulato = calcola_sd_da_stableford(
          stbl, playing_hcp, par, cr, sr, auto_buche, pcc
      )
      st.info(
          f"💡 **Score Differential (SD) Calcolato in Tempo Reale ({sel_tee}):**"
          f" `{sd_simulato}`"
      )
    else:
      sd_simulato = st.number_input("SD Manuale", value=10.0, step=0.1)

    if st.button("🧪 Esegui Simulazione Handicap"):
      riga_sim = pd.DataFrame([{
          "Data": pd.Timestamp.now(),
          "Gara": f"[SIMULATA] {nome_gara}",
          "Esecutore": f"{sel_circolo} ({sel_tee})",
          "Buche": auto_buche,
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
    st.subheader("⛳ Anagrafica Campi di Gioco Federgolf")
    st.write(
        f"Attualmente memorizzati **{len(db_campi)} percorsi ufficiali**."
        " Puoi modificare o aggiungere campi con i relativi Tee di partenza."
    )

    rows_editor = []
    for c in db_campi:
      for color_tee, vals in c.get("Tees", {}).items():
        rows_editor.append({
            "Circolo": c["Circolo"],
            "Percorso": c["Percorso"],
            "Buche": c["Buche"],
            "Par": c["Par"],
            "Tee": color_tee,
            "CR": vals.get("CR", 0.0),
            "SR": vals.get("SR", 0.0),
        })

    df_campi_curr = pd.DataFrame(rows_editor)
    df_campi_edited = st.data_editor(
        df_campi_curr,
        num_rows="dynamic",
        use_container_width=True,
        key="editor_campi",
    )

    if st.button("💾 Salva Modifiche Anagrafica Campi"):
      reconstructed = {}
      for _, r in df_campi_edited.iterrows():
        key = (r["Circolo"], r["Percorso"])
        if key not in reconstructed:
          reconstructed[key] = {
              "Circolo": r["Circolo"],
              "Percorso": r["Percorso"],
              "Buche": int(r["Buche"]),
              "Par": float(r["Par"]),
              "Tees": {},
          }
        reconstructed[key]["Tees"][r["Tee"]] = {
            "CR": float(r["CR"]),
            "SR": float(r["SR"]),
        }

      save_campi(list(reconstructed.values()))
      st.success("Anagrafica campi salvata con successo!")
      st.rerun()

  # ==========================================
  # TAB 4: REGISTRO GARE & EDITOR EXCEL
  # ==========================================
  with tab_reg:
    st.subheader("📋 Registro Gare Ufficiale (Modifica Dati)")
    st.write(
        "Modifica direttamente le celle dell'Excel e salva i risultati."
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

    col_r1, col_r2, col_r3 = st.columns(3)
    with col_r1:
      reg_circolo = st.selectbox(
          "Circolo Gara", circoli_unici, index=idx_def_circolo, key="reg_c"
      )

    reg_percorsi = [c for c in db_campi if c["Circolo"] == reg_circolo]
    reg_nomi_percorsi = [p["Percorso"] for p in reg_percorsi]

    with col_r2:
      reg_percorso_nome = st.selectbox(
          "Percorso Gara", reg_nomi_percorsi, key="reg_p"
      )

    reg_obj_p = next(
        p for p in reg_percorsi if p["Percorso"] == reg_percorso_nome
    )
    reg_tees = list(reg_obj_p["Tees"].keys())

    with col_r3:
      reg_idx_tee = reg_tees.index("Giallo") if "Giallo" in reg_tees else 0
      reg_tee = st.selectbox(
          "Tee di Partenza", reg_tees, index=reg_idx_tee, key="reg_t"
      )

    r_par = float(reg_obj_p["Par"])
    r_buche = int(reg_obj_p["Buche"])
    r_cr = float(reg_obj_p["Tees"][reg_tee]["CR"])
    r_sr = float(reg_obj_p["Tees"][reg_tee]["SR"])

    reg_playing_suggerito = calcola_playing_hcp_automatico(
        hcp_attuale, r_cr, r_sr, r_par, r_buche
    )

    with st.form("inserimento_definitivo"):
      col_a, col_b = st.columns(2)
      with col_a:
        g_data = st.date_input("Data Gara")
        g_nome = st.text_input("Nome Gara Ufficiale", "Sunday Cup")
      with col_b:
        g_stbl = st.number_input(
            "Punti Stableford", value=36 if r_buche == 18 else 18
        )
        g_playing = st.number_input(
            "Playing HCP", value=float(reg_playing_suggerito)
        )

      sd_calc_reg = calcola_sd_da_stableford(
          g_stbl, g_playing, r_par, r_cr, r_sr, r_buche
      )

      st.info(
          f"Campo selezionato: **{reg_circolo} - {reg_percorso_nome} ({reg_tee})**"
          f" | SD Calcolato: `{sd_calc_reg}`"
      )

      submit_save = st.form_submit_button("💾 Registra Gara in Excel")

      if submit_save and g_nome:
        nuova_riga = pd.DataFrame([{
            "Data": pd.to_datetime(g_data).strftime("%Y-%m-%d"),
            "Gara": g_nome,
            "Esecutore": f"{reg_circolo} ({reg_tee})",
            "Buche": r_buche,
            "Playing HCP": g_playing,
            "Stbl": g_stbl,
            "SD": sd_calc_reg,
            "CR": r_cr,
            "Par": r_par,
            "SR": r_sr,
        }])
        df_finale = pd.concat([nuova_riga, df], ignore_index=True)
        save_data(df_finale)
        st.success("Nuova gara salvata con successo!")
        st.rerun()

else:
  st.error(
      "File Excel non trovato. Assicurati che 'Handicap_2026.xlsx' sia presente"
      " su GitHub."
  )
