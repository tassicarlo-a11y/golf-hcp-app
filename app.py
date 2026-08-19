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

# --- DATABASE UFFICIALE FEDERGOLF (839 PERCORSI ESTRATTI DAL PDF) ---
CAMPI_FEDERGOLF_DEFAULT = [
    {"Nome": "ACAYA - 18 Buche Par 71", "Buche": 18, "Par": 71.0, "CR": 71.8, "SR": 130.0},
    {"Nome": "ACAYA - 18 Buche Par 72", "Buche": 18, "Par": 72.0, "CR": 72.8, "SR": 133.0},
    {"Nome": "ACAYA - Prime Nove", "Buche": 9, "Par": 36.0, "CR": 36.4, "SR": 125.0},
    {"Nome": "ACAYA - Seconde Nove Par 35", "Buche": 9, "Par": 35.0, "CR": 35.4, "SR": 135.0},
    {"Nome": "ACAYA - Seconde Nove Par 36", "Buche": 9, "Par": 36.0, "CR": 36.4, "SR": 141.0},
    {"Nome": "ACQUABONA - 18 Buche", "Buche": 18, "Par": 68.0, "CR": 68.4, "SR": 132.0},
    {"Nome": "ACQUABONA - 9 Buche", "Buche": 9, "Par": 34.0, "CR": 34.2, "SR": 132.0},
    {"Nome": "ALBISOLA - 18 Buche par 64", "Buche": 18, "Par": 64.0, "CR": 62.4, "SR": 105.0},
    {"Nome": "ALBISOLA - 18 buche par 65", "Buche": 18, "Par": 65.0, "CR": 62.6, "SR": 107.0},
    {"Nome": "ALBISOLA - 18 buche par 66", "Buche": 18, "Par": 66.0, "CR": 62.8, "SR": 108.0},
    {"Nome": "ALBISOLA - 9 Buche par 32", "Buche": 9, "Par": 32.0, "CR": 31.2, "SR": 105.0},
    {"Nome": "ALBISOLA - 9 buche par 33", "Buche": 9, "Par": 33.0, "CR": 31.4, "SR": 108.0},
    {"Nome": "ALPINO - 18 Buche", "Buche": 18, "Par": 69.0, "CR": 66.1, "SR": 131.0},
    {"Nome": "ALPINO - Prime Nove", "Buche": 9, "Par": 35.0, "CR": 33.1, "SR": 132.0},
    {"Nome": "ALTA BADIA - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 69.5, "SR": 129.0},
    {"Nome": "ALTA BADIA - 9 Buche", "Buche": 9, "Par": 36.0, "CR": 34.8, "SR": 129.0},
    {"Nome": "AMBROSIANO - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 71.9, "SR": 131.0},
    {"Nome": "AMBROSIANO - Prime Nove", "Buche": 9, "Par": 36.0, "CR": 35.9, "SR": 131.0},
    {"Nome": "AMBROSIANO - Seconde Nove", "Buche": 9, "Par": 36.0, "CR": 36.0, "SR": 130.0},
    {"Nome": "ANTOGNOLLA - 18 buche", "Buche": 18, "Par": 71.0, "CR": 71.0, "SR": 131.0},
    {"Nome": "ANTOGNOLLA - 9 Buche Misto", "Buche": 9, "Par": 35.0, "CR": 35.0, "SR": 132.0},
    {"Nome": "ANTOGNOLLA - Prime Nove", "Buche": 9, "Par": 36.0, "CR": 35.1, "SR": 136.0},
    {"Nome": "ANTOGNOLLA - Seconde Nove", "Buche": 9, "Par": 35.0, "CR": 35.9, "SR": 126.0},
    {"Nome": "AOSTA ARSANIERES - 18 Buche", "Buche": 18, "Par": 62.0, "CR": 58.6, "SR": 106.0},
    {"Nome": "AOSTA ARSANIERES - 9 Buche", "Buche": 9, "Par": 31.0, "CR": 29.3, "SR": 106.0},
    {"Nome": "AOSTA BRISSOGNE - 18 Buche", "Buche": 18, "Par": 60.0, "CR": 58.4, "SR": 102.0},
    {"Nome": "AOSTA BRISSOGNE - 9 Buche", "Buche": 9, "Par": 30.0, "CR": 29.2, "SR": 102.0},
    {"Nome": "APPIANO - GOLF & COUNTRY - 18 buche Appiano", "Buche": 18, "Par": 70.0, "CR": 72.4, "SR": 117.0},
    {"Nome": "APPIANO - GOLF & COUNTRY - 18 Buche Carezza", "Buche": 18, "Par": 70.0, "CR": 64.8, "SR": 117.0},
    {"Nome": "APPIANO - GOLF & COUNTRY - 9 Buche Appiano", "Buche": 9, "Par": 35.0, "CR": 36.2, "SR": 117.0},
    {"Nome": "APPIANO - GOLF & COUNTRY - 9 Buche Carezza", "Buche": 9, "Par": 35.0, "CR": 32.4, "SR": 117.0},
    {"Nome": "ARCHI CLAUDIO - 18 Buche Provvisorio 2025", "Buche": 18, "Par": 64.0, "CR": 62.4, "SR": 102.0},
    {"Nome": "ARCHI CLAUDIO - Provvisorio 2025 - Buca Nuova", "Buche": 18, "Par": 64.0, "CR": 62.2, "SR": 103.0},
    {"Nome": "ARCHI CLAUDIO - 18 Buche Provvisorio 2026", "Buche": 18, "Par": 64.0, "CR": 61.6, "SR": 96.0},
    {"Nome": "ARCHI CLAUDIO - 9 Buche Provvisorio 2025", "Buche": 9, "Par": 32.0, "CR": 31.2, "SR": 102.0},
    {"Nome": "ARCHI CLAUDIO - 9 Buche Provvisorio 2025 - Buca Nuova", "Buche": 9, "Par": 32.0, "CR": 31.1, "SR": 103.0},
    {"Nome": "ARCHI CLAUDIO - 9 Buche Provvisorio 2026", "Buche": 9, "Par": 32.0, "CR": 30.8, "SR": 96.0},
    {"Nome": "ARENZANO PINETA - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 67.9, "SR": 126.0},
    {"Nome": "ARENZANO PINETA - Nove Buche", "Buche": 9, "Par": 36.0, "CR": 34.0, "SR": 126.0},
    {"Nome": "ARGENTA - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 71.9, "SR": 125.0},
    {"Nome": "ARGENTA - Prime Nove", "Buche": 9, "Par": 36.0, "CR": 35.7, "SR": 127.0},
    {"Nome": "ARGENTA - Seconde Nove", "Buche": 9, "Par": 36.0, "CR": 36.2, "SR": 122.0},
    {"Nome": "ARGENTARIO - 18 Buche", "Buche": 18, "Par": 71.0, "CR": 72.4, "SR": 138.0},
    {"Nome": "ARGENTARIO - Prime Nove", "Buche": 9, "Par": 36.0, "CR": 36.0, "SR": 141.0},
    {"Nome": "ARGENTARIO - Seconde Nove", "Buche": 9, "Par": 35.0, "CR": 36.4, "SR": 135.0},
    {"Nome": "ARONA - 18 buche", "Buche": 18, "Par": 70.0, "CR": 65.9, "SR": 125.0},
    {"Nome": "ARONA - Prime Nove", "Buche": 9, "Par": 35.0, "CR": 32.7, "SR": 127.0},
    {"Nome": "ARZAGA - Gary P. 1-9", "Buche": 9, "Par": 36.0, "CR": 35.0, "SR": 128.0},
    {"Nome": "ARZAGA - Gary Player", "Buche": 18, "Par": 72.0, "CR": 70.0, "SR": 128.0},
    {"Nome": "ARZAGA - GP1+JN1", "Buche": 18, "Par": 72.0, "CR": 70.9, "SR": 125.0},
    {"Nome": "ARZAGA - GP1+JN2", "Buche": 18, "Par": 72.0, "CR": 70.8, "SR": 129.0},
    {"Nome": "ARZAGA - Jack N. 10-18", "Buche": 9, "Par": 36.0, "CR": 35.8, "SR": 129.0},
    {"Nome": "ARZAGA - Jack N. 1-9", "Buche": 9, "Par": 36.0, "CR": 35.9, "SR": 124.0},
    {"Nome": "ARZAGA - Jack Nicklaus", "Buche": 18, "Par": 72.0, "CR": 71.7, "SR": 126.0},
    {"Nome": "ASIAGO - 18 Buche", "Buche": 18, "Par": 71.0, "CR": 70.2, "SR": 124.0},
    {"Nome": "ASIAGO - 18 Buche 2025", "Buche": 18, "Par": 71.0, "CR": 70.2, "SR": 124.0},
    {"Nome": "ASIAGO - 18 Buche Provvisorio 2025", "Buche": 18, "Par": 69.0, "CR": 68.5, "SR": 120.0},
    {"Nome": "ASIAGO - Percorso Invernale Misto", "Buche": 9, "Par": 35.0, "CR": 35.3, "SR": 124.0},
    {"Nome": "ASIAGO - Percorso Invernale Misto 2 Volte", "Buche": 18, "Par": 70.0, "CR": 70.6, "SR": 124.0},
    {"Nome": "ASIAGO - Prime Nove", "Buche": 9, "Par": 36.0, "CR": 35.3, "SR": 121.0},
    {"Nome": "ASIAGO - Prime Nove 2 Volte", "Buche": 18, "Par": 72.0, "CR": 70.6, "SR": 121.0},
    {"Nome": "ASIAGO - Prime Nove 2025", "Buche": 9, "Par": 36.0, "CR": 35.0, "SR": 122.0},
    {"Nome": "ASIAGO - Seconde Nove", "Buche": 9, "Par": 35.0, "CR": 34.9, "SR": 126.0},
    {"Nome": "ASIAGO - Seconde Nove 2025", "Buche": 9, "Par": 35.0, "CR": 35.2, "SR": 126.0},
    {"Nome": "ASIAGO - Seconde Nove Provvisorio 2025", "Buche": 9, "Par": 33.0, "CR": 33.2, "SR": 118.0},
    {"Nome": "ASOLO - Giallo-Giallo", "Buche": 18, "Par": 72.0, "CR": 71.6, "SR": 132.0},
    {"Nome": "ASOLO - Giallo-Verde", "Buche": 18, "Par": 72.0, "CR": 71.6, "SR": 131.0},
    {"Nome": "ASOLO - Giallo-Verde Provvisorio 2024", "Buche": 18, "Par": 70.0, "CR": 70.0, "SR": 127.0},
    {"Nome": "ASOLO - Percorso Giallo 9 buche", "Buche": 9, "Par": 36.0, "CR": 35.8, "SR": 132.0},
    {"Nome": "ASOLO - Percorso Rosso 9 buche", "Buche": 9, "Par": 36.0, "CR": 35.8, "SR": 137.0},
    {"Nome": "ASOLO - Percorso Verde 9 buche", "Buche": 9, "Par": 36.0, "CR": 35.7, "SR": 133.0},
    {"Nome": "ASOLO - Rosso-Giallo", "Buche": 18, "Par": 72.0, "CR": 71.7, "SR": 134.0},
    {"Nome": "ASOLO - Rosso-Giallo Provvisorio 2025", "Buche": 18, "Par": 71.0, "CR": 70.6, "SR": 131.0},
    {"Nome": "ASOLO - Rosso-Verde", "Buche": 18, "Par": 72.0, "CR": 71.6, "SR": 135.0},
    {"Nome": "ASOLO - Rosso-Verde Provvisorio 2025", "Buche": 18, "Par": 71.0, "CR": 70.5, "SR": 132.0},
    {"Nome": "ASOLO - Trofeo Rocca d'Asolo 2024", "Buche": 18, "Par": 70.0, "CR": 70.1, "SR": 130.0},
    {"Nome": "BAGNAIA - 18 Buche", "Buche": 18, "Par": 71.0, "CR": 71.2, "SR": 141.0},
    {"Nome": "BAGNAIA - Prime Nove", "Buche": 9, "Par": 36.0, "CR": 35.8, "SR": 147.0},
    {"Nome": "BAGNAIA - Seconde Nove", "Buche": 9, "Par": 35.0, "CR": 35.4, "SR": 135.0},
    {"Nome": "BARIALTO GOLF - 18 Buche", "Buche": 18, "Par": 70.0, "CR": 69.8, "SR": 125.0},
    {"Nome": "BARIALTO GOLF - Prime Nove", "Buche": 9, "Par": 36.0, "CR": 35.0, "SR": 128.0},
    {"Nome": "BARIALTO GOLF - Seconde Nove", "Buche": 9, "Par": 34.0, "CR": 34.8, "SR": 122.0},
    {"Nome": "BARLASSINA - Campionato", "Buche": 18, "Par": 72.0, "CR": 71.0, "SR": 132.0},
    {"Nome": "BARLASSINA - Prime Nove", "Buche": 9, "Par": 36.0, "CR": 35.3, "SR": 130.0},
    {"Nome": "BARLASSINA - Seconde Nove", "Buche": 9, "Par": 36.0, "CR": 35.7, "SR": 133.0},
    {"Nome": "BELLOSGUARDO - Monnalisa New 2023", "Buche": 18, "Par": 71.0, "CR": 71.9, "SR": 124.0},
    {"Nome": "BELLOSGUARDO - Monnalisa New Prime Nove 2023", "Buche": 9, "Par": 35.0, "CR": 35.9, "SR": 127.0},
    {"Nome": "BELLOSGUARDO - Monnalisa New Seconde Nove 2023", "Buche": 9, "Par": 36.0, "CR": 36.0, "SR": 121.0},
    {"Nome": "BERGAMO ALBENZA - Blu", "Buche": 9, "Par": 36.0, "CR": 35.8, "SR": 137.0},
    {"Nome": "BERGAMO ALBENZA - Blu-Giallo", "Buche": 18, "Par": 72.0, "CR": 71.8, "SR": 136.0},
    {"Nome": "BERGAMO ALBENZA - Giallo", "Buche": 9, "Par": 36.0, "CR": 36.0, "SR": 135.0},
    {"Nome": "BERGAMO ALBENZA - Rosso", "Buche": 9, "Par": 36.0, "CR": 35.8, "SR": 136.0},
    {"Nome": "BERGAMO ALBENZA - Rosso-Blu", "Buche": 18, "Par": 72.0, "CR": 70.9, "SR": 134.0},
    {"Nome": "BERGAMO ALBENZA - Rosso-Giallo", "Buche": 18, "Par": 72.0, "CR": 71.1, "SR": 133.0},
    {"Nome": "BIELLA BETULLE - 18 buche 2018", "Buche": 18, "Par": 73.0, "CR": 72.9, "SR": 142.0},
    {"Nome": "BIELLA BETULLE - Le Betulle", "Buche": 18, "Par": 73.0, "CR": 72.8, "SR": 140.0},
    {"Nome": "BIELLA BETULLE - prime 9", "Buche": 9, "Par": 36.0, "CR": 36.4, "SR": 138.0},
    {"Nome": "BOGLIACO - 18 Buche", "Buche": 18, "Par": 70.0, "CR": 68.0, "SR": 137.0},
    {"Nome": "BOGLIACO - Prime Nove", "Buche": 9, "Par": 35.0, "CR": 33.5, "SR": 138.0},
    {"Nome": "BOGLIACO - Seconde Nove", "Buche": 9, "Par": 35.0, "CR": 34.5, "SR": 135.0},
    {"Nome": "BOGOGNO - 1° Nove - Conte", "Buche": 9, "Par": 36.0, "CR": 36.8, "SR": 137.0},
    {"Nome": "BOGOGNO - 2° Nove-Conte", "Buche": 9, "Par": 36.0, "CR": 36.3, "SR": 125.0},
    {"Nome": "BOGOGNO - Bonora", "Buche": 18, "Par": 72.0, "CR": 72.7, "SR": 138.0},
    {"Nome": "BOGOGNO - del Conte", "Buche": 18, "Par": 72.0, "CR": 73.1, "SR": 131.0},
    {"Nome": "BOGOGNO - Prime Nove - Bonora", "Buche": 9, "Par": 36.0, "CR": 35.9, "SR": 132.0},
    {"Nome": "BOGOGNO - Seconde Nove-Bonora", "Buche": 9, "Par": 36.0, "CR": 36.6, "SR": 140.0},
    {"Nome": "BOLLINA - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 69.7, "SR": 125.0},
    {"Nome": "BOLLINA - 9 Buche", "Buche": 9, "Par": 36.0, "CR": 34.9, "SR": 125.0},
    {"Nome": "BOLOGNA - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 71.7, "SR": 127.0},
    {"Nome": "BOLOGNA - PAR 71 CAMPIONATO", "Buche": 18, "Par": 71.0, "CR": 71.3, "SR": 125.0},
    {"Nome": "BOLOGNA - par 71.", "Buche": 18, "Par": 71.0, "CR": 70.9, "SR": 126.0},
    {"Nome": "BOLOGNA - Prime Nove", "Buche": 9, "Par": 36.0, "CR": 35.7, "SR": 135.0},
    {"Nome": "BOLOGNA - Seconde Nove", "Buche": 9, "Par": 36.0, "CR": 36.1, "SR": 118.0},
    {"Nome": "BOLOGNA - seconde nove par 35", "Buche": 9, "Par": 35.0, "CR": 35.2, "SR": 117.0},
    {"Nome": "BORGO CAMUZZAGO - Nove Buche par 27", "Buche": 9, "Par": 27.0, "CR": 28.9, "SR": 97.0},
    {"Nome": "BORGO CAMUZZAGO - 18 buche", "Buche": 18, "Par": 64.0, "CR": 62.1, "SR": 109.0},
    {"Nome": "BORGO CAMUZZAGO - 18 Buche par54", "Buche": 18, "Par": 54.0, "CR": 57.8, "SR": 97.0},
    {"Nome": "BORGO CAMUZZAGO - 9 buche", "Buche": 9, "Par": 32.0, "CR": 31.1, "SR": 109.0},
    {"Nome": "BORMIO SSD - 18 Buche", "Buche": 18, "Par": 66.0, "CR": 60.8, "SR": 107.0},
    {"Nome": "BORMIO SSD - 18 buche Par 62", "Buche": 18, "Par": 62.0, "CR": 58.8, "SR": 107.0},
    {"Nome": "BORMIO SSD - 9 Buche", "Buche": 9, "Par": 33.0, "CR": 30.4, "SR": 107.0},
    {"Nome": "BORMIO SSD - 9 buche Par 31", "Buche": 9, "Par": 31.0, "CR": 29.4, "SR": 107.0},
    {"Nome": "BOTANIC SA CUBA - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 64.8, "SR": 116.0},
    {"Nome": "BOTANIC SA CUBA - 9 Buche", "Buche": 9, "Par": 36.0, "CR": 32.4, "SR": 116.0},
    {"Nome": "BOVES - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 69.2, "SR": 125.0},
    {"Nome": "BOVES - 18 Buche Provvisorio 2024", "Buche": 18, "Par": 72.0, "CR": 69.3, "SR": 125.0},
    {"Nome": "BOVES - 9 BUCHE MISTO 2 VOLTE", "Buche": 9, "Par": 72.0, "CR": 70.6, "SR": 123.0},
    {"Nome": "BOVES - Prime Nove", "Buche": 9, "Par": 36.0, "CR": 34.8, "SR": 116.0},
    {"Nome": "BOVES - Seconde Nove", "Buche": 9, "Par": 36.0, "CR": 34.5, "SR": 134.0},
    {"Nome": "BRIANZA - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 69.0, "SR": 129.0},
    {"Nome": "BRIANZA - Prime Nove", "Buche": 9, "Par": 35.0, "CR": 33.5, "SR": 129.0},
    {"Nome": "BRIANZA - Seconde Nove", "Buche": 9, "Par": 37.0, "CR": 35.4, "SR": 129.0},
    {"Nome": "CA' AMATA - 18 Buche", "Buche": 18, "Par": 71.0, "CR": 72.2, "SR": 136.0},
    {"Nome": "CA' AMATA - 18 Buche Provvisorio 2025", "Buche": 18, "Par": 71.0, "CR": 71.7, "SR": 135.0},
    {"Nome": "CA' AMATA - Prime Nove", "Buche": 9, "Par": 36.0, "CR": 35.4, "SR": 135.0},
    {"Nome": "CA' AMATA - Prime Nove Provvisorio 2025", "Buche": 9, "Par": 36.0, "CR": 34.9, "SR": 134.0},
    {"Nome": "CA' AMATA - Seconde Nove", "Buche": 9, "Par": 35.0, "CR": 36.0, "SR": 133.0},
    {"Nome": "CA' NAVE SSD - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 72.0, "SR": 128.0},
    {"Nome": "CA' NAVE SSD - Prime Nove", "Buche": 9, "Par": 36.0, "CR": 34.8, "SR": 122.0},
    {"Nome": "CA' NAVE SSD - Seconde Nove", "Buche": 9, "Par": 36.0, "CR": 34.2, "SR": 120.0},
    {"Nome": "CA' ULIVI - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 70.0, "SR": 122.0},
    {"Nome": "CA' ULIVI - Mirabello", "Buche": 18, "Par": 62.0, "CR": 60.3, "SR": 108.0},
    {"Nome": "CA' ULIVI - Mirabello 9 B.", "Buche": 9, "Par": 31.0, "CR": 30.1, "SR": 108.0},
    {"Nome": "CAMPODOGLIO - 18 Buche Easy 2024", "Buche": 18, "Par": 70.0, "CR": 68.2, "SR": 114.0},
    {"Nome": "CAMPODOGLIO - 18 Buche Mixed 2024", "Buche": 18, "Par": 71.0, "CR": 70.7, "SR": 120.0},
    {"Nome": "CAMPODOGLIO - 18 Buche New 2024", "Buche": 18, "Par": 70.0, "CR": 70.6, "SR": 121.0},
    {"Nome": "CAMPODOGLIO - 18 Buche Old 2024", "Buche": 18, "Par": 72.0, "CR": 71.0, "SR": 121.0},
    {"Nome": "CAMPODOGLIO - 9 Buche Easy 2024", "Buche": 9, "Par": 35.0, "CR": 34.1, "SR": 114.0},
    {"Nome": "CAMPODOGLIO - 9 Buche Mixed 2024", "Buche": 9, "Par": 36.0, "CR": 35.6, "SR": 122.0},
    {"Nome": "CAMPODOGLIO - 9 Buche New 2024", "Buche": 9, "Par": 35.0, "CR": 35.3, "SR": 121.0},
    {"Nome": "CAMPODOGLIO - 9 Buche Old 2024", "Buche": 9, "Par": 36.0, "CR": 35.5, "SR": 121.0},
    {"Nome": "CANSIGLIO - 18 Buche", "Buche": 18, "Par": 71.0, "CR": 70.0, "SR": 129.0},
    {"Nome": "CANSIGLIO - Prime Nove", "Buche": 9, "Par": 35.0, "CR": 32.6, "SR": 122.0},
    {"Nome": "CANSIGLIO - Seconde Nove", "Buche": 9, "Par": 36.0, "CR": 33.9, "SR": 128.0},
    {"Nome": "CAORLE - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 68.2, "SR": 122.0},
    {"Nome": "CAORLE - Prime Nove", "Buche": 9, "Par": 36.0, "CR": 33.9, "SR": 123.0},
    {"Nome": "CAORLE - Seconde Nove", "Buche": 9, "Par": 36.0, "CR": 34.3, "SR": 121.0},
    {"Nome": "CARIMATE - 18 Buche", "Buche": 18, "Par": 71.0, "CR": 69.7, "SR": 128.0},
    {"Nome": "CARIMATE - Prime Nove", "Buche": 9, "Par": 36.0, "CR": 34.7, "SR": 128.0},
    {"Nome": "CARIMATE - Seconde Nove", "Buche": 9, "Par": 35.0, "CR": 34.9, "SR": 128.0},
    {"Nome": "CASALUNGA - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 70.6, "SR": 125.0},
    {"Nome": "CASALUNGA - 18 Buche Storm 2023 D", "Buche": 18, "Par": 72.0, "CR": 70.4, "SR": 126.0},
    {"Nome": "CASALUNGA - 9 Buche", "Buche": 9, "Par": 36.0, "CR": 35.3, "SR": 125.0},
    {"Nome": "CASALUNGA - 9 Buche Storm 2023 D", "Buche": 9, "Par": 36.0, "CR": 35.2, "SR": 126.0},
    {"Nome": "CASENTINO - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 67.5, "SR": 116.0},
    {"Nome": "CASENTINO - 9 Buche", "Buche": 9, "Par": 36.0, "CR": 33.7, "SR": 116.0},
    {"Nome": "CASTELCONTURBIA - 9 buche Azzurro", "Buche": 9, "Par": 36.0, "CR": 35.9, "SR": 138.0},
    {"Nome": "CASTELCONTURBIA - 9 buche Giallo", "Buche": 9, "Par": 36.0, "CR": 35.6, "SR": 146.0},
    {"Nome": "CASTELCONTURBIA - 9 buche Rosso", "Buche": 9, "Par": 36.0, "CR": 37.0, "SR": 141.0},
    {"Nome": "CASTELCONTURBIA - Azzurro-Giallo", "Buche": 18, "Par": 72.0, "CR": 71.5, "SR": 142.0},
    {"Nome": "CASTELCONTURBIA - Azzurro-Rosso", "Buche": 18, "Par": 72.0, "CR": 73.0, "SR": 139.0},
    {"Nome": "CASTELCONTURBIA - Giallo-Rosso", "Buche": 18, "Par": 72.0, "CR": 72.6, "SR": 143.0},
    {"Nome": "CASTELFALFI - LAKE + Mountain Prime Nove", "Buche": 18, "Par": 73.0, "CR": 70.8, "SR": 138.0},
    {"Nome": "CASTELFALFI - LAKE 18 Buche", "Buche": 18, "Par": 74.0, "CR": 71.8, "SR": 144.0},
    {"Nome": "CASTELFALFI - LAKE 9 Buche", "Buche": 9, "Par": 37.0, "CR": 35.9, "SR": 144.0},
    {"Nome": "CASTELFALFI - Mountain 18 Buche", "Buche": 18, "Par": 72.0, "CR": 75.1, "SR": 150.0},
    {"Nome": "CASTELGANDOLFO - 18 Buche - Buca 4 Par 3", "Buche": 18, "Par": 71.0, "CR": 70.0, "SR": 128.0},
    {"Nome": "CASTELGANDOLFO - 18 Buche Provvisorio 2026", "Buche": 18, "Par": 71.0, "CR": 69.7, "SR": 130.0},
    {"Nome": "CASTELGANDOLFO - Campionato", "Buche": 18, "Par": 72.0, "CR": 71.5, "SR": 134.0},
    {"Nome": "CASTELGANDOLFO - Misto (10-11-12-13-14-6-7-8-9)", "Buche": 9, "Par": 36.0, "CR": 34.8, "SR": 131.0},
    {"Nome": "CASTELGANDOLFO - Misto (10-11-12-13-14-6-7-8-9) 2 Volte", "Buche": 18, "Par": 72.0, "CR": 69.6, "SR": 131.0},
    {"Nome": "CASTELGANDOLFO - Misto (10-17-12-13-14-6-7-8-9)", "Buche": 9, "Par": 34.0, "CR": 33.2, "SR": 123.0},
    {"Nome": "CASTELGANDOLFO - Misto (10-17-12-13-14-6-7-8-9) 2 Volte", "Buche": 18, "Par": 68.0, "CR": 66.4, "SR": 123.0},
    {"Nome": "CASTELGANDOLFO - Prime Nove", "Buche": 9, "Par": 36.0, "CR": 35.6, "SR": 138.0},
    {"Nome": "CASTELGANDOLFO - Prime Nove - Buca 4 Par 3", "Buche": 9, "Par": 35.0, "CR": 34.1, "SR": 127.0},
    {"Nome": "CASTELGANDOLFO - Seconde Nove", "Buche": 9, "Par": 36.0, "CR": 35.9, "SR": 129.0},
    {"Nome": "CASTELLARO - 18 Buche", "Buche": 18, "Par": 66.0, "CR": 63.6, "SR": 114.0},
    {"Nome": "CASTELLARO - 9 Buche", "Buche": 9, "Par": 33.0, "CR": 31.8, "SR": 114.0},
    {"Nome": "CASTELLO SPESSA - 18 Buche", "Buche": 18, "Par": 71.0, "CR": 67.2, "SR": 126.0},
    {"Nome": "CASTELLO SPESSA - 2 Volte 1°Nove", "Buche": 18, "Par": 70.0, "CR": 66.0, "SR": 127.0},
    {"Nome": "CASTELLO SPESSA - Prime Nove", "Buche": 9, "Par": 35.0, "CR": 33.0, "SR": 127.0},
    {"Nome": "CASTELLO SPESSA - Seconde Nove", "Buche": 9, "Par": 36.0, "CR": 34.2, "SR": 125.0},
    {"Nome": "CAVAGLIA' - 18 buche", "Buche": 18, "Par": 68.0, "CR": 63.8, "SR": 123.0},
    {"Nome": "CAVAGLIA' - 9 buche", "Buche": 9, "Par": 35.0, "CR": 32.6, "SR": 124.0},
    {"Nome": "CAVAGLIA' - 9 buche 1-9", "Buche": 9, "Par": 34.0, "CR": 31.8, "SR": 121.0},
    {"Nome": "CERRETO MIGLIANICO - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 70.4, "SR": 126.0},
    {"Nome": "CERRETO MIGLIANICO - Prime Nove", "Buche": 9, "Par": 36.0, "CR": 33.3, "SR": 119.0},
    {"Nome": "CERRETO MIGLIANICO - Seconde Nove", "Buche": 9, "Par": 36.0, "CR": 34.0, "SR": 119.0},
    {"Nome": "CERVIA - 9 Buche Blu", "Buche": 9, "Par": 36.0, "CR": 35.3, "SR": 127.0},
    {"Nome": "CERVIA - 9 buche Giallo", "Buche": 9, "Par": 35.0, "CR": 35.4, "SR": 126.0},
    {"Nome": "CERVIA - 9 Buche Rosso", "Buche": 9, "Par": 36.0, "CR": 36.5, "SR": 125.0},
    {"Nome": "CERVIA - Blu-Giallo", "Buche": 18, "Par": 71.0, "CR": 70.7, "SR": 127.0},
    {"Nome": "CERVIA - Giallo-Rosso", "Buche": 18, "Par": 71.0, "CR": 71.9, "SR": 126.0},
    {"Nome": "CERVIA - Percorso Open", "Buche": 18, "Par": 71.0, "CR": 71.9, "SR": 126.0},
    {"Nome": "CERVIA - Rosso-Blu", "Buche": 18, "Par": 72.0, "CR": 71.8, "SR": 126.0},
    {"Nome": "CERVINO - 18 Buche", "Buche": 18, "Par": 69.0, "CR": 68.0, "SR": 128.0},
    {"Nome": "CERVINO - 18 Buche Provvisorio 2026", "Buche": 18, "Par": 68.0, "CR": 64.6, "SR": 118.0},
    {"Nome": "CERVINO - 9 Buche Misto", "Buche": 9, "Par": 36.0, "CR": 34.5, "SR": 128.0},
    {"Nome": "CERVINO - 9 Buche Provvisorio 2026", "Buche": 9, "Par": 34.0, "CR": 32.3, "SR": 118.0},
    {"Nome": "CHERASCO - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 70.8, "SR": 131.0},
    {"Nome": "CHERASCO - Prime Nove", "Buche": 9, "Par": 36.0, "CR": 36.1, "SR": 130.0},
    {"Nome": "CHERASCO - Prime Nove 2 volte", "Buche": 18, "Par": 72.0, "CR": 72.2, "SR": 130.0},
    {"Nome": "CHERASCO - Seconde Nove", "Buche": 9, "Par": 36.0, "CR": 34.7, "SR": 131.0},
    {"Nome": "CHERASCO - Seconde Nove 2 volte", "Buche": 18, "Par": 72.0, "CR": 69.4, "SR": 131.0},
    {"Nome": "CILIEGI - 9 BUCHE", "Buche": 9, "Par": 36.0, "CR": 35.2, "SR": 129.0},
    {"Nome": "CILIEGI - CHERRIES B+V", "Buche": 18, "Par": 72.0, "CR": 68.4, "SR": 124.0},
    {"Nome": "CILIEGI - KING", "Buche": 18, "Par": 72.0, "CR": 70.4, "SR": 129.0},
    {"Nome": "CITTA' D'ASTI - 18 Buche", "Buche": 18, "Par": 66.0, "CR": 62.5, "SR": 105.0},
    {"Nome": "CITTA' D'ASTI - 9 BUCHE", "Buche": 9, "Par": 33.0, "CR": 31.3, "SR": 105.0},
    {"Nome": "CLAVIERE - Buche 18", "Buche": 18, "Par": 64.0, "CR": 66.0, "SR": 121.0},
    {"Nome": "CLAVIERE - 9 Buche", "Buche": 9, "Par": 32.0, "CR": 31.8, "SR": 115.0},
    {"Nome": "COLLI BERECI - 18 Buche", "Buche": 18, "Par": 70.0, "CR": 69.7, "SR": 131.0},
    {"Nome": "CONERO - 18 Buche", "Buche": 18, "Par": 71.0, "CR": 70.2, "SR": 129.0},
    {"Nome": "CORTINA SSD - 18 Buche", "Buche": 18, "Par": 70.0, "CR": 68.2, "SR": 124.0},
    {"Nome": "COSMOPOLITAN - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 72.0, "SR": 133.0},
    {"Nome": "CROARA SSD - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 68.7, "SR": 125.0},
    {"Nome": "CUS FERRARA - 18 Buche", "Buche": 18, "Par": 68.0, "CR": 65.2, "SR": 110.0},
    {"Nome": "DES ILES BORROMEES - 18 Buche 2023", "Buche": 18, "Par": 72.0, "CR": 70.7, "SR": 132.0},
    {"Nome": "DOLOMITI - 18 Buche", "Buche": 18, "Par": 73.0, "CR": 71.6, "SR": 127.0},
    {"Nome": "DUCATO - La Rocca 18 Buche", "Buche": 18, "Par": 72.0, "CR": 72.5, "SR": 133.0},
    {"Nome": "FIORANELLO - 18 Buche", "Buche": 18, "Par": 70.0, "CR": 69.0, "SR": 127.0},
    {"Nome": "FIRENZE UGOLINO - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 67.5, "SR": 125.0},
    {"Nome": "FOLGARIA - 2026-18 Buche par 72", "Buche": 18, "Par": 72.0, "CR": 65.7, "SR": 120.0},
    {"Nome": "FONTI - 18 Buche Par 72- 2023", "Buche": 18, "Par": 72.0, "CR": 71.4, "SR": 125.0},
    {"Nome": "FRANCIACORTA - BRUT+SATEN", "Buche": 18, "Par": 72.0, "CR": 71.4, "SR": 132.0},
    {"Nome": "FRASSANELLE - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 71.5, "SR": 131.0},
    {"Nome": "FRONDE - 18 Buche", "Buche": 18, "Par": 71.0, "CR": 70.8, "SR": 132.0},
    {"Nome": "GARDAGOLF - Bianco-Giallo", "Buche": 18, "Par": 71.0, "CR": 71.0, "SR": 131.0},
    {"Nome": "GARLENDA - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 72.2, "SR": 137.0},
    {"Nome": "GLOBALE JESOLO - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 70.7, "SR": 119.0},
    {"Nome": "GREEN CLUB LAINATE - 18 Buche 2023", "Buche": 18, "Par": 70.0, "CR": 68.3, "SR": 123.0},
    {"Nome": "IS ARENAS - 18 Buche 2022", "Buche": 18, "Par": 72.0, "CR": 73.6, "SR": 140.0},
    {"Nome": "IS MOLAS SSD - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 73.6, "SR": 132.0},
    {"Nome": "LIGNANO SSD - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 72.3, "SR": 135.0},
    {"Nome": "MARCO SIMONE - 2021 18 buche", "Buche": 18, "Par": 72.0, "CR": 72.1, "SR": 129.0},
    {"Nome": "MARGARA - La Guazzetta", "Buche": 18, "Par": 72.0, "CR": 73.0, "SR": 130.0},
    {"Nome": "MARGARA - Lolli Ghetti", "Buche": 18, "Par": 72.0, "CR": 72.5, "SR": 130.0},
    {"Nome": "MENAGGIO - 18 Buche", "Buche": 18, "Par": 70.0, "CR": 69.3, "SR": 133.0},
    {"Nome": "MILANO - 1/2", "Buche": 18, "Par": 72.0, "CR": 73.4, "SR": 129.0},
    {"Nome": "MODENA - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 72.2, "SR": 127.0},
    {"Nome": "MOLINETTO - 18 Buche", "Buche": 18, "Par": 71.0, "CR": 71.4, "SR": 130.0},
    {"Nome": "MONTEVEGLIO ASD - Nove Buche", "Buche": 9, "Par": 32.0, "CR": 31.7, "SR": 114.0},
    {"Nome": "MONTEVEGLIO ASD - 18 Buche", "Buche": 18, "Par": 64.0, "CR": 63.4, "SR": 114.0},
    {"Nome": "MONTICELLO - Blu", "Buche": 18, "Par": 72.0, "CR": 71.4, "SR": 132.0},
    {"Nome": "MONTICELLO - Rosso", "Buche": 18, "Par": 72.0, "CR": 72.1, "SR": 133.0},
    {"Nome": "NAZIONALE - Campionato", "Buche": 18, "Par": 72.0, "CR": 71.7, "SR": 142.0},
    {"Nome": "OLGIATA - Percorso Ovest Par 72", "Buche": 18, "Par": 72.0, "CR": 73.2, "SR": 134.0},
    {"Nome": "PADOVA - Padova 2025-Percorso Rosso-Blu", "Buche": 18, "Par": 72.0, "CR": 71.4, "SR": 124.0},
    {"Nome": "PARCO DE' MEDICI - Championship Bianco/Blu", "Buche": 18, "Par": 72.0, "CR": 71.1, "SR": 136.0},
    {"Nome": "PAVONIERE - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 73.0, "SR": 139.0},
    {"Nome": "PEVERO - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 73.5, "SR": 140.0},
    {"Nome": "PINETINA - 18 Buche", "Buche": 18, "Par": 70.0, "CR": 70.5, "SR": 126.0},
    {"Nome": "PUNTA ALA - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 73.3, "SR": 139.0},
    {"Nome": "RAPALLO - B.2 Par 5-B.5 Rid. Par 70", "Buche": 18, "Par": 70.0, "CR": 69.6, "SR": 121.0},
    {"Nome": "RIVIERA GOLF - 18 Buche", "Buche": 18, "Par": 70.0, "CR": 70.1, "SR": 122.0},
    {"Nome": "ROBINIE - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 71.7, "SR": 125.0},
    {"Nome": "ROMA ACQUASANTA - 18 Buche", "Buche": 18, "Par": 71.0, "CR": 70.7, "SR": 131.0},
    {"Nome": "ROYAL PARK ROVERI - Hurdzan Fry", "Buche": 18, "Par": 72.0, "CR": 74.6, "SR": 140.0},
    {"Nome": "ROYAL PARK ROVERI - Trent Jones", "Buche": 18, "Par": 72.0, "CR": 74.7, "SR": 143.0},
    {"Nome": "SAN DOMENICO - EGNAZIA - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 71.3, "SR": 128.0},
    {"Nome": "SAN VIGILIO - Benaco-Solferino", "Buche": 18, "Par": 72.0, "CR": 71.7, "SR": 124.0},
    {"Nome": "SANREMO ULIVI - 18 Buche", "Buche": 18, "Par": 69.0, "CR": 68.7, "SR": 119.0},
    {"Nome": "TERRE CONSOLI - 18 buche 2019", "Buche": 18, "Par": 72.0, "CR": 74.8, "SR": 133.0},
    {"Nome": "TOLCINASCO - Blu-Giallo", "Buche": 18, "Par": 72.0, "CR": 71.1, "SR": 133.0},
    {"Nome": "TORINO - Blu", "Buche": 18, "Par": 72.0, "CR": 73.7, "SR": 142.0},
    {"Nome": "TOSCANA - 18 Buche", "Buche": 18, "Par": 71.0, "CR": 70.5, "SR": 129.0},
    {"Nome": "UDINE - 18 buche 2018", "Buche": 18, "Par": 72.0, "CR": 74.4, "SR": 139.0},
    {"Nome": "VALTELLINA - 18 Buche", "Buche": 18, "Par": 71.0, "CR": 70.3, "SR": 133.0},
    {"Nome": "VARESE - Vecchio Monastero", "Buche": 18, "Par": 72.0, "CR": 71.2, "SR": 130.0},
    {"Nome": "VENEZIA - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 72.6, "SR": 138.0},
    {"Nome": "VERDURA - East", "Buche": 18, "Par": 73.0, "CR": 74.3, "SR": 136.0},
    {"Nome": "VERONA - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 72.0, "SR": 139.0},
    {"Nome": "VILLA CAROLINA SSD - La Marchesa", "Buche": 18, "Par": 72.0, "CR": 72.9, "SR": 130.0},
    {"Nome": "VILLA CONDULMER - 2026-18 buche Giallo/Blu", "Buche": 18, "Par": 71.0, "CR": 70.6, "SR": 130.0},
    {"Nome": "VILLA D'ESTE - 18 Buche", "Buche": 18, "Par": 69.0, "CR": 70.0, "SR": 129.0},
    {"Nome": "VILLA PARADISO SSD - 18 Buche 2024", "Buche": 18, "Par": 72.0, "CR": 72.8, "SR": 131.0},
    {"Nome": "ZOATE - 18 Buche", "Buche": 18, "Par": 72.0, "CR": 70.5, "SR": 144.0}
]


# --- CARICAMENTO E SALVATAGGIO DATI EXCEL ---
@st.cache_data(ttl=1)
def load_data():
  if os.path.exists(NOME_FILE_EXCEL):
    df = pd.read_excel(NOME_FILE_EXCEL, sheet_name="Foglio2")
    if "Data" in df.columns and "SD" in df.columns:
      df["Data"] = pd.to_datetime(df["Data"])
      return df

  # Fallback se il file ha lo spazio nel nome
  if os.path.exists("Handicap 2026.xlsx"):
    df = pd.read_excel("Handicap 2026.xlsx", sheet_name="Foglio2")
    if "Data" in df.columns and "SD" in df.columns:
      df["Data"] = pd.to_datetime(df["Data"])
      return df

  return pd.DataFrame()


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

  # Se non esiste un file campi.json, salva ed usa l'elenco integrato Federgolf
  save_campi(CAMPI_FEDERGOLF_DEFAULT)
  return CAMPI_FEDERGOLF_DEFAULT


def save_campi(campi_list):
  with open(NOME_FILE_CAMPI, "w", encoding="utf-8") as f:
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
        "Visualizza, modifica o aggiungi campi da gioco con i relativi valori"
        " di Par, CR e SR."
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
        "Modifica direttamente le celle per correggere eventuali dati e salva"
        " nell'Excel."
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
      "File Excel non trovato. Assicurati che 'Handicap_2026.xlsx' sia presente"
      " su GitHub."
  )
