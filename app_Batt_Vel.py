import streamlit as st
import pandas as pd
import plotly.express as px
from github import Github, Auth
import io
import time

# --- FUNZIONI DI UTILITY PER GITHUB ---

def get_github_client():
    """Inizializza il client GitHub usando il token nei secrets"""
    auth = Auth.Token(st.secrets["github"]["token"])
    return Github(auth=auth)

# --- NUOVA FUNZIONE DI CARICAMENTO PUBBLICA ---

@st.cache_data(ttl=600)  # Conserva i dati per 10 minuti, riducendo le chiamate a GitHub
def load_from_github():
    try:
        # URL Raw del tuo file
        url = "https://raw.githubusercontent.com/alore7/app-volley-stats/main/database.parquet"
        
        # Caricamento con timeout per evitare che l'app resti appesa
        df = pd.read_parquet(url)
        return df
    except Exception as e:
        # Se GitHub risponde ancora 429, usiamo un messaggio meno invasivo
        st.warning("⚠️ I server di GitHub sono sovraccarichi. L'app userà i dati dell'ultima sessione disponibile.")
        return None
        
def save_to_github(df):
    """Salva il DataFrame aggiornato su GitHub come file Excel"""
    try:
        g = get_github_client()
        repo = g.get_repo(st.secrets["github"]["repository"])
        path = st.secrets["github"]["file_path"]
        
        # Converte il DataFrame in un file Excel in memoria
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        data = output.getvalue()
        
        # Cerca il file esistente per aggiornarlo (serve lo SHA)
        contents = repo.get_contents(path)
        repo.update_file(path, "Aggiornamento Database Battute", data, contents.sha)
        return True
    except Exception as e:
        st.error(f"❌ Errore durante il salvataggio su GitHub: {e}")
        return False

# --- CAPITOLO 3: INIZIALIZZAZIONE MEMORIA ALL'AVVIO ---
if 'df_master' not in st.session_state:
    with st.spinner("🚀 Recupero database storico..."):
        # Chiamiamo la funzione definita poco sopra
        df_iniziale = load_from_github() 
        
        if df_iniziale is not None:
            # Salviamo i dati nella "memoria a lungo termine" dell'app
            st.session_state['df_master'] = df_iniziale
        else:
            # Se GitHub è vuoto o non risponde, evitiamo il crash
            st.session_state['df_master'] = pd.DataFrame()

def stile_zebra(x):
    # Crea un DataFrame di stili vuoto (base bianca)
    df_stili = pd.DataFrame('', index=x.index, columns=x.columns)
    # Applica il colore alle righe con indice dispari (1, 3, 5...)
    for i in range(len(x)):
        if i % 2 == 1:
            df_stili.iloc[i, :] = 'background-color: #F8F9FA'
    return df_stili

def sdoppia_percentuali(df):
    # Colonne da dividere (devono corrispondere a quelle nelle tue tabelle)
    cols_to_fix = [">120", "115-120", "110-115", "100-110", "<100", "Var.ni", "Err", "Net", "Out"]
    
    data_split = {}

    # 1. Colonne fisse iniziali
    prime_colonne = [c for c in df.columns if c not in cols_to_fix]
    for col in prime_colonne:
        data_split[(col, "")] = df[col]

    # 2. Sdoppiamento colonne con numeri e percentuali
    for col in cols_to_fix:
        if col in df.columns:
            # Aggiungiamo .iloc[:, 0] per risolvere l'errore delle dimensioni
            data_split[(col, "N°")] = df[col].str.extract(r'(\d+)').iloc[:, 0].astype(float)
            data_split[(col, "%")] = df[col].str.extract(r'\((.*)%\)').iloc[:, 0].astype(float) / 100

    # 3. Creazione MultiIndex
    df_new = pd.DataFrame(data_split)
    df_new.columns = pd.MultiIndex.from_tuples(df_new.columns)
    return df_new

def sdoppia_btxbt(df):
    if df.empty: return df
    data_split = {}
    for col in df.columns:
        # Separiamo il testo (Tipo) dal numero o codice esito (Vel.)
        data_split[(col, "Tipo")] = df[col].str.extract(r'([a-zA-Z]+)').iloc[:, 0]
        data_split[(col, "Vel/Err")] = df[col].str.extract(r'(\d+|[VNFE])').iloc[:, 0]
    
    df_new = pd.DataFrame(data_split)
    df_new.columns = pd.MultiIndex.from_tuples(df_new.columns)
    return df_new

# --- 1. CONFIGURAZIONE E COSTANTI ---
COLUMNS_A_H = ['Data', 'Partita', 'Avv.', 'Team', 'Set', 'Player', 'Tipo', 'Vel.']

# --- 2. FUNZIONI DI SUPPORTO (PULIZIA E CALCOLO) ---
def check_vel(x):
    s = str(x).strip().upper()
    if s in ["", "NAN", "NONE", "0", "0.0"]: return True 
    if s in ['N', 'F', 'V']: return True
    try:
        val = float(s.replace(',', '.'))
        return 30 <= val <= 150
    except ValueError: return False

def clean_vel_val(val):
    s = str(val).strip().upper()
    if s in ['N', 'F', 'V', 'NAN', '']: return None
    try: return float(s.replace(',', '.'))
    except: return None

def calcola_stats(df_in):
    tot = len(df_in)
    df_in = df_in.copy()
    df_in['Vel_Num'] = df_in['Vel.'].apply(clean_vel_val)
    df_spin_valide = df_in[df_in['Vel_Num'].notna()].copy()
    n_spin = len(df_spin_valide)
    media = df_spin_valide['Vel_Num'].mean() if n_spin > 0 else 0
    
    n_var = len(df_in[(df_in['Tipo'].astype(str).str.upper() == 'V') | (df_in['Vel.'].astype(str).str.upper() == 'V')])
    n_net = len(df_in[(df_in['Tipo'].astype(str).str.upper() == 'N') | (df_in['Vel.'].astype(str).str.upper() == 'N')])
    n_out = len(df_in[(df_in['Tipo'].astype(str).str.upper() == 'F') | (df_in['Vel.'].astype(str).str.upper() == 'F')])
    
    p_var = (n_var / n_spin * 100) if n_spin > 0 else 0
    var_str = f"{n_var} ({p_var:.1f}%)"
    
    n_err_tot = n_net + n_out
    p_err_tot = (n_err_tot / tot * 100) if tot > 0 else 0
    err_str = f"{n_err_tot} ({p_err_tot:.1f}%)"
    
    p_net = (n_net / n_err_tot * 100) if n_err_tot > 0 else 0
    net_str = f"{n_net} ({p_net:.1f}%)"
    p_out = (n_out / n_err_tot * 100) if n_err_tot > 0 else 0
    out_str = f"{n_out} ({p_out:.1f}%)"
    
    def fmt_f(c, t):
        p = (c / t * 100) if t > 0 else 0
        return f"{c} ({p:.1f}%)"

    f1 = fmt_f(len(df_spin_valide[df_spin_valide['Vel_Num'] >= 120]), n_spin)
    f2 = fmt_f(len(df_spin_valide[(df_spin_valide['Vel_Num'] >= 115) & (df_spin_valide['Vel_Num'] < 120)]), n_spin)
    f3 = fmt_f(len(df_spin_valide[(df_spin_valide['Vel_Num'] >= 110) & (df_spin_valide['Vel_Num'] < 115)]), n_spin)
    f4 = fmt_f(len(df_spin_valide[(df_spin_valide['Vel_Num'] >= 100) & (df_spin_valide['Vel_Num'] < 110)]), n_spin)
    f5 = fmt_f(len(df_spin_valide[df_spin_valide['Vel_Num'] < 100]), n_spin)
    
    return [tot, n_spin, media, f1, f2, f3, f4, f5, var_str, err_str, net_str, out_str]

def stile_righe(row):
    fase_val = str(row.iloc[0]) 
    if fase_val == 'MATCH': return ['background-color: #ffffcc'] * len(row)
    if 'Set 2' in fase_val or 'Set 4' in fase_val: return ['background-color: #f2f2f2'] * len(row)
    return [''] * len(row)

# --- 3. FUNZIONI GITHUB ---
def get_github_client():
    token = st.secrets["github"]["access_token"]
    auth = Auth.Token(token)
    return Github(auth=auth)

def load_master_from_github():
    try:
        g = get_github_client()
        repo = g.get_repo(st.secrets["github"]["repository"])
        contents = repo.get_contents(st.secrets["github"]["file_path"])
        df = pd.read_parquet(io.BytesIO(contents.decoded_content))
        return df.astype(str).replace('nan', '') # Forza tutto a stringa al caricamento
    except Exception: return pd.DataFrame(columns=COLUMNS_A_H)

def save_to_github(df):
    g = get_github_client()
    repo = g.get_repo(st.secrets["github"]["repository"])
    path = st.secrets["github"]["file_path"]
    
    # SOLUZIONE: Forza tutte le colonne a essere stringhe prima di salvare in Parquet
    df_save = df.astype(str).replace('nan', '')
    
    buffer = io.BytesIO()
    df_save.to_parquet(buffer, index=False)
    content = buffer.getvalue()
    try:
        contents = repo.get_contents(path)
        repo.update_file(path, "Update", content, contents.sha)
    except Exception: repo.create_file(path, "Init", content)

# --- 4. INTERFACCIA STREAMLIT ---
st.set_page_config(page_title="Sir Susa Vim Perugia - Stats", layout="wide")
st.sidebar.title("🏐 Menu Analisi")
scelta = st.sidebar.radio("Scegli:", ["Caricamento Dati", "Match", "Trend", "Velocità Team/Player", "Trend Errori", "Avversari"])

df_master = load_master_from_github()

import warnings
# Nasconde i messaggi di avviso di openpyxl e pandas
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
warnings.filterwarnings("ignore", category=FutureWarning)

if scelta == "Caricamento Dati":
    st.title("🚀 Database")
    tab1, tab2 = st.tabs(["Carica Excel", "Pulisci Database"])
    
    with tab1:
            uploaded = st.file_uploader("Seleziona file .xlsm", type=["xlsm"])
            if uploaded:
                df_new = pd.read_excel(uploaded, sheet_name="Foglio1").iloc[:, 0:8]
                df_new.columns = COLUMNS_A_H
                
                # Filtro anti-NaT
                df_new = df_new.dropna(subset=['Data', 'Player'], how='all')
                df_new = df_new[df_new['Data'].astype(str).str.upper() != 'NAT']

                if st.button("🚀 Sincronizza su GitHub"):
                    with st.spinner("Sincronizzazione in corso..."):
                        df_combined = pd.concat([df_master, df_new]).drop_duplicates()
                        # Pulizia finale
                        df_combined = df_combined.dropna(subset=['Data'], how='any')
                        df_combined = df_combined[df_combined['Data'].astype(str).str.upper() != 'NAT']
                        st.session_state['df_master'] = df_combined
                        
                        save_to_github(df_combined)
                        st.success("Dati sincronizzati!")
                        
                    st.success("Dati sincronizzati!")
                    st.balloons() # <--- I palloncini ora voleranno!
                    time.sleep(2) # <--- Ritardo magico
                    st.rerun()

    with tab2:
        st.subheader("🗑️ Gestione Partite in Database")
        if not df_master.empty:
            # 1. Crea la lista dei match presenti nel database
            df_partite = df_master[['Data', 'Avv.', 'Partita']].drop_duplicates().copy()
            # 2. Inserisce la colonna per le checkbox
            df_partite.insert(0, "Elimina", False)
            
            st.write("Seleziona i match da rimuovere e clicca sul tasto 'Conferma' in fondo:")
            
            # 3. Visualizza la tabella interattiva (Data Editor)
            edited_df = st.data_editor(
                df_partite,
                column_config={
                    "Elimina": st.column_config.CheckboxColumn(
                        "Seleziona",
                        help="Spunta per eliminare il match",
                        default=False,
                    ),
                    "Data": st.column_config.TextColumn("Data Match"),
                    "Avv.": st.column_config.TextColumn("Avversario"),
                    "Partita": st.column_config.TextColumn("Competizione")
                },
                disabled=["Data", "Avv.", "Partita"], # Blocca le celle di testo per sicurezza
                hide_index=True,
                width="stretch"
            )
            
            # 4. Recupera i match selezionati
            partite_da_eliminare = edited_df[edited_df['Elimina'] == True]
            
            if not partite_da_eliminare.empty:
                st.warning(f"⚠️ Hai selezionato {len(partite_da_eliminare)} match per l'eliminazione.")
                if st.button("🔥 Conferma ed Elimina Record Selezionati"):
                    # Filtra il database master rimuovendo i match selezionati
                    for _, row in partite_da_eliminare.iterrows():
                        df_master = df_master[~(
                            (df_master['Data'].astype(str) == str(row['Data'])) & 
                            (df_master['Avv.'].astype(str) == str(row['Avv.']))
                        )]
                    
                    save_to_github(df_master)
                    st.success("Database aggiornato con successo!")
                    st.rerun()
            else:
                st.info("Nessuna partita selezionata. Usa le caselle sopra per procedere.")
        else:
            st.info("Il database è attualmente vuoto.")

elif scelta == "Match":
    st.title("🏟️ Report Match")
    if not df_master.empty:
        competizioni = sorted([str(x) for x in df_master['Partita'].dropna().unique() if str(x) != ''])
        c_sel = st.selectbox("Competizione:", competizioni)
        
        df_c = df_master[df_master['Partita'].astype(str) == c_sel].copy()
        df_c['Match_Label'] = df_c['Data'].astype(str) + " vs " + df_c['Avv.'].astype(str)
        
        match_list = sorted([str(x) for x in df_c['Match_Label'].dropna().unique() if str(x) != ''], reverse=True)
        m_sel = st.multiselect("Seleziona Match:", match_list)
        
        if m_sel:
            m_rep = st.radio("Vista:", ["REPORT", "GRAFICI"], horizontal=True)
            df_report = df_c[df_c['Match_Label'].isin(m_sel)].copy()
            df_report['Vel_Num'] = df_report['Vel.'].apply(clean_vel_val)
            info = df_report.iloc[0]
            nome_avv = str(info['Avv.']).upper()
            cols_h = ["Fase", "Tot", "Spin", "Media Km/h", ">120", "115-120", "110-115", "100-110", "<100", "Var.ni", "Err", "Net", "Out"]

            if m_rep == "REPORT":
                st.markdown("<h2 style='text-align: center;'>📋 REPORT VELOCITÀ BATTUTA SPIN</h2>", unsafe_allow_html=True)
                col1, col2, col3 = st.columns(3)
                col1.success(f"**Manifestazione**\n\n{info['Partita']}")
                col2.success(f"**Data**\n\n{info['Data']}")
                col3.success(f"**Avversario**\n\n{nome_avv}")
                
                st.markdown("### 🏐 PERUGIA")
                df_p = df_report[df_report['Team'].astype(str).str.upper() == 'PERUGIA'].copy()
                r_p = [["MATCH"] + calcola_stats(df_p)]
                for s in sorted(df_p['Set'].unique()): r_p.append([f"Set {int(float(s))}"] + calcola_stats(df_p[df_p['Set'] == s]))
                st.dataframe(pd.DataFrame(r_p, columns=cols_h).style.hide(axis="index").apply(stile_righe, axis=1).format({"Media Km/h": "{:.1f}"}, precision=1))

                st.markdown(f"### 🏐 {nome_avv}")
                df_o = df_report[df_report['Team'].astype(str).str.upper() != 'PERUGIA'].copy()
                if not df_o.empty:
                    r_o = [["MATCH"] + calcola_stats(df_o)]
                    for s in sorted(df_o['Set'].unique()): r_o.append([f"Set {int(float(s))}"] + calcola_stats(df_o[df_o['Set'] == s]))
                    st.dataframe(pd.DataFrame(r_o, columns=cols_h).style.hide(axis="index").apply(stile_righe, axis=1).format({"Media Km/h": "{:.1f}"}, precision=1))

                # --- PRIMO BOTTONE AGGIORNATO ---
                buffer_squadre = io.BytesIO()
                with pd.ExcelWriter(buffer_squadre, engine='xlsxwriter') as writer:
                    # Applichiamo lo sdoppiamento prima di salvare
                    df_p_multi = sdoppia_percentuali(pd.DataFrame(r_p, columns=cols_h))
                    df_o_multi = sdoppia_percentuali(pd.DataFrame(r_o, columns=cols_h))
                    
                    df_p_multi.to_excel(writer, sheet_name='Perugia')
                    df_o_multi.to_excel(writer, sheet_name=nome_avv[:30])

                st.download_button(
                    label="📥 Scarica Report Squadre",
                    data=buffer_squadre.getvalue(),
                    file_name="Report_Generale_Squadre.xlsx",
                    key="btn_squadre_multi"
                )

                st.divider()

                # --- SEZIONE PLAYER PERUGIA ---
                st.markdown("## 🏐 PERUGIA - PLAYER")

                # Estrazione lista giocatori Perugia
                p_list = sorted([str(x) for x in df_p[df_p['Vel_Num'].notna() | df_p['Tipo'].isin(['V','N','F','v','n','f'])]['Player'].unique()])

                # Layout Selettori Perugia
                col_fase_p, col_gioc_p = st.columns(2)
                with col_fase_p:
                    fs_p = st.selectbox("Fase Perugia (Tabella Generale):", ["MATCH"] + [f"Set {int(float(s))}" for s in sorted(df_p['Set'].unique())], key="fs_p_new")
                with col_gioc_p:
                    ps_p = st.selectbox("Seleziona Giocatore (Tabella Individuale):", p_list, key="ps_p_new")

                # 1. Tabella Generale Perugia
                rg_p = []
                for p in p_list:
                    df_pf = df_p[df_p['Player'] == p] if fs_p == "MATCH" else df_p[(df_p['Player'] == p) & (df_p['Set'].astype(float) == float(fs_p.replace("Set ","")))]
                    if not df_pf.empty: 
                        rg_p.append([p] + calcola_stats(df_pf))

                st.dataframe(
                    pd.DataFrame(rg_p, columns=["Player"] + cols_h[1:])
                    .style.apply(stile_zebra, axis=None)
                    .hide(axis="index")
                    .format({"Media Km/h": "{:.1f}"}, precision=1), 
                    use_container_width=True
                )

                # 2. Tabella Individuale Perugia
                df_i_p = df_p[df_p['Player'] == ps_p]
                ri_p = [["MATCH"] + calcola_stats(df_i_p)]
                for s in sorted(df_i_p['Set'].unique()): 
                    ri_p.append([f"Set {int(float(s))}"] + calcola_stats(df_i_p[df_i_p['Set'] == s]))

                st.dataframe(
                    pd.DataFrame(ri_p, columns=cols_h)
                    .style.apply(stile_zebra, axis=None)
                    .apply(stile_righe, axis=1)
                    .hide(axis="index")
                    .format({"Media Km/h": "{:.1f}"}, precision=1), 
                    use_container_width=True
                )

                st.divider()

                # --- SEZIONE PLAYER AVVERSARIO ---
                st.markdown(f"## 🏐 {nome_avv} - PLAYER")

                # Estrazione lista giocatori Avversario
                a_list = sorted([str(x) for x in df_o[df_o['Vel_Num'].notna() | df_o['Tipo'].isin(['V','N','F','v','n','f'])]['Player'].unique()])

                # Layout Selettori Avversario
                col_fase_a, col_gioc_a = st.columns(2)
                with col_fase_a:
                    fs_a = st.selectbox(f"Fase {nome_avv} (Tabella Generale):", ["MATCH"] + [f"Set {int(float(s))}" for s in sorted(df_o['Set'].unique())], key="fs_a_new")
                with col_gioc_a:
                    ps_a = st.selectbox(f"Seleziona Giocatore {nome_avv} (Tabella Individuale):", a_list, key="ps_a_new")

                # 1. Tabella Generale Avversario
                rg_a = []
                for a in a_list:
                    df_af = df_o[df_o['Player'] == a] if fs_a == "MATCH" else df_o[(df_o['Player'] == a) & (df_o['Set'].astype(float) == float(fs_a.replace("Set ","")))]
                    if not df_af.empty: 
                        rg_a.append([a] + calcola_stats(df_af))

                st.dataframe(
                    pd.DataFrame(rg_a, columns=["Player"] + cols_h[1:])
                    .style.apply(stile_zebra, axis=None)
                    .hide(axis="index")
                    .format({"Media Km/h": "{:.1f}"}, precision=1), 
                    use_container_width=True
                )

                # 2. Tabella Individuale Avversario
                df_i_a = df_o[df_o['Player'] == ps_a]
                ri_a = [["MATCH"] + calcola_stats(df_i_a)]
                for s in sorted(df_i_a['Set'].unique()): 
                    ri_a.append([f"Set {int(float(s))}"] + calcola_stats(df_i_a[df_i_a['Set'] == s]))

                st.dataframe(
                    pd.DataFrame(ri_a, columns=cols_h)
                    .style.apply(stile_zebra, axis=None)
                    .apply(stile_righe, axis=1)
                    .hide(axis="index")
                    .format({"Media Km/h": "{:.1f}"}, precision=1), 
                    use_container_width=True
                )

                # --- SECONDO BOTTONE: REPORT PLAYER (CON COLONNE DOPPIE) ---
                # Preparazione dati per il download - Perugia
                df_p_raw = pd.DataFrame(rg_p, columns=["Player"] + cols_h[1:])

                # Preparazione dati per il download - Avversario (usiamo 'o' per evitare il NameError)
                df_o_raw = pd.DataFrame(rg_a, columns=["Player"] + cols_h[1:])

                buffer_player = io.BytesIO()
                with pd.ExcelWriter(buffer_player, engine='xlsxwriter') as writer:
                    # Applichiamo lo sdoppiamento (se la struttura delle colonne è compatibile con cols_h)
                    df_p_player_multi = sdoppia_percentuali(df_p_raw)
                    df_o_player_multi = sdoppia_percentuali(df_o_raw)
                    
                    df_p_player_multi.to_excel(writer, sheet_name='Perugia_Player')
                    df_o_player_multi.to_excel(writer, sheet_name=f'{nome_avv}_Player'[:30])

                st.download_button(
                    label="📥 Scarica Report Player",
                    data=buffer_player.getvalue(),
                    file_name="Report_Player_Match.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_player_multi"
                )

            # --- SEZIONE FINALE: BtxBT PLAYER (Sequenza Cronologica) ---
            st.divider()
            st.subheader("📊 Sequenza Cronologica Battute (BtxBT)")

            if not df_report.empty:
                # 1. Filtro Set
                set_disponibili = sorted(df_report['Set'].unique())
                set_scelto = st.selectbox("Seleziona il Set:", set_disponibili, format_func=lambda x: f"Set {int(float(x))}")
                df_set = df_report[df_report['Set'] == set_scelto].copy()

                # 2. Logica di separazione Team
                nomi_team = df_report['Team'].unique()
                
                # Definiamo Perugia
                team_sir = next((t for t in nomi_team if any(key in str(t).upper() for key in ['SIR', 'PERUGIA'])), None)
                
                # Definiamo l'Avversario
                team_avv = next((t for t in nomi_team if t != team_sir), None)

                def genera_tabella_per_team(df_input, nome_team_target):
                    if not nome_team_target: return pd.DataFrame()
                    df_team = df_input[df_input['Team'] == nome_team_target]
                    giocatori = sorted(df_team['Player'].unique())
                    if not giocatori: return pd.DataFrame()

                    data_rows = []
                    max_battute = df_team.groupby('Player').size().max()
                    for i in range(max_battute):
                        row = {}
                        for p in giocatori:
                            p_battute = df_team[df_team['Player'] == p]
                            if i < len(p_battute):
                                b = p_battute.iloc[i]
                                row[p] = f"{b['Tipo']} {b['Vel.']}"
                            else:
                                row[p] = ""
                        data_rows.append(row)
                    return pd.DataFrame(data_rows)

                # Funzione per il colore giallo chiarissimo
                def evidenzia_errori(val):
                    v = str(val).upper()
                    if ' N' in v or ' F' in v:
                        return 'background-color: #ffffcc; font-weight: bold;'
                    return ''

                # 3. Visualizzazione Tabella PERUGIA
                st.markdown("#### 🏐 PERUGIA")
                df_p_bt = pd.DataFrame() # Inizializziamo per sicurezza
                if team_sir:
                    df_p_bt = genera_tabella_per_team(df_set, team_sir)
                    if not df_p_bt.empty:
                        st.dataframe(df_p_bt.style.applymap(evidenzia_errori), use_container_width=True, hide_index=True)
                    else:
                        st.info("Nessuna battuta per Perugia in questo set.")

                # 4. Visualizzazione Tabella AVVERSARI
                st.markdown(f"#### 🏐 {nome_avv}")
                df_a_bt = pd.DataFrame() # Inizializziamo per sicurezza
                if team_avv:
                    df_a_bt = genera_tabella_per_team(df_set, team_avv)
                    if not df_a_bt.empty:
                        st.dataframe(df_a_bt.style.applymap(evidenzia_errori), use_container_width=True, hide_index=True)
                    else:
                        st.info(f"Nessuna battuta per l'avversario in questo set.")

                # --- TERZO BOTTONE AGGIORNATO (BtxBT) ---
                if not df_p_bt.empty or not df_a_bt.empty:
                    buffer_bt = io.BytesIO()
                    with pd.ExcelWriter(buffer_bt, engine='xlsxwriter') as writer:
                        # Applichiamo lo sdoppiamento specifico per BtxBT
                        df_p_bt_multi = sdoppia_btxbt(df_p_bt)
                        df_a_bt_multi = sdoppia_btxbt(df_a_bt)
                        
                        df_p_bt_multi.to_excel(writer, sheet_name='Perugia')
                        df_a_bt_multi.to_excel(writer, sheet_name='Avversario')
                    
                    st.download_button(
                        label=f"📥 Scarica BtxBT Set {int(float(set_scelto))}",
                        data=buffer_bt.getvalue(),
                        file_name=f"BtxBT_Set_{set_scelto}.xlsx",
                        key=f"btn_btxbt_multi_{set_scelto}"
                    )

            elif m_rep == "GRAFICI":
                st.subheader("📊 Analisi Visiva della Battuta")
                df_graf = df_report[df_report['Vel_Num'].notna()].copy()
                if not df_graf.empty:
                    st.markdown("##### 🏎️ Distribuzione Potenza")
                    fig_box = px.box(df_graf, x="Team", y="Vel_Num", color="Team", points="all",
                                     color_discrete_map={'PERUGIA': '#C41E3A', info['Avv.']: '#0047AB'})
                    st.plotly_chart(fig_box, width='stretch')

                    st.markdown("##### 📈 Trend Velocità per Set")
                    df_trend = df_graf.groupby(['Set', 'Team'])['Vel_Num'].mean().reset_index()
                    fig_line = px.line(df_trend, x="Set", y="Vel_Num", color="Team", markers=True)
                    st.plotly_chart(fig_line, width='stretch')

                    st.markdown("##### 🏆 Top 8 Performance")
                    df_top = df_graf.groupby(['Player', 'Team'])['Vel_Num'].mean().reset_index()
                    df_top = df_top.sort_values(by='Vel_Num', ascending=False).head(8)
                    fig_bar = px.bar(df_top, x='Vel_Num', y='Player', color='Team', orientation='h', text_auto='.1f')
                    st.plotly_chart(fig_bar, width='stretch')
                else:
                    st.warning("Dati di velocità non disponibili.")

# ... qui finisce tutto il blocco indentato di "Match" ...

elif scelta == "Trend":
    st.markdown("<h2 style='text-align: center;'>📈 TREND STORICO PARTITE</h2>", unsafe_allow_html=True)
    
    # Verifichiamo se il database master è carico
    if 'df_master' in st.session_state and not st.session_state['df_master'].empty:
        df_trend_base = st.session_state['df_master'].copy()
        
        # --- AGGIUNGI QUESTE RIGHE PER RISOLVERE IL FILTRO ---
        # Forza la colonna Data ad essere un formato data pulito (senza ore/minuti)
        df_trend_base['Data'] = pd.to_datetime(df_trend_base['Data'], format='mixed').dt.date
        # ----------------------------------------------------

        # Ora crea la lista dei match per il multiselect
        opzioni_match = sorted(df_trend_base['Data'].unique(), reverse=True)
        
        selected_matches = st.multiselect(
            "Scegli le partite:", 
            options=opzioni_match,
            default=opzioni_match # Inizia mostrandole tutte
        )

        # Filtra il database in base alla selezione
        df_filtrato = df_trend_base[df_trend_base['Data'].isin(selected_matches)]

        if not df_filtrato.empty:
            # QUI METTI IL CODICE DEI TUOI GRAFICI (px.line, px.bar, ecc.)
            st.success(f"Analisi di {len(df_filtrato)} record")
        else:
            st.warning("Nessun dato disponibile per i match selezionati.")
        
        # --- 1. SELEZIONE PARTITE ---
        # Creiamo una lista di opzioni combinando Data e Avversario per chiarezza
        df_trend_base['Match_Label'] = df_trend_base['Data'].astype(str) + " - vs " + df_trend_base['Avv.']
        lista_match = sorted(df_trend_base['Match_Label'].unique())
                
        # Filtriamo il dataframe usando il nome del PRIMO selettore (quello dei tag rossi)
        if not selected_matches:
            df_trend_filtered = df_trend_base.copy()
        else:
            # Qui usiamo 'Data' perché è la colonna che hai pulito sopra
            df_trend_filtered = df_trend_base[df_trend_base['Data'].isin(selected_matches)].copy()

        # --- 2. ELABORAZIONE DATI PERUGIA ---
        # Filtriamo solo i dati di Perugia dai match selezionati
        df_p_trend = df_trend_filtered[df_trend_filtered['Team'] == 'PERUGIA'].copy()
        
        trend_data = []
        # Cicliamo sulle etichette dei match per mantenere l'ordine cronologico
        for m_label in lista_match:
            df_m = df_p_trend[df_p_trend['Match_Label'] == m_label]
            
            if not df_m.empty:
                # Convertiamo Vel in numerico per il calcolo
                df_m['Vel_Num'] = pd.to_numeric(df_m['Vel.'], errors='coerce')
                media_vel = df_m['Vel_Num'].mean()
                
                # Calcolo % Errori
                battute_tot = len(df_m)
                err_tot = len(df_m[df_m['Tipo'].str.upper().isin(['E','N','O'])])
                perc_err = (err_tot / battute_tot * 100) if battute_tot > 0 else 0
                
                trend_data.append({
                    "Gara": m_label,
                    "Media Km/h": round(media_vel, 1),
                    "% Errori": round(perc_err, 1)
                })
        
        df_plot = pd.DataFrame(trend_data)

        # --- 3. GRAFICI ---
        if not df_plot.empty:
            # Grafico Potenza
            st.subheader("🚀 Evoluzione Potenza Media")
            fig_vel = px.line(df_plot, x="Gara", y="Media Km/h", markers=True, text="Media Km/h",
                              color_discrete_sequence=["#C00000"])
            fig_vel.update_traces(textposition="top center")
            st.plotly_chart(fig_vel, use_container_width=True)

            # Grafico Fallosità
            st.subheader("📉 Analisi della Precisione (% Errori)")
            fig_err = px.bar(df_plot, x="Gara", y="% Errori", text="% Errori",
                             color_discrete_sequence=["#FFA500"])
            fig_err.update_traces(texttemplate='%{text}%', textposition='outside')
            st.plotly_chart(fig_err, use_container_width=True)
            
            # Tabella di riepilogo
            st.write("### 📋 Dati Comparativi")
            st.dataframe(df_plot, use_container_width=True, hide_index=True)
        if not df_filtrato.empty:
            st.success(f"Analisi di {len(df_filtrato)} record") # Quello che vedi ora

            # --- AGGIUNGIAMO UN GRAFICO DI ESEMPIO ---
            st.subheader("📈 Andamento Velocità Media")
            
            # Raggruppiamo i dati per data per calcolare la media
        # --- INIZIO CORREZIONE SICURA ---
        if not df_filtrato.empty:
            # Identifichiamo automaticamente la colonna della velocità
            # Cerchiamo tra i nomi possibili: 'Vel', 'Vel.', 'Velocità'
            colonne_possibili = ['Vel', 'Vel.', 'Velocità']
            colonna_vel = next((c for c in colonne_possibili if c in df_filtrato.columns), None)

            if colonna_vel:
                st.subheader("📈 Andamento Velocità Media")
                
                # Convertiamo in numero la colonna trovata
                df_filtrato['Vel_Num'] = pd.to_numeric(df_filtrato[colonna_vel], errors='coerce')
                
                # Calcoliamo la media raggruppata per Data
                df_media = df_filtrato.groupby('Data')['Vel_Num'].mean().reset_index()
                
                # Creiamo il grafico usando il nome corretto della colonna calcolata
                fig = px.line(
                    df_media, 
                    x='Data', 
                    y='Vel_Num',  # <--- Deve essere Vel_Num, non Velocità
                    title="Evoluzione Velocità Media nel Tempo",
                    labels={'Vel_Num': 'Velocità Media (km/h)', 'Data': 'Partita'},
                    markers=True
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.error(f"❌ Colonna velocità non trovata. Colonne disponibili: {list(df_filtrato.columns)}")
        # --- FINE CORREZIONE ---
            
            fig = px.line(
                df_media, 
                x='Data', 
                y='Velocità', 
                title="Evoluzione Velocità Media nel Tempo",
                markers=True
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # --- FINE GRAFICO ---

        else:
            st.warning("Seleziona almeno una partita per vedere i grafici.")
                    


