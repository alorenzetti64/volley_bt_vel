import streamlit as st
import pandas as pd
import plotly.express as px
from github import Github, Auth
import io
import time
import textwrap
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

def _safe_pdf_text(value):
    return str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def build_trend_pdf(selected_matches, df_table_pdf, df_plot_pdf, modalita, etichetta_metriche):
    buffer = io.BytesIO()

    def _new_landscape_fig():
        return plt.figure(figsize=(11.69, 8.27))

    with PdfPages(buffer) as pdf:
        # Pagina 1: partite selezionate + tabella
        fig = _new_landscape_fig()
        ax = fig.add_axes([0, 0, 1, 1])
        ax.axis('off')

        fig.text(0.03, 0.95, 'Trend Report Battuta', fontsize=18, fontweight='bold', ha='left', va='top')
        fig.text(0.03, 0.915, etichetta_metriche, fontsize=10, ha='left', va='top')
        fig.text(0.03, 0.89, f'Modalità: {modalita}', fontsize=10, ha='left', va='top')

        def _match_label_pdf(m):
            s = str(m)
            if ' - vs ' in s:
                s = s.split(' - vs ', 1)[1]
            return s.upper()

        match_items = [f"{i}. {_match_label_pdf(m)}" for i, m in enumerate(selected_matches, 1)]
        if not match_items:
            match_items = ['Nessuna partita selezionata']

        # Vera disposizione orizzontale in 5 colonne tramite tabella
        n_cols_matches = 5
        rows_needed = max(1, (len(match_items) + n_cols_matches - 1) // n_cols_matches)
        match_rows = []
        for r in range(rows_needed):
            row = []
            for c in range(n_cols_matches):
                idx = r * n_cols_matches + c
                row.append(match_items[idx] if idx < len(match_items) else '')
            match_rows.append(row)

        matches_height = min(0.26, 0.04 + rows_needed * 0.028)
        matches_bottom = 0.84 - matches_height
        matches_ax = fig.add_axes([0.03, matches_bottom, 0.94, matches_height])
        matches_ax.axis('off')
        matches_tbl = matches_ax.table(
            cellText=match_rows,
            cellLoc='left',
            colLoc='left',
            loc='upper left',
            bbox=[0, 0, 1, 1]
        )
        matches_tbl.auto_set_font_size(False)
        matches_tbl.set_fontsize(9)
        matches_tbl.scale(1, 1.2)
        for (_, _), cell in matches_tbl.get_celld().items():
            cell.set_linewidth(0)
            cell.set_edgecolor('white')
            cell.set_facecolor('white')
            cell.PAD = 0.02
            cell.get_text().set_fontfamily('monospace')
            cell.get_text().set_ha('left')
            cell.get_text().set_va('center')

        # Nessun titolo intermedio: la prima pagina deve restare pulita
        table_df = df_table_pdf.copy()
        rename_map = {
            'Giocatore': 'Gioc.',
            'Data': 'Data',
            'Avversario': 'Avv.',
            'Battute Totali': 'Spin',
            'Errori Totali': 'Err',
            'Errori N': 'N',
            'Errori F': 'F',
            'Media Km/h': 'Km/h',
            '% Errori': '% Err',
            '% Errori N': '% N',
            '% Errori F': '% F',
        }
        table_df = table_df.rename(columns=rename_map)

        for col in table_df.columns:
            if col in ['Km/h', '% Err', '% N', '% F']:
                table_df[col] = table_df[col].map(lambda x: f"{float(x):.1f}" if pd.notna(x) and str(x) != '' else '')
            else:
                table_df[col] = table_df[col].fillna('').astype(str)

        max_rows = 18
        table_show = table_df.head(max_rows)
        if len(table_df) > max_rows:
            extra = pd.DataFrame([{c: '...' for c in table_show.columns}])
            table_show = pd.concat([table_show, extra], ignore_index=True)

        table_top = matches_bottom - 0.03
        table_bottom = 0.07
        table_height = max(0.36, table_top - table_bottom)
        ax_tbl = fig.add_axes([0.03, table_bottom, 0.94, table_height])
        ax_tbl.axis('off')
        tbl = ax_tbl.table(
            cellText=table_show.values,
            colLabels=table_show.columns,
            loc='upper left',
            cellLoc='center',
            bbox=[0, 0, 1, 1]
        )
        tbl.auto_set_font_size(False)
        tbl.set_fontsize(8)
        tbl.scale(1, 1.25)
        for (row, col), cell in tbl.get_celld().items():
            cell.set_linewidth(0.4)
            if row == 0:
                cell.set_text_props(weight='bold')
                cell.set_facecolor('#D9E2F3')
            elif row % 2 == 0:
                cell.set_facecolor('#F7F7F7')

        pdf.savefig(fig, bbox_inches='tight')
        plt.close(fig)

        x = list(range(len(df_plot_pdf)))
        labels = df_plot_pdf['Avv_Breve'].tolist()

        # Pagina 2 - velocità
        fig, ax = plt.subplots(figsize=(11.69, 8.27))
        ax.plot(x, df_plot_pdf['Media Km/h'], marker='o', linewidth=2)
        ax.plot(x, df_plot_pdf['Trend_Vel'], linestyle='--', linewidth=1.8)
        for xi, yi in zip(x, df_plot_pdf['Media Km/h']):
            ax.text(xi, yi + 0.35, f"{yi:.1f}", ha='center', va='bottom', fontsize=8)
        ax.set_xticks(x)
        ax.set_xticklabels(labels)
        ax.set_title('Andamento Velocità Media nel Tempo')
        ax.set_ylabel('Velocità Media (km/h)')
        ax.grid(axis='y', alpha=0.3)
        fig.tight_layout()
        pdf.savefig(fig, bbox_inches='tight')
        plt.close(fig)

        # Pagina 3 - errori
        fig, ax = plt.subplots(figsize=(11.69, 8.27))
        bars = ax.bar(x, df_plot_pdf['% Errori'])
        ax.plot(x, df_plot_pdf['Trend_Err'], linestyle='--', linewidth=1.8)
        for rect, val in zip(bars, df_plot_pdf['% Errori']):
            ax.text(rect.get_x() + rect.get_width()/2, rect.get_height() + 0.4, f"{val:.1f}%", ha='center', va='bottom', fontsize=8)
        ax.set_xticks(x)
        ax.set_xticklabels(labels)
        ax.set_title('Andamento ERRORE')
        ax.set_ylabel('% Errori')
        ax.grid(axis='y', alpha=0.3)
        fig.tight_layout()
        pdf.savefig(fig, bbox_inches='tight')
        plt.close(fig)

        # Pagina 4 - tipologie errore
        fig, ax = plt.subplots(figsize=(11.69, 8.27))
        width = 0.38
        x_n = [i - width/2 for i in x]
        x_f = [i + width/2 for i in x]
        bars_n = ax.bar(x_n, df_plot_pdf['% Errori N'], width=width, label='N')
        bars_f = ax.bar(x_f, df_plot_pdf['% Errori F'], width=width, label='F')
        ax.plot(x, df_plot_pdf['Trend_Err_N'], linestyle='--', linewidth=1.6, label='Trend N')
        ax.plot(x, df_plot_pdf['Trend_Err_F'], linestyle=':', linewidth=1.8, label='Trend F')
        for rect, val in zip(bars_n, df_plot_pdf['% Errori N']):
            if val > 0:
                ax.text(rect.get_x() + rect.get_width()/2, rect.get_height() + 0.25, f"{val:.1f}%", ha='center', va='bottom', fontsize=7)
        for rect, val in zip(bars_f, df_plot_pdf['% Errori F']):
            if val > 0:
                ax.text(rect.get_x() + rect.get_width()/2, rect.get_height() + 0.25, f"{val:.1f}%", ha='center', va='bottom', fontsize=7)
        ax.set_xticks(x)
        ax.set_xticklabels(labels)
        ax.set_title('Tipologie di ERRORE')
        ax.set_ylabel('% Errori per tipologia')
        ax.grid(axis='y', alpha=0.3)
        ax.legend(loc='upper left', ncol=4, fontsize=8)
        fig.tight_layout()
        pdf.savefig(fig, bbox_inches='tight')
        plt.close(fig)

    buffer.seek(0)
    return buffer.getvalue()


# --- FUNZIONI DI UTILITY PER GITHUB ---

def get_github_client():
    """Inizializza il client GitHub usando il token nei secrets"""
    auth = Auth.Token(st.secrets["github"]["token"])
    return Github(auth=auth)

# --- NUOVA FUNZIONE DI CARICAMENTO PUBBLICA ---

def load_from_github():
    """Versione standard che usavi prima - Richiede il Token nei Secrets"""
    try:
        from github import Auth, Github
        import io
        
        # Questa riga cerca il token che DEVE essere nei Secrets di Streamlit Cloud
        auth = Auth.Token(st.secrets["github"]["token"])
        g = Github(auth=auth)
        
        repo = g.get_repo(st.secrets["github"]["repository"])
        path = st.secrets["github"]["file_path"]
        
        file_content = repo.get_contents(path)
        df = pd.read_parquet(io.BytesIO(file_content.decoded_content))
        return df
    except Exception as e:
        st.error(f"⚠️ Errore: {e}")
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
    df = df_in.copy()

    st.write("DEBUG ROWS IN INGRESSO")
    st.write(df[['Team', 'Set', 'Player', 'Tipo', 'Vel.']].to_string())

    df['TipoU'] = df['Tipo'].astype(str).str.upper().str.strip()
    df['VelS'] = df['Vel.'].astype(str).str.upper().str.strip()

    st.write("DEBUG SOLO PERUGIA SPIN")
    st.write(df[(df['Team'].astype(str).str.upper().str.strip() == 'PERUGIA') & (df['TipoU'] == 'SPIN')][['Set','Player','Tipo','Vel.','VelS']].to_string())

    st.write("N righe totali ricevute:", len(df))
    st.write("N righe SPIN:", len(df[df['TipoU'] == 'SPIN']))
    st.write("N righe SPIN numeriche:", pd.to_numeric(df[df['TipoU'] == 'SPIN']['VelS'], errors='coerce').notna().sum())

    # Tengo solo le battute SPIN
    df_spin = df[df['TipoU'] == 'SPIN'].copy()

    # TOT = tutte le SPIN
    tot = len(df_spin)

    # Valide = solo quelle numeriche
    df_spin['Vel_Num'] = pd.to_numeric(df_spin['VelS'], errors='coerce')
    df_valide = df_spin[df_spin['Vel_Num'].notna()].copy()
    n_spin = len(df_valide)

    media = df_valide['Vel_Num'].mean() if n_spin > 0 else 0

    # Codici speciali sulle sole SPIN
    n_var = (df_spin['VelS'].str.startswith('V')).sum()
    n_net = (df_spin['VelS'].str.startswith('N')).sum()
    n_out = (df_spin['VelS'].str.startswith('F')).sum()
    n_err_tot = n_net + n_out

    def fmt_f(c, t):
        p = (c / t * 100) if t > 0 else 0
        return f"{c} ({p:.1f}%)"

    f1 = fmt_f(len(df_valide[df_valide['Vel_Num'] > 120]), n_spin)
    f2 = fmt_f(len(df_valide[(df_valide['Vel_Num'] >= 115) & (df_valide['Vel_Num'] <= 120)]), n_spin)
    f3 = fmt_f(len(df_valide[(df_valide['Vel_Num'] >= 110) & (df_valide['Vel_Num'] < 115)]), n_spin)
    f4 = fmt_f(len(df_valide[(df_valide['Vel_Num'] >= 100) & (df_valide['Vel_Num'] < 110)]), n_spin)
    f5 = fmt_f(len(df_valide[df_valide['Vel_Num'] < 100]), n_spin)

    var_str = fmt_f(n_var, tot)
    err_str = fmt_f(n_err_tot, tot)
    net_str = fmt_f(n_net, n_err_tot) if n_err_tot > 0 else '0 (0.0%)'
    out_str = fmt_f(n_out, n_err_tot) if n_err_tot > 0 else '0 (0.0%)'

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
st.warning("DEBUG VERSIONE NUOVA - TEST 11 MARZO")
st.sidebar.title("🏐 Menu Analisi")
scelta = st.sidebar.radio("Scegli:", ["Caricamento Dati", "Match", "Trend Team/Player", "Storico Avversari"])

# Usa prima i dati caricati nella sessione; solo in assenza totale ripiega su GitHub.
if 'df_master' in st.session_state and isinstance(st.session_state['df_master'], pd.DataFrame):
    df_master = st.session_state['df_master'].copy()
else:
    df_master = load_master_from_github()
    st.session_state['df_master'] = df_master.copy()

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
                    
                    st.session_state['df_master'] = df_master.copy()
                    save_to_github(df_master)
                    st.success("Database aggiornato con successo!")
                    st.rerun()
            else:
                st.info("Nessuna partita selezionata. Usa le caselle sopra per procedere.")
        else:
            st.info("Il database è attualmente vuoto.")

elif scelta == "Match":
    st.title("🏟️ Report Match")

    if df_master.empty:
        st.info("Carica prima i dati per visualizzare il Report Match.")
    else:
        competizioni = sorted([str(x) for x in df_master['Partita'].dropna().unique() if str(x) != ''])

        if not competizioni:
            st.info("Nessuna competizione disponibile nel database.")
        else:
            c_sel = st.selectbox("Competizione:", competizioni)

            df_c = df_master[df_master['Partita'].astype(str) == c_sel].copy()
            df_c['Match_Label'] = df_c['Data'].astype(str) + " vs " + df_c['Avv.'].astype(str)

            match_list = sorted([str(x) for x in df_c['Match_Label'].dropna().unique() if str(x) != ''], reverse=True)
            m_sel = st.multiselect("Seleziona Match:", match_list)

            if not m_sel:
                st.info("Seleziona almeno una partita per aprire il Report Match.")
            else:
                m_rep = st.radio("Vista:", ["REPORT", "GRAFICI"], horizontal=True)
                df_report = df_c[df_c['Match_Label'].isin(m_sel)].copy()

                if df_report.empty:
                    st.warning("Nessun dato disponibile per le partite selezionate.")
                else:
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
                        df_p = df_report[df_report['Team'].astype(str).str.upper().str.strip() == 'PERUGIA'].copy()
                        r_p = [["MATCH"] + calcola_stats(df_p)]
                        for s in sorted(df_p['Set'].dropna().unique()):
                            r_p.append([f"Set {int(float(s))}"] + calcola_stats(df_p[df_p['Set'] == s]))
                        st.dataframe(pd.DataFrame(r_p, columns=cols_h).style.hide(axis="index").apply(stile_righe, axis=1).format({"Media Km/h": "{:.1f}"}, precision=1))

                        st.markdown(f"### 🏐 {nome_avv}")
                        df_o = df_report[df_report['Team'].astype(str).str.upper().str.strip() != 'PERUGIA'].copy()
                        r_o = []
                        if not df_o.empty:
                            r_o = [["MATCH"] + calcola_stats(df_o)]
                            for s in sorted(df_o['Set'].dropna().unique()):
                                r_o.append([f"Set {int(float(s))}"] + calcola_stats(df_o[df_o['Set'] == s]))
                            st.dataframe(pd.DataFrame(r_o, columns=cols_h).style.hide(axis="index").apply(stile_righe, axis=1).format({"Media Km/h": "{:.1f}"}, precision=1))
                        else:
                            st.info("Nessun dato disponibile per l'avversario.")

                        buffer_squadre = io.BytesIO()
                        with pd.ExcelWriter(buffer_squadre, engine='xlsxwriter') as writer:
                            df_p_multi = sdoppia_percentuali(pd.DataFrame(r_p, columns=cols_h))
                            df_p_multi.to_excel(writer, sheet_name='Perugia')
                            if r_o:
                                df_o_multi = sdoppia_percentuali(pd.DataFrame(r_o, columns=cols_h))
                                df_o_multi.to_excel(writer, sheet_name=nome_avv[:30])

                        st.download_button(
                            label="📥 Scarica Report Squadre",
                            data=buffer_squadre.getvalue(),
                            file_name="Report_Generale_Squadre.xlsx",
                            key="btn_squadre_multi"
                        )

                        st.divider()
                        st.markdown("## 🏐 PERUGIA - PLAYER")
                        p_list = sorted([str(x) for x in df_p[df_p['Vel_Num'].notna() | df_p['Tipo'].isin(['V','N','F','v','n','f'])]['Player'].dropna().unique()])

                        if p_list:
                            col_fase_p, col_gioc_p = st.columns(2)
                            with col_fase_p:
                                fs_p = st.selectbox("Fase Perugia (Tabella Generale):", ["MATCH"] + [f"Set {int(float(s))}" for s in sorted(df_p['Set'].dropna().unique())], key="fs_p_new")
                            with col_gioc_p:
                                ps_p = st.selectbox("Seleziona Giocatore (Tabella Individuale):", p_list, key="ps_p_new")

                            rg_p = []
                            for p_name in p_list:
                                df_pf = df_p[df_p['Player'] == p_name] if fs_p == "MATCH" else df_p[(df_p['Player'] == p_name) & (df_p['Set'].astype(float) == float(fs_p.replace("Set ","")))]
                                if not df_pf.empty:
                                    rg_p.append([p_name] + calcola_stats(df_pf))

                            st.dataframe(
                                pd.DataFrame(rg_p, columns=["Player"] + cols_h[1:]).style.apply(stile_zebra, axis=None).hide(axis="index").format({"Media Km/h": "{:.1f}"}, precision=1),
                                use_container_width=True
                            )

                            df_i_p = df_p[df_p['Player'] == ps_p]
                            ri_p = [["MATCH"] + calcola_stats(df_i_p)]
                            for s in sorted(df_i_p['Set'].dropna().unique()):
                                ri_p.append([f"Set {int(float(s))}"] + calcola_stats(df_i_p[df_i_p['Set'] == s]))

                            st.dataframe(
                                pd.DataFrame(ri_p, columns=cols_h).style.apply(stile_zebra, axis=None).apply(stile_righe, axis=1).hide(axis="index").format({"Media Km/h": "{:.1f}"}, precision=1),
                                use_container_width=True
                            )
                        else:
                            rg_p = []
                            st.info("Nessun giocatore Perugia disponibile in questa selezione.")

                        st.divider()
                        st.markdown(f"## 🏐 {nome_avv} - PLAYER")
                        a_list = sorted([str(x) for x in df_o[df_o['Vel_Num'].notna() | df_o['Tipo'].isin(['V','N','F','v','n','f'])]['Player'].dropna().unique()])

                        if a_list:
                            col_fase_a, col_gioc_a = st.columns(2)
                            with col_fase_a:
                                fs_a = st.selectbox(f"Fase {nome_avv} (Tabella Generale):", ["MATCH"] + [f"Set {int(float(s))}" for s in sorted(df_o['Set'].dropna().unique())], key="fs_a_new")
                            with col_gioc_a:
                                ps_a = st.selectbox(f"Seleziona Giocatore {nome_avv} (Tabella Individuale):", a_list, key="ps_a_new")

                            rg_a = []
                            for a_name in a_list:
                                df_af = df_o[df_o['Player'] == a_name] if fs_a == "MATCH" else df_o[(df_o['Player'] == a_name) & (df_o['Set'].astype(float) == float(fs_a.replace("Set ","")))]
                                if not df_af.empty:
                                    rg_a.append([a_name] + calcola_stats(df_af))

                            st.dataframe(
                                pd.DataFrame(rg_a, columns=["Player"] + cols_h[1:]).style.apply(stile_zebra, axis=None).hide(axis="index").format({"Media Km/h": "{:.1f}"}, precision=1),
                                use_container_width=True
                            )

                            df_i_a = df_o[df_o['Player'] == ps_a]
                            ri_a = [["MATCH"] + calcola_stats(df_i_a)]
                            for s in sorted(df_i_a['Set'].dropna().unique()):
                                ri_a.append([f"Set {int(float(s))}"] + calcola_stats(df_i_a[df_i_a['Set'] == s]))

                            st.dataframe(
                                pd.DataFrame(ri_a, columns=cols_h).style.apply(stile_zebra, axis=None).apply(stile_righe, axis=1).hide(axis="index").format({"Media Km/h": "{:.1f}"}, precision=1),
                                use_container_width=True
                            )
                        else:
                            rg_a = []
                            st.info("Nessun giocatore avversario disponibile in questa selezione.")

                        df_p_raw = pd.DataFrame(rg_p, columns=["Player"] + cols_h[1:]) if rg_p else pd.DataFrame(columns=["Player"] + cols_h[1:])
                        df_o_raw = pd.DataFrame(rg_a, columns=["Player"] + cols_h[1:]) if rg_a else pd.DataFrame(columns=["Player"] + cols_h[1:])

                        if not df_p_raw.empty or not df_o_raw.empty:
                            buffer_player = io.BytesIO()
                            with pd.ExcelWriter(buffer_player, engine='xlsxwriter') as writer:
                                sdoppia_percentuali(df_p_raw).to_excel(writer, sheet_name='Perugia_Player')
                                if not df_o_raw.empty:
                                    sdoppia_percentuali(df_o_raw).to_excel(writer, sheet_name=f'{nome_avv}_Player'[:30])
                            st.download_button(
                                label="📥 Scarica Report Player",
                                data=buffer_player.getvalue(),
                                file_name="Report_Player_Match.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="btn_player_multi"
                            )

                        st.divider()
                        st.subheader("📊 Sequenza Cronologica Battute (BtxBT)")

                        set_disponibili = sorted(df_report['Set'].dropna().unique())
                        if set_disponibili:
                            set_scelto = st.selectbox("Seleziona il Set:", set_disponibili, format_func=lambda x: f"Set {int(float(x))}")
                            df_set = df_report[df_report['Set'] == set_scelto].copy()

                            nomi_team = df_report['Team'].dropna().unique()
                            team_sir = next((t for t in nomi_team if any(key in str(t).upper() for key in ['SIR', 'PERUGIA'])), None)
                            team_avv = next((t for t in nomi_team if t != team_sir), None)

                            def genera_tabella_per_team(df_input, nome_team_target):
                                if not nome_team_target:
                                    return pd.DataFrame()
                                df_team = df_input[df_input['Team'] == nome_team_target]
                                giocatori = sorted(df_team['Player'].dropna().unique())
                                if not giocatori:
                                    return pd.DataFrame()
                                data_rows = []
                                max_battute = df_team.groupby('Player').size().max()
                                for i in range(max_battute):
                                    row = {}
                                    for p_name in giocatori:
                                        p_battute = df_team[df_team['Player'] == p_name]
                                        if i < len(p_battute):
                                            b = p_battute.iloc[i]
                                            row[p_name] = f"{b['Tipo']} {b['Vel.']}"
                                        else:
                                            row[p_name] = ""
                                    data_rows.append(row)
                                return pd.DataFrame(data_rows)

                            def evidenzia_errori(val):
                                v = str(val).upper()
                                if ' N' in v or ' F' in v:
                                    return 'background-color: #ffffcc; font-weight: bold;'
                                return ''

                            st.markdown("#### 🏐 PERUGIA")
                            df_p_bt = genera_tabella_per_team(df_set, team_sir)
                            if not df_p_bt.empty:
                                st.dataframe(df_p_bt.style.applymap(evidenzia_errori), use_container_width=True, hide_index=True)
                            else:
                                st.info("Nessuna battuta per Perugia in questo set.")

                            st.markdown(f"#### 🏐 {nome_avv}")
                            df_a_bt = genera_tabella_per_team(df_set, team_avv)
                            if not df_a_bt.empty:
                                st.dataframe(df_a_bt.style.applymap(evidenzia_errori), use_container_width=True, hide_index=True)
                            else:
                                st.info("Nessuna battuta per l'avversario in questo set.")

                            if not df_p_bt.empty or not df_a_bt.empty:
                                buffer_bt = io.BytesIO()
                                with pd.ExcelWriter(buffer_bt, engine='xlsxwriter') as writer:
                                    sdoppia_btxbt(df_p_bt).to_excel(writer, sheet_name='Perugia')
                                    if not df_a_bt.empty:
                                        sdoppia_btxbt(df_a_bt).to_excel(writer, sheet_name='Avversario')
                                st.download_button(
                                    label=f"📥 Scarica BtxBT Set {int(float(set_scelto))}",
                                    data=buffer_bt.getvalue(),
                                    file_name=f"BtxBT_Set_{set_scelto}.xlsx",
                                    key=f"btn_btxbt_multi_{set_scelto}"
                                )
                    else:
                        st.subheader("📊 Analisi Visiva della Battuta")
                        df_graf = df_report[df_report['Vel_Num'].notna()].copy()
                        if not df_graf.empty:
                            st.markdown("##### 🏎️ Distribuzione Potenza")
                            fig_box = px.box(df_graf, x="Team", y="Vel_Num", color="Team", points="all",
                                             color_discrete_map={'PERUGIA': '#C41E3A', info['Avv.']: '#0047AB'})
                            st.plotly_chart(fig_box, use_container_width=True)

                            st.markdown("##### 📈 Trend Velocità per Set")
                            df_trend = df_graf.groupby(['Set', 'Team'])['Vel_Num'].mean().reset_index()
                            fig_line = px.line(df_trend, x="Set", y="Vel_Num", color="Team", markers=True)
                            st.plotly_chart(fig_line, use_container_width=True)

                            st.markdown("##### 🏆 Top 8 Performance")
                            df_top = df_graf.groupby(['Player', 'Team'])['Vel_Num'].mean().reset_index()
                            df_top = df_top.sort_values(by='Vel_Num', ascending=False).head(8)
                            fig_bar = px.bar(df_top, x='Vel_Num', y='Player', color='Team', orientation='h', text_auto='.1f')
                            st.plotly_chart(fig_bar, use_container_width=True)
                        else:
                            st.warning("Dati di velocità non disponibili.")

elif scelta == "Trend Team/Player":

    st.markdown(
        "<h2 style='text-align: center;'>📈 TREND STORICO PARTITE</h2>",
        unsafe_allow_html=True
    )

    # =========================================================
    # 1. CONTROLLO DATABASE
    # =========================================================
    if 'df_master' not in st.session_state or st.session_state['df_master'].empty:
        st.warning("⚠️ Carica prima il database master per visualizzare il Trend.")
    else:
        df_trend_base = st.session_state['df_master'].copy()

        colonne_necessarie = ['Data', 'Avv.', 'Team', 'Player', 'Vel.']
        colonne_mancanti = [c for c in colonne_necessarie if c not in df_trend_base.columns]

        if colonne_mancanti:
            st.error(f"❌ Mancano queste colonne nel database master: {colonne_mancanti}")
        else:
            # =========================================================
            # 2. PULIZIA DATI
            # =========================================================
            df_trend_base['Data_raw'] = df_trend_base['Data'].astype(str).str.strip()
            df_trend_base['Data'] = pd.to_datetime(
                df_trend_base['Data_raw'],
                errors='coerce',
                dayfirst=True
            )

            mask_nat = df_trend_base['Data'].isna()
            if mask_nat.any():
                df_trend_base.loc[mask_nat, 'Data'] = pd.to_datetime(
                    df_trend_base.loc[mask_nat, 'Data_raw'],
                    errors='coerce',
                    format='mixed'
                )

            df_trend_base = df_trend_base.dropna(subset=['Data']).copy()

            if df_trend_base.empty:
                st.warning("⚠️ Nessuna data valida trovata nel database master.")
            else:
                df_trend_base['Data_Solo'] = df_trend_base['Data'].dt.date
                df_trend_base['Avv.'] = df_trend_base['Avv.'].astype(str).str.strip()
                df_trend_base['Team'] = df_trend_base['Team'].astype(str).str.strip()
                df_trend_base['Player'] = df_trend_base['Player'].astype(str).str.strip()
                df_trend_base['Player_Key'] = (
                    df_trend_base['Player']
                    .astype(str)
                    .str.strip()
                    .str.replace(r'\s+', ' ', regex=True)
                    .str.upper()
                )
                df_trend_base['Player_Display'] = (
                    df_trend_base['Player_Key']
                    .str.lower()
                    .str.title()
                )
                df_trend_base['Vel_Str'] = df_trend_base['Vel.'].astype(str).str.upper().str.strip()
                df_trend_base['Vel_Num'] = df_trend_base['Vel.'].apply(clean_vel_val)

                # Battute SPIN: quelle con velocità numerica oppure con codice errore N/F nel campo velocità
                df_trend_base['Is_Spin'] = df_trend_base['Vel_Num'].notna() | df_trend_base['Vel_Str'].isin(['N', 'F'])
                df_trend_base['Is_Error'] = df_trend_base['Vel_Str'].isin(['N', 'F'])

                df_trend_base['Match_Label'] = (
                    df_trend_base['Data_Solo'].astype(str) + " - vs " + df_trend_base['Avv.']
                )

                match_info = (
                    df_trend_base[['Data', 'Match_Label']]
                    .drop_duplicates()
                    .sort_values('Data', ascending=False)
                )
                lista_match = match_info['Match_Label'].tolist()

                # =========================================================
                # 3. SCELTA MODALITÀ E FILTRI
                # =========================================================
                modalita = st.radio(
                    "Modalità analisi:",
                    options=["PERUGIA", "INDIVIDUALE"],
                    horizontal=True
                )

                giocatore_scelto = None
                giocatore_key = None
                if modalita == "INDIVIDUALE":
                    df_giocatori_perugia = df_trend_base[
                        df_trend_base['Team'].str.upper().str.contains('PERUGIA', na=False)
                    ][['Player_Key', 'Player_Display']].copy()
                    df_giocatori_perugia = (
                        df_giocatori_perugia
                        .replace('', pd.NA)
                        .dropna(subset=['Player_Key'])
                        .drop_duplicates(subset=['Player_Key'])
                        .sort_values('Player_Display')
                    )

                    lista_giocatori = df_giocatori_perugia['Player_Display'].tolist()
                    mappa_giocatori = dict(zip(df_giocatori_perugia['Player_Display'], df_giocatori_perugia['Player_Key']))

                    if not lista_giocatori:
                        st.warning("⚠️ Non risultano giocatori di Perugia nel database master.")
                        st.stop()

                    giocatore_scelto = st.selectbox(
                        "Scegli il giocatore di Perugia:",
                        options=lista_giocatori,
                        index=0
                    )
                    giocatore_key = mappa_giocatori.get(giocatore_scelto)

                selected_matches = st.multiselect(
                    "Scegli le partite:",
                    options=lista_match,
                    default=lista_match
                )

                if not selected_matches:
                    st.warning("Seleziona almeno una partita per vedere i grafici.")
                else:
                    df_trend_filtered = df_trend_base[
                        df_trend_base['Match_Label'].isin(selected_matches)
                    ].copy()

                    if df_trend_filtered.empty:
                        st.warning("Nessun dato disponibile per i match selezionati.")
                    else:
                        df_p_trend = df_trend_filtered[
                            df_trend_filtered['Team'].str.upper().str.contains('PERUGIA', na=False)
                        ].copy()

                        if df_p_trend.empty:
                            st.warning("⚠️ Nei match selezionati non ci sono dati per PERUGIA.")
                        else:
                            if modalita == "INDIVIDUALE":
                                df_target = df_p_trend[df_p_trend['Player_Key'] == giocatore_key].copy()
                                etichetta_metriche = f"Giocatore: {giocatore_scelto}"
                            else:
                                df_target = df_p_trend.copy()
                                etichetta_metriche = "Squadra: PERUGIA"

                            if df_target.empty:
                                st.warning("⚠️ Nessun dato disponibile per il filtro selezionato.")
                            else:
                                # =========================================================
                                # 4. COSTRUZIONE DATAFRAME TREND
                                # =========================================================
                                trend_data = []

                                for m_label in selected_matches:
                                    df_m = df_target[df_target['Match_Label'] == m_label].copy()

                                    if not df_m.empty:
                                        df_spin = df_m[df_m['Is_Spin']].copy()
                                        battute_tot = len(df_spin)
                                        media_vel = df_spin['Vel_Num'].mean() if battute_tot > 0 else None
                                        err_tot = int(df_spin['Is_Error'].sum()) if battute_tot > 0 else 0
                                        err_n = int((df_spin['Vel_Str'] == 'N').sum()) if battute_tot > 0 else 0
                                        err_f = int((df_spin['Vel_Str'] == 'F').sum()) if battute_tot > 0 else 0
                                        perc_err = (err_tot / battute_tot * 100) if battute_tot > 0 else 0
                                        perc_err_n = (err_n / battute_tot * 100) if battute_tot > 0 else 0
                                        perc_err_f = (err_f / battute_tot * 100) if battute_tot > 0 else 0

                                        data_match = df_m['Data_Solo'].iloc[0]
                                        avv_match = df_m['Avv.'].iloc[0]

                                        trend_data.append({
                                            "Data": data_match,
                                            "Gara": m_label,
                                            "Avversario": avv_match,
                                            "Avv_Breve": str(avv_match).strip()[:3].upper(),
                                            "Battute Totali": battute_tot,
                                            "Errori Totali": err_tot,
                                            "Errori N": err_n,
                                            "Errori F": err_f,
                                            "Media Km/h": round(media_vel, 1) if pd.notna(media_vel) else 0,
                                            "% Errori": round(perc_err, 1),
                                            "% Errori N": round(perc_err_n, 1),
                                            "% Errori F": round(perc_err_f, 1)
                                        })

                                df_plot = pd.DataFrame(trend_data)

                                if df_plot.empty:
                                    st.warning("Nessun dato utile da mostrare per i match selezionati.")
                                else:
                                    df_plot = df_plot.sort_values('Data').reset_index(drop=True)
                                    df_plot['X_Pos'] = list(range(1, len(df_plot) + 1))

                                    def linea_trend(valori_x, valori_y):
                                        if len(valori_y) >= 2:
                                            x_mean = sum(valori_x) / len(valori_x)
                                            y_mean = sum(valori_y) / len(valori_y)
                                            den = sum((x - x_mean) ** 2 for x in valori_x)
                                            if den != 0:
                                                m = sum((x - x_mean) * (y - y_mean) for x, y in zip(valori_x, valori_y)) / den
                                                b = y_mean - m * x_mean
                                                return [m * x + b for x in valori_x]
                                        media = sum(valori_y) / len(valori_y) if valori_y else 0
                                        return [media] * len(valori_y)

                                    x_vals = df_plot['X_Pos'].tolist()
                                    df_plot['Trend_Vel'] = linea_trend(x_vals, df_plot['Media Km/h'].fillna(0).tolist())
                                    df_plot['Trend_Err'] = linea_trend(x_vals, df_plot['% Errori'].fillna(0).tolist())
                                    df_plot['Trend_Err_N'] = linea_trend(x_vals, df_plot['% Errori N'].fillna(0).tolist())
                                    df_plot['Trend_Err_F'] = linea_trend(x_vals, df_plot['% Errori F'].fillna(0).tolist())

                                    # =========================================================
                                    # 5. METRICHE RIASSUNTIVE
                                    # =========================================================
                                    col1, col2, col3, col4 = st.columns(4)
                                    col1.metric("Partite selezionate", len(df_plot))
                                    col2.metric("Velocità media complessiva", f"{df_plot['Media Km/h'].mean():.1f} km/h")
                                    col3.metric("Errori medi", f"{df_plot['% Errori'].mean():.1f}%")
                                    col4.metric("Analisi", etichetta_metriche)

                                    st.divider()

                                    tickvals = df_plot['X_Pos'].tolist()
                                    ticktext = df_plot['Avv_Breve'].tolist()
                                    custom_hover = df_plot[['Avversario', 'Data', 'Battute Totali', 'Errori Totali']].astype(str).values

                                    # =========================================================
                                    # 6. GRAFICO VELOCITÀ MEDIA NEL TEMPO
                                    # =========================================================
                                    st.subheader("📈 Andamento Velocità Media nel Tempo")

                                    df_plot['Vel_Label'] = df_plot['Media Km/h'].map(lambda v: f"{v:.1f}")

                                    fig_and = px.line(
                                        df_plot,
                                        x='X_Pos',
                                        y='Media Km/h',
                                        text='Vel_Label',
                                        markers=True,
                                        color_discrete_sequence=['#C00000']
                                    )
                                    fig_and.update_traces(
                                        textposition="top center",
                                        texttemplate='%{text}',
                                        line=dict(width=3),
                                        marker=dict(size=9),
                                        hovertemplate="<b>%{customdata[0]}</b><br>Data: %{customdata[1]}<br>Spin totali: %{customdata[2]}<br>Velocità media: %{y:.1f} km/h<extra></extra>",
                                        customdata=custom_hover
                                    )
                                    fig_and.add_scatter(
                                        x=df_plot['X_Pos'],
                                        y=df_plot['Trend_Vel'],
                                        mode='lines',
                                        name='Trend',
                                        line=dict(color='#666666', width=2, dash='dash'),
                                        hovertemplate='Linea di tendenza<extra></extra>'
                                    )
                                    fig_and.update_layout(
                                        xaxis_title="",
                                        yaxis_title="Velocità Media (km/h)",
                                        legend_title_text='',
                                        height=450,
                                        xaxis=dict(
                                            tickmode='array',
                                            tickvals=tickvals,
                                            ticktext=ticktext
                                        )
                                    )
                                    st.plotly_chart(fig_and, use_container_width=True)

                                    # =========================================================
                                    # 7. GRAFICO ERRORI
                                    # =========================================================
                                    st.subheader("📉 Andamento ERRORE")

                                    fig_err = px.bar(
                                        df_plot,
                                        x="X_Pos",
                                        y="% Errori",
                                        text="% Errori",
                                        color_discrete_sequence=["#FFA500"]
                                    )
                                    fig_err.update_traces(
                                        texttemplate='%{text}%',
                                        textposition='outside',
                                        hovertemplate="<b>%{customdata[0]}</b><br>Data: %{customdata[1]}<br>Spin totali: %{customdata[2]}<br>Errori totali: %{customdata[3]}<br>% Errori: %{y:.1f}%<extra></extra>",
                                        customdata=custom_hover
                                    )
                                    fig_err.add_scatter(
                                        x=df_plot['X_Pos'],
                                        y=df_plot['Trend_Err'],
                                        mode='lines',
                                        name='Trend',
                                        line=dict(color='#666666', width=2, dash='dash'),
                                        hovertemplate='Linea di tendenza<extra></extra>'
                                    )
                                    fig_err.update_layout(
                                        xaxis_title="",
                                        yaxis_title="% Errori",
                                        xaxis=dict(
                                            tickmode='array',
                                            tickvals=tickvals,
                                            ticktext=ticktext,
                                            tickangle=0
                                        ),
                                        legend_title_text='',
                                        height=450
                                    )
                                    st.plotly_chart(fig_err, use_container_width=True)

                                    # =========================================================
                                    # 8. GRAFICO TIPOLOGIE DI ERRORE
                                    # =========================================================
                                    st.subheader("🔎 Tipologie di ERRORE")

                                    df_err_type = df_plot[['X_Pos', 'Avv_Breve', 'Avversario', 'Data', 'Battute Totali', 'Errori N', 'Errori F', '% Errori N', '% Errori F', 'Trend_Err_N', 'Trend_Err_F']].copy()
                                    df_err_long = df_err_type.melt(
                                        id_vars=['X_Pos', 'Avv_Breve', 'Avversario', 'Data', 'Battute Totali', 'Errori N', 'Errori F', 'Trend_Err_N', 'Trend_Err_F'],
                                        value_vars=['% Errori N', '% Errori F'],
                                        var_name='Tipo Errore',
                                        value_name='Percentuale'
                                    )
                                    df_err_long['Tipo Errore'] = df_err_long['Tipo Errore'].replace({
                                        '% Errori N': 'N',
                                        '% Errori F': 'F'
                                    })
                                    df_err_long['Trend'] = df_err_long.apply(
                                        lambda r: r['Trend_Err_N'] if r['Tipo Errore'] == 'N' else r['Trend_Err_F'],
                                        axis=1
                                    )
                                    df_err_long['Errori Tipo'] = df_err_long.apply(
                                        lambda r: r['Errori N'] if r['Tipo Errore'] == 'N' else r['Errori F'],
                                        axis=1
                                    )
                                    df_err_long['Label'] = df_err_long['Percentuale'].map(lambda v: f"{v:.1f}%" if v > 0 else "")

                                    fig_err_type = px.bar(
                                        df_err_long,
                                        x='X_Pos',
                                        y='Percentuale',
                                        color='Tipo Errore',
                                        barmode='group',
                                        text='Label',
                                        color_discrete_map={'N': '#D62728', 'F': '#1F77B4'},
                                        category_orders={'Tipo Errore': ['N', 'F']}
                                    )
                                    fig_err_type.update_traces(
                                        textposition='outside',
                                        hovertemplate="<b>%{customdata[0]}</b><br>Data: %{customdata[1]}<br>Spin totali: %{customdata[2]}<br>Tipo errore: %{customdata[3]}<br>Errori tipo: %{customdata[4]}<br>% Tipo errore: %{y:.1f}%<extra></extra>",
                                        customdata=df_err_long[['Avversario', 'Data', 'Battute Totali', 'Tipo Errore', 'Errori Tipo']].astype(str).values
                                    )
                                    for tipo, nome_trend, colore in [('N', 'Trend N', '#8B0000'), ('F', 'Trend F', '#0B4F8A')]:
                                        df_tipo = df_err_long[df_err_long['Tipo Errore'] == tipo].sort_values('X_Pos')
                                        fig_err_type.add_scatter(
                                            x=df_tipo['X_Pos'],
                                            y=df_tipo['Trend'],
                                            mode='lines',
                                            name=nome_trend,
                                            line=dict(color=colore, width=2, dash='dash'),
                                            hovertemplate=f'Linea di tendenza {tipo}<extra></extra>'
                                        )
                                    fig_err_type.update_layout(
                                        xaxis_title='',
                                        yaxis_title='% Errori per tipologia',
                                        xaxis=dict(
                                            tickmode='array',
                                            tickvals=tickvals,
                                            ticktext=ticktext,
                                            tickangle=0
                                        ),
                                        legend_title_text='',
                                        height=500
                                    )
                                    st.plotly_chart(fig_err_type, use_container_width=True)

                                    # =========================================================
                                    # 9. TABELLA FINALE
                                    # =========================================================
                                    st.write("### 📋 Dati Comparativi")
                                    colonne_tabella = ['Data', 'Avversario', 'Battute Totali', 'Errori Totali', 'Errori N', 'Errori F', 'Media Km/h', '% Errori', '% Errori N', '% Errori F']
                                    if modalita == "INDIVIDUALE":
                                        df_plot['Giocatore'] = giocatore_scelto
                                        colonne_tabella = ['Giocatore'] + colonne_tabella
                                    st.dataframe(
                                        df_plot[colonne_tabella],
                                        use_container_width=True,
                                        hide_index=True
                                    )


                                    pdf_bytes = build_trend_pdf(
                                        selected_matches=selected_matches,
                                        df_table_pdf=df_plot[colonne_tabella].copy(),
                                        df_plot_pdf=df_plot.copy(),
                                        modalita=modalita,
                                        etichetta_metriche=etichetta_metriche
                                    )
                                    st.download_button(
                                        label="📄 Scarica PDF Trend",
                                        data=pdf_bytes,
                                        file_name="trend_report_battuta.pdf",
                                        mime="application/pdf"
                                    )

elif scelta == "Storico Avversari":
    st.markdown("<h2 style='text-align: center;'>📚 STORICO AVVERSARI</h2>", unsafe_allow_html=True)

    df_base = st.session_state['df_master'].copy() if ('df_master' in st.session_state and not st.session_state['df_master'].empty) else df_master.copy()

    if df_base.empty:
        st.warning("⚠️ Carica prima il database per visualizzare lo Storico Avversari.")
    else:
        df_base = df_base.copy()

        # colonne minime richieste
        colonne_minime = ['Data', 'Avv.', 'Team', 'Tipo', 'Vel.']
        colonne_mancanti = [c for c in colonne_minime if c not in df_base.columns]
        if colonne_mancanti:
            st.error(f"❌ Mancano queste colonne nel database: {colonne_mancanti}")
        else:
            # pulizia base (conversione date robusta)
            df_base['Data_raw'] = df_base['Data'].astype(str).str.strip()
            df_base['Data'] = pd.to_datetime(df_base['Data_raw'], errors='coerce', dayfirst=True)
            mask_nat = df_base['Data'].isna()
            if mask_nat.any():
                df_base.loc[mask_nat, 'Data'] = pd.to_datetime(
                    df_base.loc[mask_nat, 'Data_raw'],
                    errors='coerce',
                    format='mixed'
                )
            df_base = df_base.dropna(subset=['Data']).copy()
            df_base['Data_Solo'] = df_base['Data'].dt.date
            df_base['Avv.'] = df_base['Avv.'].astype(str).str.strip()
            df_base['Team'] = df_base['Team'].astype(str).str.strip()
            df_base['Tipo'] = df_base['Tipo'].astype(str).str.upper().str.strip()
            df_base['Vel_Str'] = df_base['Vel.'].astype(str).str.strip().str.upper()
            df_base['Vel_Num'] = pd.to_numeric(df_base['Vel.'], errors='coerce')

            def _is_perugia(val):
                return 'PERUGIA' in str(val).upper()

            def _opponent_name(row):
                team = str(row['Team']).strip()
                avv = str(row['Avv.']).strip()
                if _is_perugia(team) and not _is_perugia(avv):
                    return avv
                if _is_perugia(avv) and not _is_perugia(team):
                    return team
                if not _is_perugia(avv):
                    return avv
                return team

            def _side_name(row):
                return 'Perugia' if _is_perugia(row['Team']) else 'Avversario'

            df_base['Opponent'] = df_base.apply(_opponent_name, axis=1).astype(str).str.strip()
            df_base['Side'] = df_base.apply(_side_name, axis=1)
            df_base['Match_Label'] = df_base['Data_Solo'].astype(str) + ' - vs ' + df_base['Opponent']

            match_info = (
                df_base[['Data', 'Match_Label']]
                .drop_duplicates()
                .sort_values('Data', ascending=False)
            )
            lista_match = match_info['Match_Label'].tolist()

            st.markdown("### ⚙️ Parametri")
            selected_matches = st.multiselect(
                "Seleziona le partite da analizzare:",
                options=lista_match,
                default=lista_match
            )

            if not selected_matches:
                st.warning("Seleziona almeno una partita per vedere i grafici.")
            else:
                df_sel = df_base[df_base['Match_Label'].isin(selected_matches)].copy()

                if df_sel.empty:
                    st.warning("Nessun dato disponibile per le partite selezionate.")
                else:
                    # Ordine partite: manteniamo esattamente quello selezionato nel filtro
                    ordine_match = selected_matches.copy()

                    # Etichette corte ma univoche per singola partita
                    def _short_match_label(lbl):
                        try:
                            data_txt, opp_txt = lbl.split(' - vs ', 1)
                            opp_txt = str(opp_txt).strip().upper()
                            data_txt = pd.to_datetime(data_txt, errors='coerce').strftime('%d-%m') if pd.notna(pd.to_datetime(data_txt, errors='coerce')) else data_txt
                            return f"{opp_txt} | {data_txt}"
                        except Exception:
                            return str(lbl).upper()

                    match_label_map = {lbl: _short_match_label(lbl) for lbl in ordine_match}
                    ordine_match_short = [match_label_map[lbl] for lbl in ordine_match]
                    df_sel['Match_Short'] = df_sel['Match_Label'].map(match_label_map)

                    # 1) grafico confronto velocità medie: solo SPIN, per singola partita, barre orizzontali
                    # Per coerenza con il resto dell'app, consideriamo SPIN solo le battute con velocità numerica valida
                    df_vel = df_sel.copy()
                    if 'clean_vel_val' in globals():
                        df_vel['Vel_Num_Spin'] = df_vel['Vel.'].apply(clean_vel_val)
                    else:
                        df_vel['Vel_Num_Spin'] = pd.to_numeric(df_vel['Vel.'], errors='coerce')
                    df_vel = df_vel[df_vel['Vel_Num_Spin'].notna()].copy()

                    if not df_vel.empty:
                        df_vel_plot = (
                            df_vel.groupby(['Match_Label', 'Match_Short', 'Side'])['Vel_Num_Spin']
                            .mean()
                            .reset_index()
                        )
                        if not df_vel_plot.empty:
                            df_vel_plot['Match_Short'] = pd.Categorical(
                                df_vel_plot['Match_Short'],
                                categories=list(reversed(ordine_match_short)),
                                ordered=True
                            )
                            df_vel_plot = df_vel_plot.sort_values(['Match_Short', 'Side'])

                            st.subheader("Confronto Velocità Medie SPIN: PERUGIA vs AVVERSARIO")
                            fig_vel = px.bar(
                                df_vel_plot,
                                y='Match_Short',
                                x='Vel_Num_Spin',
                                color='Side',
                                barmode='group',
                                text='Vel_Num_Spin',
                                orientation='h',
                                category_orders={'Match_Short': list(reversed(ordine_match_short)), 'Side': ['Perugia', 'Avversario']},
                                color_discrete_map={'Perugia': '#C00000', 'Avversario': '#1F77B4'},
                                hover_data={'Match_Label': True, 'Match_Short': False, 'Vel_Num_Spin': ':.1f'}
                            )
                            fig_vel.update_traces(texttemplate='%{text:.1f}', textposition='outside', cliponaxis=False)
                            fig_vel.update_layout(
                                xaxis_title='Velocità media SPIN (km/h)',
                                yaxis_title='',
                                legend_title='',
                                height=max(500, 55 * len(ordine_match_short))
                            )
                            st.plotly_chart(fig_vel, use_container_width=True)
                        else:
                            st.info("Nessuna battuta SPIN con velocità numerica valida trovata per il confronto Perugia/Avversario.")
                    else:
                        st.info("Nessuna battuta SPIN con velocità numerica valida trovata per il confronto Perugia/Avversario.")

                    # 2) grafico variazioni avversari (solo codice V), per singola partita, barre orizzontali
                    mask_var = (df_sel['Side'] == 'Avversario') & (
                        (df_sel['Tipo'] == 'V') | (df_sel['Vel_Str'] == 'V')
                    )
                    df_var = df_sel[mask_var].copy()

                    if not df_var.empty:
                        df_var_plot = (
                            df_var.groupby(['Match_Label', 'Match_Short'])
                            .size()
                            .reset_index(name='Variazioni')
                        )
                        df_var_plot['Match_Short'] = pd.Categorical(
                            df_var_plot['Match_Short'],
                            categories=list(reversed(ordine_match_short)),
                            ordered=True
                        )
                        df_var_plot = df_var_plot.sort_values('Match_Short')

                        st.subheader("Variazioni Avversario (codice V)")
                        fig_var = px.bar(
                            df_var_plot,
                            y='Match_Short',
                            x='Variazioni',
                            text='Variazioni',
                            orientation='h',
                            color_discrete_sequence=['#1F77B4'],
                            hover_data={'Match_Label': True, 'Match_Short': False}
                        )
                        fig_var.update_traces(textposition='outside', cliponaxis=False)
                        fig_var.update_layout(
                            xaxis_title='Numero assoluto variazioni',
                            yaxis_title='',
                            showlegend=False,
                            height=max(420, 50 * len(df_var_plot))
                        )
                        st.plotly_chart(fig_var, use_container_width=True)
                    else:
                        st.info("Nessuna variazione con codice V trovata per gli avversari contro Perugia.")

                    # 3) filtro fasce velocità + grafico battute avversarie sopra soglia, per singola partita
                    st.markdown("### 🎯 Fasce di velocità")
                    soglia_label = st.selectbox(
                        "Seleziona la fascia di velocità da analizzare:",
                        options=[">= 110 km/h", ">= 115 km/h", ">= 120 km/h"],
                        index=0,
                        key="storico_avv_fascia_vel"
                    )
                    soglia = int(soglia_label.split()[1])

                    df_band = df_sel[(df_sel['Side'] == 'Avversario') & (df_sel['Vel_Num_Spin'].notna() if 'Vel_Num_Spin' in df_sel.columns else df_sel['Vel_Num'].notna())].copy()
                    if 'Vel_Num_Spin' not in df_band.columns:
                        if 'clean_vel_val' in globals():
                            df_band['Vel_Num_Spin'] = df_band['Vel.'].apply(clean_vel_val)
                        else:
                            df_band['Vel_Num_Spin'] = pd.to_numeric(df_band['Vel.'], errors='coerce')
                    df_band = df_band[df_band['Vel_Num_Spin'] >= soglia].copy()

                    if not df_band.empty:
                        df_band_plot = (
                            df_band.groupby(['Match_Label', 'Match_Short'])
                            .size()
                            .reset_index(name='Battute_Fascia')
                        )
                        df_band_plot['Match_Short'] = pd.Categorical(
                            df_band_plot['Match_Short'],
                            categories=list(reversed(ordine_match_short)),
                            ordered=True
                        )
                        df_band_plot = df_band_plot.sort_values('Match_Short')

                        st.subheader(f"Battute Avversario nella fascia {soglia_label}")
                        fig_band = px.bar(
                            df_band_plot,
                            y='Match_Short',
                            x='Battute_Fascia',
                            text='Battute_Fascia',
                            orientation='h',
                            color_discrete_sequence=['#1F77B4'],
                            hover_data={'Match_Label': True, 'Match_Short': False}
                        )
                        fig_band.update_traces(textposition='outside', cliponaxis=False)
                        fig_band.update_layout(
                            xaxis_title='Numero assoluto battute',
                            yaxis_title='',
                            showlegend=False,
                            height=max(420, 50 * len(df_band_plot))
                        )
                        st.plotly_chart(fig_band, use_container_width=True)
                    else:
                        st.info(f"Nessuna battuta avversaria trovata nella fascia {soglia_label}.")

