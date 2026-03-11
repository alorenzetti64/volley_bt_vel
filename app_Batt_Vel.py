import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from github import Github, Auth
import io
import unicodedata
import time
import textwrap
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages


def parse_match_dates(series):
    """Converte date eterogenee del database in datetime.
    Gestisce anche formati tipo '07-mar' tipici dei file battute.
    """
    s = series.astype(str).str.strip()
    s = s.replace({'': pd.NA, 'nan': pd.NA, 'NaT': pd.NA, 'None': pd.NA})

    # 1) tentativi standard
    dt = pd.to_datetime(s, errors='coerce', dayfirst=True)
    mask = dt.isna() & s.notna()
    if mask.any():
        dt.loc[mask] = pd.to_datetime(s.loc[mask], errors='coerce', format='mixed', dayfirst=True)

    # 2) numeri seriali Excel
    mask = dt.isna() & s.notna() & s.str.fullmatch(r'\d+(?:\.0+)?', na=False)
    if mask.any():
        nums = pd.to_numeric(s.loc[mask], errors='coerce')
        dt.loc[mask] = pd.to_datetime(nums, unit='D', origin='1899-12-30', errors='coerce')

    # 3) formato breve senza anno: 07-mar / 7 set / 07_mar
    mask = dt.isna() & s.notna()
    if mask.any():
        mesi = {
            'gen': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'mag': 5, 'giu': 6,
            'lug': 7, 'ago': 8, 'set': 9, 'sett': 9, 'sep': 9,
            'ott': 10, 'nov': 11, 'dic': 12
        }
        oggi = pd.Timestamp.today()
        season_end_year = oggi.year if oggi.month <= 6 else oggi.year + 1

        estratti = s.loc[mask].str.lower().str.extract(r'^(\d{1,2})[\s\-\/_\.]+([a-zàéìòù]{3,4})(?:[\s\-\/_\.]+(\d{2,4}))?$')
        parsed = []
        for giorno, mese_txt, anno_txt in estratti.itertuples(index=False, name=None):
            if pd.isna(giorno) or pd.isna(mese_txt):
                parsed.append(pd.NaT)
                continue
            mese_key = str(mese_txt).strip().replace('.', '')[:4]
            mese = mesi.get(mese_key) or mesi.get(mese_key[:3])
            if not mese:
                parsed.append(pd.NaT)
                continue
            if pd.notna(anno_txt):
                anno = int(anno_txt)
                if anno < 100:
                    anno += 2000
            else:
                anno = season_end_year if mese <= 6 else season_end_year - 1
            try:
                parsed.append(pd.Timestamp(year=anno, month=int(mese), day=int(giorno)))
            except Exception:
                parsed.append(pd.NaT)
        dt.loc[mask] = pd.Series(parsed, index=s.loc[mask].index)

    return dt

def _rimuovi_accenti_testo(valore):
    valore = str(valore)
    valore = unicodedata.normalize("NFKD", valore)
    return "".join(ch for ch in valore if not unicodedata.combining(ch))


def normalizza_nomi_giocatori(df, colonna='Player'):
    if colonna not in df.columns:
        return df
    df = df.copy()
    s = df[colonna].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)

    alias_display = {
        'PLOTNYSKYI': 'PLOTNYTSKYI',
        'PLOTNYTSKYI': 'PLOTNYTSKYI',
        'SE': 'SEMENIUK',
        'SEMENIUK': 'SEMENIUK',
        'SOLE': 'SOLÉ',
        'SOLÉ': 'SOLÉ',
    }

    def _canon(v):
        raw = str(v).strip()
        if raw == '' or raw.lower() in ('nan', 'none'):
            return raw
        up = raw.upper()
        up_noacc = _rimuovi_accenti_testo(up)
        return alias_display.get(up, alias_display.get(up_noacc, up))

    df[colonna] = s.apply(_canon)
    return df

def _safe_pdf_text(value):
    return str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")



def build_player_sheet_pdf(player_name, selected_matches, metrics_dict, df_table_pdf):
    """Crea un PDF semplice per la Scheda Battitore."""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    import tempfile

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    doc = SimpleDocTemplate(tmp.name, pagesize=A4, leftMargin=32, rightMargin=32, topMargin=32, bottomMargin=32)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(f"Scheda Battitore - {player_name}", styles["Title"]))
    story.append(Spacer(1, 10))
    story.append(Paragraph(f"Partite selezionate: {selected_matches}", styles["Normal"]))
    story.append(Spacer(1, 10))

    metric_rows = [["Indicatore", "Valore"]]
    for k, v in metrics_dict.items():
        metric_rows.append([str(k), str(v)])

    metric_table = Table(metric_rows, repeatRows=1)
    metric_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8EEF7")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8F8F8")]),
    ]))
    story.append(metric_table)
    story.append(Spacer(1, 14))

    if df_table_pdf is not None and not df_table_pdf.empty:
        data = [list(df_table_pdf.columns)] + df_table_pdf.astype(str).values.tolist()
        table = Table(data, repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDEBF7")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#FAFAFA")]),
            ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
        ]))
        story.append(Paragraph("Dettaglio partite", styles["Heading2"]))
        story.append(Spacer(1, 6))
        story.append(table)

    doc.build(story)
    with open(tmp.name, "rb") as f:
        return f.read()


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
    """Prepara un DataFrame piatto per Excel, evitando MultiIndex.
    Le colonne tipo '12 (34.5%)' vengono divise in due colonne: N e %.
    """
    if df is None or df.empty:
        return df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()

    cols_to_fix = [
        ">=120", ">=115 <120", ">=110 <115", ">=100 <110", "<100",
        "[V] var.ni", "Errori [N]+[F]", "Rete [N]", "Fuori [NF]"
    ]

    df_out = pd.DataFrame(index=df.index)

    for col in df.columns:
        if col in cols_to_fix:
            s = df[col].astype(str)
            num = pd.to_numeric(s.str.extract(r'(\d+)').iloc[:, 0], errors='coerce')
            pct = pd.to_numeric(s.str.extract(r'\(([-\d.,]+)%\)').iloc[:, 0].str.replace(',', '.', regex=False), errors='coerce') / 100
            df_out[f'{col} N'] = num
            df_out[f'{col} %'] = pct
        else:
            df_out[col] = df[col]

    return df_out




def uniforma_decimali(df):
    """Arrotonda a un solo decimale tutte le colonne di velocità, percentuali e indici."""
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return df
    df = df.copy()
    chiavi = ['km/h', '%', 'index']
    for col in df.columns:
        col_str = str(col).lower()
        if any(k in col_str for k in chiavi):
            df[col] = pd.to_numeric(df[col], errors='coerce').round(1)
    return df

def sdoppia_btxbt(df):
    """Prepara un DataFrame piatto per Excel, evitando MultiIndex.
    Ogni colonna viene sdoppiata in Tipo e Vel/Err.
    """
    if df is None or df.empty:
        return df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()

    df_out = pd.DataFrame(index=df.index)
    for col in df.columns:
        s = df[col].astype(str)
        df_out[f'{col} Tipo'] = s.str.extract(r'([A-Za-z]+)').iloc[:, 0]
        df_out[f'{col} Vel/Err'] = s.str.extract(r'(\d+|[VNFE])').iloc[:, 0]

    return df_out

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

    df['TipoU'] = df['Tipo'].astype(str).str.upper().str.strip()
    df['VelS'] = df['Vel.'].astype(str).str.upper().str.strip()

    def clean_spin_numeric(val):
        s = str(val).strip().upper()
        if s in ['', 'NAN', 'NONE', 'V', 'N', 'F']:
            return None
        try:
            return float(s.replace(',', '.'))
        except Exception:
            return None

    # CONTA SOLO LE VERE SPIN
    df_spin = df[df['TipoU'] == 'SPIN'].copy()
    df_spin['Vel_Num'] = df_spin['Vel.'].apply(clean_spin_numeric)

    tot = len(df_spin)
    df_spin_valide = df_spin[df_spin['Vel_Num'].notna()].copy()
    n_spin = len(df_spin_valide)

    media = df_spin_valide['Vel_Num'].mean() if n_spin > 0 else 0

    n_var = int((df_spin['VelS'] == 'V').sum())
    n_net = int((df_spin['VelS'] == 'N').sum())
    n_out = int((df_spin['VelS'] == 'F').sum())
    n_err_tot = n_net + n_out

    def fmt_f(c, t):
        p = (c / t * 100) if t > 0 else 0
        return f"{c} ({p:.1f}%)"

    f1 = fmt_f(len(df_spin_valide[df_spin_valide['Vel_Num'] >= 120]), n_spin)
    f2 = fmt_f(len(df_spin_valide[(df_spin_valide['Vel_Num'] >= 115) & (df_spin_valide['Vel_Num'] < 120)]), n_spin)
    f3 = fmt_f(len(df_spin_valide[(df_spin_valide['Vel_Num'] >= 110) & (df_spin_valide['Vel_Num'] < 115)]), n_spin)
    f4 = fmt_f(len(df_spin_valide[(df_spin_valide['Vel_Num'] >= 100) & (df_spin_valide['Vel_Num'] < 110)]), n_spin)
    f5 = fmt_f(len(df_spin_valide[df_spin_valide['Vel_Num'] < 100]), n_spin)

    var_str = fmt_f(n_var, tot)
    err_str = fmt_f(n_err_tot, tot)
    net_str = fmt_f(n_net, n_err_tot) if n_err_tot > 0 else "0 (0.0%)"
    out_str = fmt_f(n_out, n_err_tot) if n_err_tot > 0 else "0 (0.0%)"

    return [tot, n_spin, media, f1, f2, f3, f4, f5, var_str, err_str, net_str, out_str]

def stile_righe(row):
    fase_val = str(row.iloc[0]) 
    if fase_val == 'MATCH': return ['background-color: #ffffcc'] * len(row)
    if 'Set 2' in fase_val or 'Set 4' in fase_val: return ['background-color: #f2f2f2'] * len(row)
    return [''] * len(row)

# --- 3. FUNZIONI GITHUB ---
def get_github_client():
    token = st.secrets["github"].get("access_token") or st.secrets["github"].get("token")
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
scelta = st.sidebar.radio("Scegli:", ["Caricamento Dati", "Report Partita", "Trend Team/Player", "Scheda Battitore", "Confronto Partite", "Storico Avversari", "Ranking Battitori", "Insight"])

# Usa prima i dati caricati nella sessione; solo in assenza totale ripiega su GitHub.
if 'df_master' in st.session_state and isinstance(st.session_state['df_master'], pd.DataFrame):
    df_master = normalizza_nomi_giocatori(st.session_state['df_master'].copy(), 'Player')
    st.session_state['df_master'] = df_master.copy()
else:
    df_master = normalizza_nomi_giocatori(load_master_from_github(), 'Player')
    st.session_state['df_master'] = df_master.copy()

import warnings
# Nasconde i messaggi di avviso di openpyxl e pandas
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
warnings.filterwarnings("ignore", category=FutureWarning)

def load_uploaded_match_file(uploaded_file):
    """Legge CSV/XLSM/XLSX e restituisce il dataframe normalizzato con le sole 8 colonne utili."""
    file_name = uploaded_file.name.lower()

    if file_name.endswith('.csv'):
        try:
            df_raw = pd.read_csv(uploaded_file, sep=';', engine='python', dtype=str)
            if df_raw.shape[1] < 8:
                uploaded_file.seek(0)
                df_raw = pd.read_csv(uploaded_file, sep=',', engine='python', dtype=str)
        except Exception:
            uploaded_file.seek(0)
            df_raw = pd.read_csv(uploaded_file, sep=',', engine='python', dtype=str)
    else:
        df_raw = pd.read_excel(uploaded_file, sheet_name="Foglio1", engine="openpyxl", dtype=str)

    if df_raw is None or df_raw.shape[1] < 8:
        raise ValueError("Il file non contiene almeno 8 colonne utili nelle prime colonne.")

    df_new = df_raw.iloc[:, 0:8].copy()
    df_new.columns = COLUMNS_A_H

    # Riempie eventuali celle vuote dovute a merge / celle ripetute mancanti
    for col in ['Data', 'Partita', 'Avv.', 'Team', 'Set']:
        df_new[col] = df_new[col].ffill()

    # Normalizzazione
    df_new['Data'] = df_new['Data'].astype(str).str.strip()
    df_new['Partita'] = df_new['Partita'].astype(str).str.strip()
    df_new['Avv.'] = df_new['Avv.'].astype(str).str.upper().str.strip()
    df_new['Team'] = df_new['Team'].astype(str).str.upper().str.strip()
    df_new['Set'] = df_new['Set'].astype(str).str.strip()
    df_new['Player'] = df_new['Player'].astype(str).str.strip()
    df_new = normalizza_nomi_giocatori(df_new, 'Player')
    df_new['Tipo'] = df_new['Tipo'].astype(str).str.upper().str.strip()
    df_new['Vel.'] = df_new['Vel.'].astype(str).str.upper().str.strip()

    # Togli righe sporche
    df_new = df_new.dropna(subset=['Player'], how='all')
    df_new = df_new[df_new['Player'].ne('')]
    df_new = df_new[df_new['Data'].str.upper() != 'NAT']
    df_new = df_new[df_new['Tipo'].isin(['SPIN', 'FLOAT'])]

    return df_new.reset_index(drop=True)

if scelta == "Caricamento Dati":
    st.title("🚀 Database")
    tab1, tab2 = st.tabs(["Carica CSV / Excel", "Pulisci Database"])
    
    with tab1:
        uploaded = st.file_uploader("Seleziona file CSV / XLSM / XLSX")
        st.caption("Formati supportati: .csv, .xlsm, .xlsx")
        if uploaded:
            file_name_check = uploaded.name.lower()
            if not file_name_check.endswith((".csv", ".xlsm", ".xlsx")):
                st.error("Carica un file .csv, .xlsm oppure .xlsx")
                st.stop()
            try:
                df_new = load_uploaded_match_file(uploaded)
            except Exception as e:
                st.error(f"Errore nel caricamento del file: {e}")
                st.stop()

            st.write("DEBUG IMPORT - Team x Tipo:")
            st.write(df_new.groupby(['Team', 'Tipo']).size().reset_index(name='conteggio'))

            df_spin_debug = df_new[df_new['Tipo'].eq('SPIN')].copy()
            df_spin_debug['Vel_Num'] = df_spin_debug['Vel.'].apply(clean_vel_val)
            st.write("DEBUG IMPORT - Spin totali / spin con velocità numerica:")
            st.write(
                df_spin_debug.groupby('Team').agg(
                    Spin_Totali=('Tipo', 'size'),
                    Spin_Valide=('Vel_Num', lambda x: int(x.notna().sum())),
                    Variazioni_V=('Vel.', lambda x: int(x.astype(str).eq('V').sum())),
                    Errori_N=('Vel.', lambda x: int(x.astype(str).eq('N').sum())),
                    Errori_F=('Vel.', lambda x: int(x.astype(str).eq('F').sum())),
                ).reset_index()
            )

            if st.button("🚀 Sincronizza su GitHub"):
                with st.spinner("Sincronizzazione in corso..."):
                    df_combined = normalizza_nomi_giocatori(df_master.copy(), 'Player')

                    # SOSTITUISCE la stessa partita, non la somma
                    chiavi_match = df_new[['Data', 'Avv.', 'Partita']].drop_duplicates()

                    for _, r in chiavi_match.iterrows():
                        mask = (
                            df_combined['Data'].astype(str).str.strip().eq(str(r['Data']).strip()) &
                            df_combined['Avv.'].astype(str).str.upper().str.strip().eq(str(r['Avv.']).strip()) &
                            df_combined['Partita'].astype(str).str.strip().eq(str(r['Partita']).strip())
                        )
                        df_combined = df_combined[~mask]

                    df_combined = pd.concat([df_combined, df_new], ignore_index=True)
                    df_combined = normalizza_nomi_giocatori(df_combined, 'Player')

                    st.session_state['df_master'] = df_combined.copy()
                    save_to_github(df_combined)

                    st.success("Dati sincronizzati!")
                    st.balloons()
                    time.sleep(2)
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
                            (df_master['Data'].astype(str).str.strip() == str(row['Data']).strip()) &
                            (df_master['Avv.'].astype(str).str.upper().str.strip() == str(row['Avv.']).strip().upper()) &
                            (df_master['Partita'].astype(str).str.strip() == str(row['Partita']).strip())
                        )]
                    
                    st.session_state['df_master'] = df_master.copy()
                    save_to_github(df_master)
                    st.success("Database aggiornato con successo!")
                    st.rerun()
            else:
                st.info("Nessuna partita selezionata. Usa le caselle sopra per procedere.")
        else:
            st.info("Il database è attualmente vuoto.")

elif scelta == "Report Partita":
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
                    df_report['Set_Num'] = pd.to_numeric(df_report['Set'], errors='coerce')
                    df_report['Tipo_Clean'] = df_report['Tipo'].astype(str).str.upper().str.strip()
                    df_report['Team_Clean'] = df_report['Team'].astype(str).str.upper().str.strip()

                    info = df_report.iloc[0]
                    nome_avv = str(info['Avv.']).upper()
                    cols_h = ["Fase", "Tot.SPIN", "Spin valide", "Media Km/h", ">=120", ">=115 <120", ">=110 <115", ">=100 <110", "<100", "[V] var.ni", "Errori [N]+[F]", "Rete [N]", "Fuori [NF]"]

                    mask_perugia = df_report['Team_Clean'].str.contains('PERUGIA', na=False) | df_report['Team_Clean'].str.contains('SIR', na=False)
                    df_p = df_report[mask_perugia].copy()
                    df_o = df_report[~mask_perugia].copy()

                    set_disponibili_report = sorted(df_report['Set_Num'].dropna().unique())
                    set_labels_report = [f"Set {int(float(s))}" for s in set_disponibili_report]
                    tipo_errori = ['V', 'N', 'F']

                    def build_team_rows(df_team):
                        rows = []
                        if df_team.empty:
                            return rows
                        rows.append(["MATCH"] + calcola_stats(df_team))
                        for s in sorted(df_team['Set_Num'].dropna().unique()):
                            rows.append([f"Set {int(float(s))}"] + calcola_stats(df_team[df_team['Set_Num'] == s]))
                        return rows

                    def build_player_rows(df_team, fase_label):
                        if df_team.empty:
                            return []
                        if fase_label == "MATCH":
                            df_fase = df_team.copy()
                        else:
                            set_num = float(fase_label.replace("Set ", ""))
                            df_fase = df_team[df_team['Set_Num'] == set_num].copy()
                        player_list = sorted([
                            str(x) for x in df_fase[
                                df_fase['Vel_Num'].notna() | df_fase['Tipo_Clean'].isin(tipo_errori)
                            ]['Player'].dropna().unique() if str(x).strip() != ''
                        ])
                        rows = []
                        for player_name in player_list:
                            df_player = df_fase[df_fase['Player'] == player_name]
                            if not df_player.empty:
                                rows.append([player_name] + calcola_stats(df_player))
                        return rows

                    def build_single_player_rows(df_team, player_name):
                        if df_team.empty or not player_name:
                            return []
                        df_player = df_team[df_team['Player'] == player_name].copy()
                        if df_player.empty:
                            return []
                        rows = [["MATCH"] + calcola_stats(df_player)]
                        for s in sorted(df_player['Set_Num'].dropna().unique()):
                            rows.append([f"Set {int(float(s))}"] + calcola_stats(df_player[df_player['Set_Num'] == s]))
                        return rows

                    def genera_tabella_per_team(df_input, nome_team_target):
                        if not nome_team_target:
                            return pd.DataFrame()
                        df_team = df_input[df_input['Team'] == nome_team_target].copy()
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

                    if m_rep == "REPORT":
                        st.markdown("<h2 style='text-align: center;'>📋 REPORT VELOCITÀ BATTUTA SPIN</h2>", unsafe_allow_html=True)
                        col1, col2, col3 = st.columns(3)
                        col1.success(f"**Manifestazione**\n\n{info['Partita']}")
                        col2.success(f"**Data**\n\n{info['Data']}")
                        col3.success(f"**Avversario**\n\n{nome_avv}")

                        st.markdown("### 🏐 PERUGIA")
                        r_p = build_team_rows(df_p)
                        if r_p:
                            st.dataframe(
                                pd.DataFrame(r_p, columns=cols_h)
                                .style.hide(axis='index')
                                .apply(stile_righe, axis=1)
                                .format({"Media Km/h": "{:.1f}"}),
                                use_container_width=True
                            )
                        else:
                            st.info("Nessun dato disponibile per Perugia.")

                        st.markdown(f"### 🏐 {nome_avv}")
                        r_o = build_team_rows(df_o)
                        if r_o:
                            st.dataframe(
                                pd.DataFrame(r_o, columns=cols_h)
                                .style.hide(axis='index')
                                .apply(stile_righe, axis=1)
                                .format({"Media Km/h": "{:.1f}"}),
                                use_container_width=True
                            )
                        else:
                            st.info("Nessun dato disponibile per l'avversario.")

                        if r_p or r_o:
                            buffer_squadre = io.BytesIO()
                            with pd.ExcelWriter(buffer_squadre, engine='xlsxwriter') as writer:
                                if r_p:
                                    uniforma_decimali(sdoppia_percentuali(pd.DataFrame(r_p, columns=cols_h))).to_excel(writer, sheet_name='Perugia', index=False)
                                if r_o:
                                    uniforma_decimali(sdoppia_percentuali(pd.DataFrame(r_o, columns=cols_h))).to_excel(writer, sheet_name=nome_avv[:30], index=False)
                            st.download_button(
                                label="📥 Scarica Report Squadre",
                                data=buffer_squadre.getvalue(),
                                file_name="Report_Generale_Squadre.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="btn_squadre_multi"
                            )

                        st.divider()
                        st.markdown("## 🏐 PERUGIA - PLAYER")
                        fasi_perugia = ["MATCH"] + [f"Set {int(float(s))}" for s in sorted(df_p['Set_Num'].dropna().unique())]
                        fase_p_default = fasi_perugia if fasi_perugia else ["MATCH"]
                        rg_p = []
                        if not df_p.empty:
                            col_fase_p, col_gioc_p = st.columns(2)
                            with col_fase_p:
                                fs_p = st.selectbox("Fase Perugia:", fase_p_default, key="fs_p_new")
                            rg_p = build_player_rows(df_p, fs_p)
                            p_list = [r[0] for r in rg_p]
                            if p_list:
                                with col_gioc_p:
                                    ps_p = st.selectbox("Giocatore Perugia:", p_list, key="ps_p_new")
                                st.dataframe(
                                    pd.DataFrame(rg_p, columns=["Player"] + cols_h[1:])
                                    .style.apply(stile_zebra, axis=None)
                                    .hide(axis="index")
                                    .format({"Media Km/h": "{:.1f}"}, precision=1),
                                    use_container_width=True
                                )
                                ri_p = build_single_player_rows(df_p, ps_p)
                                st.dataframe(
                                    pd.DataFrame(ri_p, columns=cols_h)
                                    .style.apply(stile_zebra, axis=None)
                                    .apply(stile_righe, axis=1)
                                    .hide(axis="index")
                                    .format({"Media Km/h": "{:.1f}"}, precision=1),
                                    use_container_width=True
                                )
                            else:
                                st.info("Nessun giocatore Perugia disponibile in questa selezione.")
                        else:
                            st.info("Nessun dato disponibile per Perugia.")

                        st.divider()
                        st.markdown(f"## 🏐 {nome_avv} - PLAYER")
                        fasi_avv = ["MATCH"] + [f"Set {int(float(s))}" for s in sorted(df_o['Set_Num'].dropna().unique())]
                        fase_a_default = fasi_avv if fasi_avv else ["MATCH"]
                        rg_a = []
                        if not df_o.empty:
                            col_fase_a, col_gioc_a = st.columns(2)
                            with col_fase_a:
                                fs_a = st.selectbox(f"Fase {nome_avv}:", fase_a_default, key="fs_a_new")
                            rg_a = build_player_rows(df_o, fs_a)
                            a_list = [r[0] for r in rg_a]
                            if a_list:
                                with col_gioc_a:
                                    ps_a = st.selectbox(f"Giocatore {nome_avv}:", a_list, key="ps_a_new")
                                st.dataframe(
                                    pd.DataFrame(rg_a, columns=["Player"] + cols_h[1:])
                                    .style.apply(stile_zebra, axis=None)
                                    .hide(axis="index")
                                    .format({"Media Km/h": "{:.1f}"}, precision=1),
                                    use_container_width=True
                                )
                                ri_a = build_single_player_rows(df_o, ps_a)
                                st.dataframe(
                                    pd.DataFrame(ri_a, columns=cols_h)
                                    .style.apply(stile_zebra, axis=None)
                                    .apply(stile_righe, axis=1)
                                    .hide(axis="index")
                                    .format({"Media Km/h": "{:.1f}"}, precision=1),
                                    use_container_width=True
                                )
                            else:
                                st.info("Nessun giocatore avversario disponibile in questa selezione.")
                        else:
                            st.info("Nessun dato disponibile per l'avversario.")

                        df_p_raw = pd.DataFrame(rg_p, columns=["Player"] + cols_h[1:]) if rg_p else pd.DataFrame(columns=["Player"] + cols_h[1:])
                        df_o_raw = pd.DataFrame(rg_a, columns=["Player"] + cols_h[1:]) if rg_a else pd.DataFrame(columns=["Player"] + cols_h[1:])

                        if not df_p_raw.empty or not df_o_raw.empty:
                            buffer_player = io.BytesIO()
                            with pd.ExcelWriter(buffer_player, engine='xlsxwriter') as writer:
                                if not df_p_raw.empty:
                                    uniforma_decimali(sdoppia_percentuali(df_p_raw)).to_excel(writer, sheet_name='Perugia_Player', index=False)
                                if not df_o_raw.empty:
                                    uniforma_decimali(sdoppia_percentuali(df_o_raw)).to_excel(writer, sheet_name=f'{nome_avv}_Player'[:30], index=False)
                            st.download_button(
                                label="📥 Scarica Report Player",
                                data=buffer_player.getvalue(),
                                file_name="Report_Player_Match.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="btn_player_multi"
                            )

                        st.divider()
                        st.subheader("📊 Sequenza Cronologica Battute (BtxBT)")
                        if set_disponibili_report:
                            set_scelto = st.selectbox(
                                "Seleziona il Set:",
                                set_disponibili_report,
                                format_func=lambda x: f"Set {int(float(x))}"
                            )
                            df_set = df_report[df_report['Set_Num'] == float(set_scelto)].copy()

                            nomi_team = [x for x in df_set['Team'].dropna().unique() if str(x).strip() != '']
                            team_sir = next((t for t in nomi_team if any(key in str(t).upper() for key in ['SIR', 'PERUGIA'])), None)
                            team_avv = next((t for t in nomi_team if t != team_sir), None)

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
                                    if not df_p_bt.empty:
                                        uniforma_decimali(sdoppia_btxbt(df_p_bt)).to_excel(writer, sheet_name='Perugia', index=False)
                                    if not df_a_bt.empty:
                                        uniforma_decimali(sdoppia_btxbt(df_a_bt)).to_excel(writer, sheet_name='Avversario', index=False)
                                st.download_button(
                                    label=f"📥 Scarica BtxBT Set {int(float(set_scelto))}",
                                    data=buffer_bt.getvalue(),
                                    file_name=f"BtxBT_Set_{int(float(set_scelto))}.xlsx",
                                    key=f"btn_btxbt_multi_{int(float(set_scelto))}"
                                )
                        else:
                            st.info("Nessun set disponibile per la sequenza cronologica delle battute.")

                    elif m_rep == "GRAFICI":
                        st.markdown("<h2 style='text-align: center;'>📈 GRAFICI BATTUTA</h2>", unsafe_allow_html=True)
                        col1, col2, col3 = st.columns(3)
                        col1.success(f"**Manifestazione**\n\n{info['Partita']}")
                        col2.success(f"**Data**\n\n{info['Data']}")
                        col3.success(f"**Avversario**\n\n{nome_avv}")

                        df_graf = df_report[df_report['Vel_Num'].notna()].copy()
                        if not df_graf.empty:
                            st.markdown("##### 🏎️ Distribuzione Potenza")
                            fig_box = px.box(df_graf, x="Team", y="Vel_Num", color="Team", points="all")
                            st.plotly_chart(fig_box, use_container_width=True)

                            st.markdown("##### 📈 Trend Velocità per Set")
                            df_trend = df_graf.groupby(['Set_Num', 'Team'])['Vel_Num'].mean().reset_index()

                            fig_line = go.Figure()
                            for team_name in df_trend['Team'].dropna().unique():
                                df_team = df_trend[df_trend['Team'] == team_name].sort_values('Set_Num')
                                fig_line.add_trace(go.Scatter(
                                    x=df_team['Set_Num'],
                                    y=df_team['Vel_Num'],
                                    mode='lines+markers+text',
                                    name=team_name,
                                    text=[f"{v:.1f}" for v in df_team['Vel_Num']],
                                    textposition='top center'
                                ))

                            df_diff = (
                                df_trend.pivot(index='Set_Num', columns='Team', values='Vel_Num')
                                .dropna()
                                .sort_index()
                            )
                            if df_diff.shape[1] >= 2:
                                team_a = df_diff.columns[0]
                                team_b = df_diff.columns[1]
                                for set_num, row in df_diff.iterrows():
                                    y1 = float(row[team_a])
                                    y2 = float(row[team_b])
                                    y_mid = (y1 + y2) / 2
                                    diff = abs(y1 - y2)
                                    fig_line.add_shape(
                                        type='line',
                                        x0=set_num,
                                        x1=set_num,
                                        y0=min(y1, y2),
                                        y1=max(y1, y2),
                                        line=dict(width=2, dash='dot', color='rgba(80,80,80,0.7)'),
                                    )
                                    fig_line.add_annotation(
                                        x=set_num,
                                        y=y_mid,
                                        text=f"Δ {diff:.1f}",
                                        showarrow=False,
                                        xshift=18,
                                        bgcolor='rgba(255,255,255,0.75)',
                                        bordercolor='rgba(80,80,80,0.35)',
                                        borderwidth=1
                                    )

                            fig_line.update_layout(
                                xaxis_title="Set",
                                yaxis_title="Velocità media (km/h)",
                                xaxis=dict(
                                    tickmode='linear',
                                    tick0=1,
                                    dtick=1
                                )
                            )
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
            df_trend_base['Data'] = parse_match_dates(df_trend_base['Data_raw'])
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
                                        uniforma_decimali(df_plot[colonne_tabella]),
                                        use_container_width=True,
                                        hide_index=True
                                    )


                                    pdf_bytes = build_trend_pdf(
                                        selected_matches=selected_matches,
                                        df_table_pdf=uniforma_decimali(df_plot[colonne_tabella].copy()),
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



elif scelta == "Scheda Battitore":
    st.markdown("<h2 style='text-align: center;'>🎯 SCHEDA BATTITORE</h2>", unsafe_allow_html=True)

    df_base = st.session_state['df_master'].copy() if ('df_master' in st.session_state and not st.session_state['df_master'].empty) else df_master.copy()

    if df_base.empty:
        st.warning("⚠️ Carica prima il database per visualizzare la Scheda Battitore.")
    else:
        colonne_minime = ['Data', 'Avv.', 'Team', 'Player', 'Vel.']
        colonne_mancanti = [c for c in colonne_minime if c not in df_base.columns]
        if colonne_mancanti:
            st.error(f"❌ Mancano queste colonne nel database: {colonne_mancanti}")
        else:
            df_base = df_base.copy()
            df_base['Data_raw'] = df_base['Data'].astype(str).str.strip()
            df_base['Data'] = parse_match_dates(df_base['Data_raw'])
            df_base = df_base.dropna(subset=['Data']).copy()
            df_base['Data_Solo'] = df_base['Data'].dt.date
            df_base['Avv.'] = df_base['Avv.'].astype(str).str.strip()
            df_base['Team'] = df_base['Team'].astype(str).str.strip()
            df_base['Player'] = df_base['Player'].astype(str).str.strip()
            df_base['Vel_Str'] = df_base['Vel.'].astype(str).str.upper().str.strip()
            df_base['Vel_Num'] = df_base['Vel.'].apply(clean_vel_val)
            df_base['Is_Spin'] = df_base['Vel_Num'].notna() | df_base['Vel_Str'].isin(['N', 'F'])
            df_base['Is_Error'] = df_base['Vel_Str'].isin(['N', 'F'])
            df_base['Is_Valid'] = df_base['Vel_Num'].notna()
            df_base['Match_Label'] = df_base['Data_Solo'].astype(str) + " - vs " + df_base['Avv.']

            df_perugia = df_base[df_base['Team'].astype(str).str.upper().str.contains('PERUGIA', na=False)].copy()
            if df_perugia.empty:
                st.warning("⚠️ Nel database non risultano dati associati a Perugia.")
            else:
                giocatori = sorted([g for g in df_perugia['Player'].dropna().unique().tolist() if str(g).strip()])
                col_a, col_b = st.columns([1, 2])
                with col_a:
                    giocatore_scelto = st.selectbox("Giocatore:", options=giocatori)
                with col_b:
                    match_disponibili = (
                        df_perugia[['Data', 'Match_Label']]
                        .drop_duplicates()
                        .sort_values('Data', ascending=False)['Match_Label']
                        .tolist()
                    )
                    selected_matches = st.multiselect(
                        "Partite da includere:",
                        options=match_disponibili,
                        default=match_disponibili
                    )

                df_g = df_perugia[(df_perugia['Player'] == giocatore_scelto) & (df_perugia['Match_Label'].isin(selected_matches))].copy()
                df_g = df_g[df_g['Is_Spin']].copy()

                if df_g.empty:
                    st.info("Nessun dato disponibile per il giocatore e le partite selezionate.")
                else:
                    partite_n = df_g['Match_Label'].nunique()
                    spin_tot = int(len(df_g))
                    media_vel = float(df_g['Vel_Num'].mean()) if df_g['Vel_Num'].notna().any() else 0.0
                    pct_err = float(df_g['Is_Error'].mean() * 100) if len(df_g) else 0.0
                    pct_val = float(df_g['Is_Valid'].mean() * 100) if len(df_g) else 0.0

                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Partite", partite_n)
                    c2.metric("Spin Totali", spin_tot)
                    c3.metric("Media Km/h", f"{media_vel:.1f}")
                    c4.metric("% Errori", f"{pct_err:.1f}%")
                    st.caption(f"% Valide: {pct_val:.1f}%")

                    df_match = (
                        df_g.groupby(['Match_Label', 'Data_Solo', 'Avv.'])
                        .agg(
                            **{
                                'Spin Totali': ('Vel_Str', 'size'),
                                'Media Km/h': ('Vel_Num', 'mean'),
                                '% Errori': ('Is_Error', lambda x: x.mean() * 100),
                                '% Valide': ('Is_Valid', lambda x: x.mean() * 100),
                                'Max Km/h': ('Vel_Num', 'max')
                            }
                        )
                        .reset_index()
                        .sort_values('Data_Solo')
                    )
                    df_match['Etichetta'] = df_match['Avv.'].astype(str).str.upper() + ' | ' + pd.to_datetime(df_match['Data_Solo']).dt.strftime('%d-%m')

                    col1, col2 = st.columns(2)
                    with col1:
                        fig_vel = px.line(df_match, x='Etichetta', y='Media Km/h', markers=True, text='Media Km/h')
                        fig_vel.update_traces(texttemplate='%{text:.1f}', textposition='top center')
                        fig_vel.update_layout(title='Velocità media per partita', xaxis_title='', yaxis_title='Km/h')
                        st.plotly_chart(fig_vel, use_container_width=True)
                    with col2:
                        fig_err = px.bar(df_match, x='Etichetta', y='% Errori', text='% Errori')
                        fig_err.update_traces(texttemplate='%{text:.1f}%', textposition='outside', cliponaxis=False)
                        fig_err.update_layout(title='% errori per partita', xaxis_title='', yaxis_title='% Errori')
                        st.plotly_chart(fig_err, use_container_width=True)

                    st.write('### 📋 Dettaglio per partita')
                    player_pdf_df = uniforma_decimali(
                        df_match[['Data_Solo', 'Avv.', 'Spin Totali', 'Media Km/h', 'Max Km/h', '% Valide', '% Errori']]
                        .rename(columns={'Data_Solo': 'Data', 'Avv.': 'Avversario'})
                        .copy()
                    )
                    st.dataframe(
                        player_pdf_df,
                        use_container_width=True,
                        hide_index=True
                    )

                    player_metrics = {
                        'Spin Totali': int(spin_tot) if pd.notna(spin_tot) else '-',
                        'Media Km/h': f"{media_kmh:.1f}" if pd.notna(media_kmh) else '-',
                        'Max Km/h': f"{max_kmh:.1f}" if pd.notna(max_kmh) else '-',
                        '% Valide': f"{perc_valide:.1f}" if pd.notna(perc_valide) else '-',
                        '% Errori': f"{perc_errori:.1f}" if pd.notna(perc_errori) else '-',
                        'Serve Impact Index': f"{player_index:.1f}" if pd.notna(player_index) else '-',
                    }
                    pdf_player = build_player_sheet_pdf(
                        player_sel,
                        len(df_match),
                        player_metrics,
                        player_pdf_df
                    )
                    st.download_button(
                        "📄 Scarica PDF Scheda Battitore",
                        data=pdf_player,
                        file_name=f"scheda_battitore_{player_sel.lower().replace(' ', '_')}.pdf",
                        mime="application/pdf",
                        key=f"pdf_scheda_battitore_{player_sel}"
                    )

elif scelta == "Confronto Partite":
    st.markdown("<h2 style='text-align: center;'>🆚 CONFRONTO PARTITE</h2>", unsafe_allow_html=True)

    df_base = st.session_state['df_master'].copy() if ('df_master' in st.session_state and not st.session_state['df_master'].empty) else df_master.copy()

    if df_base.empty:
        st.warning("⚠️ Carica prima il database per visualizzare il Confronto Partite.")
    else:
        colonne_minime = ['Data', 'Avv.', 'Team', 'Vel.']
        colonne_mancanti = [c for c in colonne_minime if c not in df_base.columns]
        if colonne_mancanti:
            st.error(f"❌ Mancano queste colonne nel database: {colonne_mancanti}")
        else:
            df_base = df_base.copy()
            df_base['Data_raw'] = df_base['Data'].astype(str).str.strip()
            df_base['Data'] = parse_match_dates(df_base['Data_raw'])
            df_base = df_base.dropna(subset=['Data']).copy()
            df_base['Data_Solo'] = df_base['Data'].dt.date
            df_base['Avv.'] = df_base['Avv.'].astype(str).str.strip()
            df_base['Team'] = df_base['Team'].astype(str).str.strip()
            df_base['Vel_Str'] = df_base['Vel.'].astype(str).str.upper().str.strip()
            df_base['Vel_Num'] = df_base['Vel.'].apply(clean_vel_val)
            df_base['Is_Spin'] = df_base['Vel_Num'].notna() | df_base['Vel_Str'].isin(['N', 'F'])
            df_base['Is_Error'] = df_base['Vel_Str'].isin(['N', 'F'])
            df_base['Match_Label'] = df_base['Data_Solo'].astype(str) + " - vs " + df_base['Avv.']
            df_base['Side'] = df_base['Team'].apply(lambda x: 'Perugia' if 'PERUGIA' in str(x).upper() else 'Avversario')

            match_list = (
                df_base[['Data', 'Match_Label']]
                .drop_duplicates()
                .sort_values('Data', ascending=False)['Match_Label']
                .tolist()
            )
            default_matches = match_list[:5] if len(match_list) >= 5 else match_list
            selected_matches = st.multiselect("Seleziona le partite:", options=match_list, default=default_matches)

            if not selected_matches:
                st.info("Seleziona almeno una partita.")
            else:
                df_sel = df_base[df_base['Match_Label'].isin(selected_matches) & df_base['Is_Spin']].copy()
                if df_sel.empty:
                    st.info("Nessun dato disponibile per le partite selezionate.")
                else:
                    df_cmp = (
                        df_sel.groupby(['Match_Label', 'Data_Solo', 'Avv.', 'Side'])
                        .agg(
                            **{
                                'Spin Totali': ('Vel_Str', 'size'),
                                'Media Km/h': ('Vel_Num', 'mean'),
                                '% Errori': ('Is_Error', lambda x: x.mean() * 100)
                            }
                        )
                        .reset_index()
                        .sort_values(['Data_Solo', 'Side'])
                    )
                    df_cmp['Etichetta'] = df_cmp['Avv.'].astype(str).str.upper() + ' | ' + pd.to_datetime(df_cmp['Data_Solo']).dt.strftime('%d-%m')

                    c1, c2, c3 = st.columns(3)
                    c1.metric('Partite selezionate', len(selected_matches))
                    c2.metric('Media Km/h Perugia', f"{df_cmp[df_cmp['Side']=='Perugia']['Media Km/h'].mean():.1f}" if not df_cmp[df_cmp['Side']=='Perugia'].empty else '0.0')
                    c3.metric('Media % Errori Perugia', f"{df_cmp[df_cmp['Side']=='Perugia']['% Errori'].mean():.1f}%" if not df_cmp[df_cmp['Side']=='Perugia'].empty else '0.0%')

                    col1, col2 = st.columns(2)
                    with col1:
                        fig_vel = px.bar(df_cmp, x='Etichetta', y='Media Km/h', color='Side', barmode='group', text='Media Km/h')
                        fig_vel.update_traces(texttemplate='%{text:.1f}', textposition='outside', cliponaxis=False)
                        fig_vel.update_layout(title='Velocità media confronto partite', xaxis_title='', yaxis_title='Km/h')
                        st.plotly_chart(fig_vel, use_container_width=True)
                    with col2:
                        fig_err = px.bar(df_cmp, x='Etichetta', y='% Errori', color='Side', barmode='group', text='% Errori')
                        fig_err.update_traces(texttemplate='%{text:.1f}%', textposition='outside', cliponaxis=False)
                        fig_err.update_layout(title='% errori confronto partite', xaxis_title='', yaxis_title='% Errori')
                        st.plotly_chart(fig_err, use_container_width=True)

                    st.write('### 📋 Tabella comparativa')
                    st.dataframe(
                        uniforma_decimali(df_cmp[['Data_Solo', 'Avv.', 'Side', 'Spin Totali', 'Media Km/h', '% Errori']].rename(
                            columns={'Data_Solo': 'Data', 'Avv.': 'Avversario', 'Side': 'Squadra'}
                        )),
                        use_container_width=True,
                        hide_index=True
                    )

elif scelta == "Ranking Battitori":
    st.markdown("<h2 style='text-align: center;'>🏅 RANKING BATTITORI</h2>", unsafe_allow_html=True)

    df_base = st.session_state['df_master'].copy() if ('df_master' in st.session_state and not st.session_state['df_master'].empty) else df_master.copy()

    if df_base.empty:
        st.warning("⚠️ Carica prima il database per visualizzare il Ranking Battitori.")
    else:
        colonne_minime = ['Data', 'Team', 'Player', 'Vel.']
        colonne_mancanti = [c for c in colonne_minime if c not in df_base.columns]
        if colonne_mancanti:
            st.error(f"❌ Mancano queste colonne nel database: {colonne_mancanti}")
        else:
            df_base = df_base.copy()
            df_base['Data_raw'] = df_base['Data'].astype(str).str.strip()
            df_base['Data'] = parse_match_dates(df_base['Data_raw'])
            df_base = df_base.dropna(subset=['Data']).copy()
            df_base['Team'] = df_base['Team'].astype(str).str.strip()
            df_base['Player'] = df_base['Player'].astype(str).str.strip()
            df_base['Vel_Str'] = df_base['Vel.'].astype(str).str.upper().str.strip()
            df_base['Vel_Num'] = df_base['Vel.'].apply(clean_vel_val)
            df_base['Is_Spin'] = df_base['Vel_Num'].notna() | df_base['Vel_Str'].isin(['N', 'F'])
            df_base['Is_Error'] = df_base['Vel_Str'].isin(['N', 'F'])
            df_base['Is_Valid'] = df_base['Vel_Num'].notna()
            df_p = df_base[df_base['Team'].astype(str).str.upper().str.contains('PERUGIA', na=False) & df_base['Is_Spin']].copy()

            if df_p.empty:
                st.info("Nessun dato utile per costruire il ranking.")
            else:
                min_spin = st.slider('Minimo spin per entrare nel ranking:', min_value=5, max_value=100, value=15, step=5)
                df_rank = (
                    df_p.groupby('Player')
                    .agg(
                        **{
                            'Spin Totali': ('Vel_Str', 'size'),
                            'Media Km/h': ('Vel_Num', 'mean'),
                            '% Errori': ('Is_Error', lambda x: x.mean() * 100),
                            '% Valide': ('Is_Valid', lambda x: x.mean() * 100),
                            'Max Km/h': ('Vel_Num', 'max')
                        }
                    )
                    .reset_index()
                )
                df_rank = df_rank[df_rank['Spin Totali'] >= min_spin].copy()
                if df_rank.empty:
                    st.info('Nessun giocatore supera la soglia minima selezionata.')
                else:
                    df_rank['Serve Impact Index'] = df_rank['Media Km/h'].fillna(0) * (df_rank['% Valide'].fillna(0) / 100) * (1 - df_rank['% Errori'].fillna(0) / 100)
                    df_rank = uniforma_decimali(df_rank)
                    df_rank = df_rank.sort_values('Serve Impact Index', ascending=False).reset_index(drop=True)
                    df_rank.insert(0, 'Rank', range(1, len(df_rank) + 1))

                    st.caption('Indice sintetico = Media Km/h × % Valide × (1 - % Errori). Nessun riferimento ai finali: qui guardiamo solo il rendimento complessivo.')
                    st.dataframe(uniforma_decimali(df_rank), use_container_width=True, hide_index=True)

                    fig_rank = px.bar(df_rank.head(10), x='Serve Impact Index', y='Player', orientation='h', text='Serve Impact Index')
                    fig_rank.update_traces(texttemplate='%{text:.1f}', textposition='outside', cliponaxis=False)
                    fig_rank.update_layout(title='Top ranking battitori', yaxis={'categoryorder': 'total ascending'})
                    st.plotly_chart(fig_rank, use_container_width=True)

elif scelta == "Insight":
    st.markdown("<h2 style='text-align: center;'>💡 INSIGHT AUTOMATICI</h2>", unsafe_allow_html=True)

    df_base = st.session_state['df_master'].copy() if ('df_master' in st.session_state and not st.session_state['df_master'].empty) else df_master.copy()

    if df_base.empty:
        st.warning("⚠️ Carica prima il database per visualizzare gli Insight.")
    else:
        colonne_minime = ['Data', 'Avv.', 'Team', 'Player', 'Vel.']
        colonne_mancanti = [c for c in colonne_minime if c not in df_base.columns]
        if colonne_mancanti:
            st.error(f"❌ Mancano queste colonne nel database: {colonne_mancanti}")
        else:
            df_base = df_base.copy()
            df_base['Data_raw'] = df_base['Data'].astype(str).str.strip()
            df_base['Data'] = parse_match_dates(df_base['Data_raw'])
            df_base = df_base.dropna(subset=['Data']).copy()
            df_base['Data_Solo'] = df_base['Data'].dt.date
            df_base['Avv.'] = df_base['Avv.'].astype(str).str.strip()
            df_base['Team'] = df_base['Team'].astype(str).str.strip()
            df_base['Player'] = df_base['Player'].astype(str).str.strip()
            df_base['Vel_Str'] = df_base['Vel.'].astype(str).str.upper().str.strip()
            df_base['Vel_Num'] = df_base['Vel.'].apply(clean_vel_val)
            df_base['Is_Spin'] = df_base['Vel_Num'].notna() | df_base['Vel_Str'].isin(['N', 'F'])
            df_base['Is_Error'] = df_base['Vel_Str'].isin(['N', 'F'])
            df_base['Is_Valid'] = df_base['Vel_Num'].notna()
            df_base['Match_Label'] = df_base['Data_Solo'].astype(str) + " - vs " + df_base['Avv.']
            df_p = df_base[df_base['Team'].astype(str).str.upper().str.contains('PERUGIA', na=False) & df_base['Is_Spin']].copy()

            if df_p.empty:
                st.info('Nessun dato utile per generare insight.')
            else:
                df_match = (
                    df_p.groupby(['Match_Label', 'Data_Solo', 'Avv.'])
                    .agg(
                        **{
                            'Spin Totali': ('Vel_Str', 'size'),
                            'Media Km/h': ('Vel_Num', 'mean'),
                            '% Errori': ('Is_Error', lambda x: x.mean() * 100)
                        }
                    )
                    .reset_index()
                    .sort_values('Data_Solo')
                )
                df_player = (
                    df_p.groupby('Player')
                    .agg(
                        **{
                            'Spin Totali': ('Vel_Str', 'size'),
                            'Media Km/h': ('Vel_Num', 'mean'),
                            '% Errori': ('Is_Error', lambda x: x.mean() * 100),
                            '% Valide': ('Is_Valid', lambda x: x.mean() * 100)
                        }
                    )
                    .reset_index()
                )
                df_player_min = df_player[df_player['Spin Totali'] >= 15].copy()
                insights = []
                if not df_match.empty:
                    best_speed = df_match.loc[df_match['Media Km/h'].idxmax()]
                    worst_err = df_match.loc[df_match['% Errori'].idxmax()]
                    insights.append(f"La partita con velocità media più alta è stata contro {best_speed['Avv.']} il {pd.to_datetime(best_speed['Data_Solo']).strftime('%d/%m/%Y')} con {best_speed['Media Km/h']:.1f} km/h.")
                    insights.append(f"La partita con la percentuale errori più alta è stata contro {worst_err['Avv.']} il {pd.to_datetime(worst_err['Data_Solo']).strftime('%d/%m/%Y')} con {worst_err['% Errori']:.1f}%.")
                    if len(df_match) >= 6:
                        recent = df_match.tail(3)
                        prev = df_match.iloc[-6:-3]
                        if not prev.empty:
                            diff_vel = recent['Media Km/h'].mean() - prev['Media Km/h'].mean()
                            diff_err = recent['% Errori'].mean() - prev['% Errori'].mean()
                            verso_vel = 'in crescita' if diff_vel > 0 else 'in calo'
                            verso_err = 'in crescita' if diff_err > 0 else 'in calo'
                            insights.append(f"Nelle ultime 3 partite la velocità media è {verso_vel} di {abs(diff_vel):.1f} km/h rispetto alle 3 precedenti.")
                            insights.append(f"Nelle ultime 3 partite la percentuale errori è {verso_err} di {abs(diff_err):.1f} punti rispetto alle 3 precedenti.")
                if not df_player_min.empty:
                    best_valid = df_player_min.loc[df_player_min['% Valide'].idxmax()]
                    best_speed_player = df_player_min.loc[df_player_min['Media Km/h'].idxmax()]
                    insights.append(f"Il battitore più affidabile, sopra la soglia minima, è {best_valid['Player']} con {best_valid['% Valide']:.1f}% di battute valide.")
                    insights.append(f"Il battitore con la velocità media più alta, sopra la soglia minima, è {best_speed_player['Player']} con {best_speed_player['Media Km/h']:.1f} km/h.")

                if insights:
                    for i, txt in enumerate(insights, 1):
                        st.markdown(f"**{i}.** {txt}")
                else:
                    st.info('Dati ancora insufficienti per produrre insight utili.')

                if not df_match.empty:
                    fig_ins = px.line(df_match, x='Data_Solo', y='Media Km/h', markers=True, text='Media Km/h')
                    fig_ins.update_traces(texttemplate='%{text:.1f}', textposition='top center')
                    fig_ins.update_layout(title='Velocità media Perugia nel tempo', xaxis_title='', yaxis_title='Km/h')
                    st.plotly_chart(fig_ins, use_container_width=True)

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
            df_base['Data'] = parse_match_dates(df_base['Data_raw'])
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

