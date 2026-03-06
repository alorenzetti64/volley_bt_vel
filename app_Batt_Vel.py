import streamlit as st
import pandas as pd
import plotly.express as px
from github import Github, Auth
import io

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

if scelta == "Caricamento Dati":
    st.title("🚀 Database")
    tab1, tab2 = st.tabs(["Carica Excel", "Pulisci Database"])
    with tab1:
        uploaded = st.file_uploader("Seleziona file .xlsm", type=["xlsm"])
        if uploaded:
            df_new = pd.read_excel(uploaded, sheet_name="Foglio1").iloc[:, 0:8]
            df_new.columns = COLUMNS_A_H
            if st.button("🚀 Sincronizza su GitHub"):
                # Concateniamo e salviamo con la nuova logica string-only
                df_combined = pd.concat([df_master, df_new]).drop_duplicates()
                save_to_github(df_combined)
                st.balloons()
                st.success("Dati sincronizzati!")
    with tab2:
        if not df_master.empty:
            date_list = sorted([str(x) for x in df_master['Data'].dropna().unique() if str(x) != ''])
            d_to_del = st.selectbox("Seleziona Data da eliminare:", date_list)
            if st.button("🔥 Elimina Record"):
                save_to_github(df_master[~(df_master['Data'].astype(str) == d_to_del)])
                st.rerun()

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
                st.table(pd.DataFrame(r_p, columns=cols_h).style.hide(axis="index").apply(stile_righe, axis=1).format({"Media Km/h": "{:.1f}"}, precision=1))

                st.markdown(f"### 🏐 {nome_avv}")
                df_o = df_report[df_report['Team'].astype(str).str.upper() != 'PERUGIA'].copy()
                if not df_o.empty:
                    r_o = [["MATCH"] + calcola_stats(df_o)]
                    for s in sorted(df_o['Set'].unique()): r_o.append([f"Set {int(float(s))}"] + calcola_stats(df_o[df_o['Set'] == s]))
                    st.table(pd.DataFrame(r_o, columns=cols_h).style.hide(axis="index").apply(stile_righe, axis=1).format({"Media Km/h": "{:.1f}"}, precision=1))

                st.divider()
                st.markdown("## 🏐 PERUGIA - PLAYER")
                cp1, cp2 = st.columns([1, 2])
                with cp1: tp_p = st.radio("Modalità Perugia:", ["GENERALE", "INDIVIDUALE"], horizontal=True, key="tp_p")
                p_list = sorted([str(x) for x in df_p[df_p['Vel_Num'].notna() | df_p['Tipo'].isin(['V','N','F','v','n','f'])]['Player'].unique()])
                if tp_p == "GENERALE":
                    with cp2: fs_p = st.selectbox("Fase Perugia:", ["MATCH"] + [f"Set {int(float(s))}" for s in sorted(df_p['Set'].unique())], key="fs_p")
                    rg_p = []
                    for p in p_list:
                        df_pf = df_p[df_p['Player'] == p] if fs_p == "MATCH" else df_p[(df_p['Player'] == p) & (df_p['Set'].astype(float) == float(fs_p.replace("Set ","")))]
                        if not df_pf.empty: rg_p.append([p] + calcola_stats(df_pf))
                    st.table(pd.DataFrame(rg_p, columns=["Player"] + cols_h[1:]).style.hide(axis="index").format({"Media Km/h": "{:.1f}"}, precision=1))
                else:
                    with cp2: ps_p = st.selectbox("Giocatore Perugia:", p_list, key="ps_p")
                    df_i_p = df_p[df_p['Player'] == ps_p]
                    ri_p = [["MATCH"] + calcola_stats(df_i_p)]
                    for s in sorted(df_i_p['Set'].unique()): ri_p.append([f"Set {int(float(s))}"] + calcola_stats(df_i_p[df_i_p['Set'] == s]))
                    st.table(pd.DataFrame(ri_p, columns=cols_h).style.hide(axis="index").apply(stile_righe, axis=1).format({"Media Km/h": "{:.1f}"}, precision=1))

                st.divider()
                st.markdown(f"## 🏐 {nome_avv} - PLAYER")
                ca1, ca2 = st.columns([1, 2])
                with ca1: tp_a = st.radio(f"Modalità {nome_avv}:", ["GENERALE", "INDIVIDUALE"], horizontal=True, key="tp_a")
                a_list = sorted([str(x) for x in df_o[df_o['Vel_Num'].notna() | df_o['Tipo'].isin(['V','N','F','v','n','f'])]['Player'].unique()])
                if tp_a == "GENERALE":
                    with ca2: fs_a = st.selectbox(f"Fase {nome_avv}:", ["MATCH"] + [f"Set {int(float(s))}" for s in sorted(df_o['Set'].unique())], key="fs_a")
                    rg_a = []
                    for a in a_list:
                        df_af = df_o[df_o['Player'] == a] if fs_a == "MATCH" else df_o[(df_o['Player'] == a) & (df_o['Set'].astype(float) == float(fs_a.replace("Set ","")))]
                        if not df_af.empty: rg_a.append([a] + calcola_stats(df_af))
                    st.table(pd.DataFrame(rg_a, columns=["Player"] + cols_h[1:]).style.hide(axis="index").format({"Media Km/h": "{:.1f}"}, precision=1))
                else:
                    with ca2: ps_a = st.selectbox(f"Giocatore {nome_avv}:", a_list, key="ps_a")
                    df_i_a = df_o[df_o['Player'] == ps_a]
                    ri_a = [["MATCH"] + calcola_stats(df_i_a)]
                    for s in sorted(df_i_a['Set'].unique()): ri_a.append([f"Set {int(float(s))}"] + calcola_stats(df_i_a[df_i_a['Set'] == s]))
                    st.table(pd.DataFrame(ri_a, columns=cols_h).style.hide(axis="index").apply(stile_righe, axis=1).format({"Media Km/h": "{:.1f}"}, precision=1))

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