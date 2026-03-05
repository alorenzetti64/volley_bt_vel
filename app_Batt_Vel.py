import streamlit as st
import pandas as pd
from github import Github
import io

# --- CONFIGURAZIONE ---
COLUMNS_A_H = ['Data', 'Partita', 'Avv.', 'Team', 'Set', 'Player', 'Tipo', 'Vel.']

def check_vel(x):
    s = str(x).strip().upper()
    if s in ["", "NAN", "NONE", "0", "0.0"]:
        return True 
    if s in ['N', 'F', 'V']: 
        return True
    try:
        val = float(s.replace(',', '.'))
        return 30 <= val <= 150
    except: 
        return False

def validate_data(df):
    mask_invalid = ~df['Vel.'].apply(check_vel)
    if mask_invalid.any():
        rows = df[mask_invalid].index.tolist()
        return False, f"Valori non validi in 'Vel.' alle righe: {rows}"
    return True, "Validazione superata!"

# --- FUNZIONI GITHUB ---
def get_github_client():
    token = st.secrets["github"]["access_token"]
    return Github(token)

def load_master_from_github():
    try:
        g = get_github_client()
        repo = g.get_repo(st.secrets["github"]["repository"])
        contents = repo.get_contents(st.secrets["github"]["file_path"])
        return pd.read_parquet(io.BytesIO(contents.decoded_content))
    except:
        return pd.DataFrame(columns=COLUMNS_A_H)

def save_to_github(df):
    g = get_github_client()
    repo = g.get_repo(st.secrets["github"]["repository"])
    path = st.secrets["github"]["file_path"]
    buffer = io.BytesIO()
    df.to_parquet(buffer, index=False)
    content = buffer.getvalue()
    try:
        contents = repo.get_contents(path)
        repo.update_file(path, "Update master battute", content, contents.sha)
    except:
        repo.create_file(path, "Initial master battute", content)

# --- INTERFACCIA STREAMLIT ---
st.set_page_config(page_title="Sir Stats", layout="wide")
st.title("🏐 Analisi Battute Sir Susa Vim Perugia")

uploaded_file = st.file_uploader("Carica file partita (.xlsm)", type=["xlsm"])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, sheet_name="Foglio1")
    df_new = df_raw.iloc[:, 0:8] 
    df_new.columns = COLUMNS_A_H
    df_new['Vel.'] = df_new['Vel.'].astype(str)
    df_new = df_new.dropna(subset=['Player', 'Data'], how='all')

    is_valid, msg = validate_data(df_new)
    
    if not is_valid:
        st.error(msg)
    else:
        st.success(msg)
        if st.button("🚀 Carica e Aggiorna Master su GitHub"):
            with st.spinner("Sincronizzazione..."):
                master_df = load_master_from_github()
                updated_df = pd.concat([master_df, df_new]).drop_duplicates()
                save_to_github(updated_df)
                st.balloons()
                st.success("Dati inviati!")

st.divider()

# --- SEZIONE STATISTICHE FILTRATE PER PERUGIA ---
full_data = load_master_from_github()

if not full_data.empty:
    # FILTRO: Prendiamo solo le righe dove il Team è Perugia
    # (Uso .str.contains per sicurezza nel caso ci siano spazi o maiuscole diverse)
    perugia_data = full_data[full_data['Team'].astype(str).str.contains('Perugia', case=False, na=False)].copy()
    
    st.subheader("📊 Statistiche Esclusive: Sir Susa Vim Perugia")
    
    if not perugia_data.empty:
        # Conversione velocità per i calcoli
        perugia_data['Vel_Num'] = pd.to_numeric(perugia_data['Vel.'].str.replace(',', '.'), errors='coerce')
        
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("Record Velocità", f"{perugia_data['Vel_Num'].max():.1f} km/h")
        with c2:
            st.metric("Battute Totali", len(perugia_data))
        with c3:
            aces = len(perugia_data[perugia_data['Vel.'].str.upper() == 'V'])
            st.metric("Ace (V)", aces)
        with c4:
            media_sq = perugia_data['Vel_Num'].mean()
            st.metric("Media Squadra", f"{media_sq:.1f} km/h")

        # Grafico Medie Giocatori Perugia
        st.write("### 🚀 Classifica Potenza Giocatori (Media km/h)")
        player_avg = perugia_data.dropna(subset=['Vel_Num']).groupby('Player')['Vel_Num'].mean().sort_values(ascending=False)
        st.bar_chart(player_avg)
        
        # Tabella degli esiti (V, N, F) per Perugia
        st.write("### 🎯 Riepilogo Esiti Battuta")
        esiti = perugia_data['Vel.'].str.upper().value_counts()
        # Filtriamo solo le lettere per chiarezza
        lettere = esiti[esiti.index.isin(['V', 'N', 'F'])]
        if not lettere.empty:
            st.table(lettere.rename("Conteggio"))
    else:
        st.warning("Nessun dato trovato per 'Perugia' nella colonna Team.")

    with st.expander("Visualizza tutto il Database (Inclusi Avversari)"):
        st.dataframe(full_data)
else:
    st.info("Database vuoto. Carica i file per iniziare.")