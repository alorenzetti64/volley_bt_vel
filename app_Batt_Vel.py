
import re
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
import streamlit as st

import hmac

# =============================================================
# STAFF LOCK (import pages only for Angelo)
# =============================================================
STAFF_PASSWORD = "Agl64@Volley"

def staff_unlocked() -> bool:
    """Return True if staff area is unlocked via password."""
    if "staff_ok" not in st.session_state:
        st.session_state.staff_ok = False

    if st.session_state.staff_ok:
        return True

    with st.sidebar.expander("Area riservata", expanded=False):
        pw = st.text_input("Password", type="password", key="staff_pw")
        if st.button("Sblocca", key="staff_unlock_btn"):
            st.session_state.staff_ok = hmac.compare_digest(str(pw), STAFF_PASSWORD)
            if not st.session_state.staff_ok:
                st.error("Password errata.")

    return st.session_state.staff_ok

from sqlalchemy import create_engine, text

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Volley App", layout="wide")

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Volley App", layout="wide")

# =========================
# DB CONFIG (Supabase Postgres)
# =========================
DB_URL = st.secrets.get("DATABASE_URL", "").strip()
if not DB_URL:
    st.error("DATABASE_URL mancante in Secrets (Streamlit Cloud).")
    st.stop()

engine = create_engine(
    DB_URL,
    future=True,
    pool_pre_ping=True,
    pool_size=5,
    max_overflow=2,
    connect_args={
        "connect_timeout": 8,
        "sslmode": "require",
    },
)
# =========================
# HELPERS
# =========================
def norm(s: str | None) -> str:
    if not s:
        return ""
    s = s.strip().lower()
    s = " ".join(s.split())
    return s

# =============================================================
# TABLE RENDERER (auto width like Excel-ish) + Perugia highlight
# =============================================================
import html as _html

def render_table_auto(
    df: pd.DataFrame,
    fmt: dict | None = None,
    team_col: str | None = None,
    perugia_word: str = "perugia",
    highlight_cols: list[str] | None = None,
    max_rows: int | None = None,
):
    """Render a DataFrame as HTML table with column widths based on content.
    - Auto column widths (character-based)
    - 1 decimal for floats by default
    - Highlight Perugia row (if team column found)
    - Highlight selected columns (e.g., % columns or guide column)
    """
    if df is None or getattr(df, "empty", True):
        st.info("Nessun dato.")
        return

    df = df.copy()
    if max_rows is not None:
        df = df.head(int(max_rows)).copy()

    fmt = fmt or {}
    highlight_cols = highlight_cols or []

    # pick team column automatically if not provided
    if team_col is None:
        for cand in ["squadra", "Team", "Squadra", "Nome Team"]:
            if cand in df.columns:
                team_col = cand
                break

    cols = list(df.columns)

    # format cells to strings
    def _format_cell(col, v):
        f = fmt.get(col)
        if f is not None:
            try:
                return f.format(v)
            except Exception:
                return "" if pd.isna(v) else str(v)
        # default formats
        if isinstance(v, (int,)) and not isinstance(v, bool):
            return f"{v:d}"
        if isinstance(v, (float,)):
            return f"{v:.1f}"
        return "" if pd.isna(v) else str(v)

    str_rows = []
    for _, row in df.iterrows():
        str_rows.append([_format_cell(c, row[c]) for c in cols])

    # compute widths (in ch)
    sample_n = min(40, len(str_rows))
    widths = []
    for j, c in enumerate(cols):
        header_len = len(str(c))
        cell_lens = [len(str_rows[i][j]) for i in range(sample_n)]
        mx = max([header_len] + cell_lens) if cell_lens else header_len
        mx = max(6, min(28, mx))  # clamp
        widths.append(mx)

    team_idx = cols.index(team_col) if (team_col in cols) else None
    hi_idx = {cols.index(c) for c in highlight_cols if c in cols}

    # if no explicit highlight cols, highlight all % columns
    if not hi_idx:
        hi_idx = {i for i,c in enumerate(cols) if "%" in str(c)}

    css = """
    <style>
    .tblwrap { overflow-x:auto; }
    table.autotbl { border-collapse: collapse; width: max-content; min-width: 100%; }
    table.autotbl th, table.autotbl td { border: 1px solid #e9ecef; padding: 8px 10px; white-space: nowrap; }
    table.autotbl th { position: sticky; top: 0; background: #f8f9fa; font-weight: 800; text-align: left; }
    table.autotbl td.num { text-align: right; }
    table.autotbl tr.perugia td { background: #fff3cd; font-weight: 800; }
    table.autotbl td.hicol { background: #e7f5ff; font-weight: 900; }
    </style>
    """

    # header
    ths = []
    for c,w in zip(cols, widths):
        ths.append(f'<th style="min-width:{w}ch">{_html.escape(str(c))}</th>')
    html_rows = ["<tr>" + "".join(ths) + "</tr>"]

    # body
    for r in str_rows:
        is_perugia = False
        if team_idx is not None:
            is_perugia = perugia_word in (r[team_idx] or "").lower()

        tds = []
        for j,(val,w) in enumerate(zip(r, widths)):
            classes = []
            # numeric align heuristic
            sval = val or ""
            if j != team_idx and sval and sval.replace(".", "", 1).replace("-", "", 1).isdigit():
                classes.append("num")
            if j in hi_idx:
                classes.append("hicol")
            cls = " ".join(classes).strip()
            cls_attr = f' class="{cls}"' if cls else ""
            tds.append(f'<td{cls_attr} style="min-width:{w}ch">{_html.escape(sval)}</td>')

        tr_cls = ' class="perugia"' if is_perugia else ""
        html_rows.append(f"<tr{tr_cls}>" + "".join(tds) + "</tr>")

    st.markdown(css + '<div class="tblwrap"><table class="autotbl">' + "".join(html_rows) + "</table></div>", unsafe_allow_html=True)


def render_any_table(obj, fmt: dict | None = None, team_col: str | None = None, highlight_cols: list[str] | None = None, max_rows: int | None = None):
    """Accept a DataFrame OR a pandas Styler and render with auto-width HTML."""
    df = None
    try:
        # pandas Styler has .data
        df = obj.data  # type: ignore
    except Exception:
        df = obj
    render_table_auto(df, fmt=fmt, team_col=team_col, highlight_cols=highlight_cols, max_rows=max_rows)


# =========================
# TEAM CANON (sponsor/variants -> 12 teams)
# =========================

def render_focus_4_players(out: pd.DataFrame, key_base: str, name_col: str = "Nome giocatore"):
    """Render a 'Focus 4 players' selector + mini table, keeping ranking from the main table."""
    if out is None or getattr(out, "empty", True):
        return
    # Ranking column can be 'Ranking' or 'Rank'
    rank_col = "Ranking" if "Ranking" in out.columns else ("Rank" if "Rank" in out.columns else None)
    if rank_col is None or name_col not in out.columns:
        return

    st.subheader("Focus 4 giocatori")
    names = out[name_col].dropna().astype(str).tolist()
    if not names:
        return

    default = names[:4] if len(names) >= 4 else names
    picked = st.multiselect(
        "Scegli fino a 4 giocatori (dal filtro attuale)",
        options=names,
        default=default,
        max_selections=4,
        key=f"focus_{key_base}",
    )

    if picked:
        focus_df = out[out[name_col].isin(picked)].copy()
        focus_df = focus_df.sort_values(by=rank_col, ascending=True)
        render_any_table(focus_df, highlight_cols=[rank_col], max_rows=4)
    st.divider()

def canonical_team(name: str | None) -> str:
    """Map raw team names (with sponsors/variants) to canonical labels."""
    if not name:
        return ""
    s = str(name).strip().lower()
    s = " ".join(s.split())
    # substring rules (broad, to catch sponsor variants)
    rules = [
        ("PERUGIA",    ["perugia", "sir", "susa"]),
        ("MODENA",     ["modena", "valsa"]),
        ("TRENTO",     ["trentino", "trento", "itas"]),
        ("MILANO",     ["milano", "allianz"]),
        ("PIACENZA",   ["piacenza", "gas sales", "bluenergy"]),
        ("CIVITANOVA", ["civitanova", "lube", "cucine lube"]),
        ("VERONA",     ["verona", "rana"]),
        ("CUNEO",      ["cuneo", "bernardo", "s. bernardo", "s bernardo", "ma acqua"]),
        ("CISTERNA",   ["cisterna"]),
        ("PADOVA",     ["padova", "sonepar"]),
        ("MONZA",      ["monza", "vero volley"]),
        ("GROTTA",     ["grotta", "grottazzolina", "yuasa"]),
    ]
    for canon, needles in rules:
        if any(n in s for n in needles):
            return canon
    # fallback: keep readable
    return str(name).strip().upper()

def build_match_key(team_a: str, team_b: str, competition: str | None, phase: str, round_number: int) -> str:
    return f"{norm(team_a)}|{norm(team_b)}|{norm(competition)}|{phase}{round_number:02d}"


def extract_round_code(filename: str) -> tuple[str, int]:
    """Estrae (phase, round_number) dal nome file.

    Supporta:
      - Andata:  A01 .. A11
      - Ritorno: R01 .. R11
      - Playoff: POQ1..POQ5, POS1..POS5, POF1..POF5
    """
    fname = filename.strip()

    # 1) Pattern playoff espliciti (prioritari)
    m = re.search(r"_(POQ|POS|POF)(\d{1,2})_", fname)
    if m:
        phase = m.group(1)
        rnd = int(m.group(2))
        return phase, rnd

    # 2) Pattern Andata/Ritorno classici
    m = re.search(r"_([AR])(\d{2})_", fname)
    if m:
        return m.group(1), int(m.group(2))

    # 3) Fallback: cerca comunque PO* o ARxx dentro la stringa
    m = re.search(r"(POQ|POS|POF)(\d{1,2})", fname)
    if m:
        return m.group(1), int(m.group(2))

    m = re.search(r"([AR])(\d{2})", fname)
    if m:
        return m.group(1), int(m.group(2))

    raise ValueError("Impossibile estrarre il codice giornata dal filename")

# =========================
# MATCH FILTER (fasi + range)
# =========================
PHASE_ORDER = {"A": 1, "R": 2, "POQ": 3, "POS": 4, "POF": 5}


def match_order_value(phase: str, round_number: int) -> int:
    return (PHASE_ORDER.get(phase, 99) * 100) + int(round_number or 0)


def match_order_sql(alias: str = "m") -> str:
    return f"""(
        (CASE {alias}.phase
            WHEN 'A' THEN 1
            WHEN 'R' THEN 2
            WHEN 'POQ' THEN 3
            WHEN 'POS' THEN 4
            WHEN 'POF' THEN 5
            ELSE 99
        END) * 100 + COALESCE({alias}.round_number, 0)
    )"""


def parse_match_date(filename: str) -> str:
    m = re.match(r"^(\d{4})[\.\-](\d{2})[\.\-](\d{2})", filename.strip())
    if not m:
        return ""
    return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"


def load_matches_index() -> pd.DataFrame:
    with engine.begin() as conn:
        rows = conn.execute(text("""
            SELECT id, filename, team_a, team_b, competition, phase, round_number, match_key
            FROM matches
            WHERE phase IS NOT NULL AND round_number IS NOT NULL
        """)).mappings().all()

    if not rows:
        return pd.DataFrame(columns=["id", "filename", "team_a", "team_b", "competition", "phase", "round_number", "match_key", "order", "date"])

    df = pd.DataFrame(rows)
    df["date"] = df["filename"].apply(parse_match_date)
    df["order"] = df.apply(lambda r: match_order_value(str(r["phase"]), int(r["round_number"])), axis=1)
    df = df.sort_values(["order", "date", "filename"]).reset_index(drop=True)
    return df


def sidebar_match_filters() -> tuple[int, int, str]:
    st.sidebar.subheader("Selezione Partite")
    df = load_matches_index()

    if df.empty:
        st.sidebar.info("Nessuna partita trovata nel DB (importa prima i DVW).")
        st.session_state["mf_from"] = 0
        st.session_state["mf_to"] = 99999
        st.session_state["mf_label"] = "Tutte"
        return 0, 99999, "Tutte"

    def phase_box(label: str, phase: str, max_n: int, key_prefix: str):
        checked = st.sidebar.checkbox(label, value=False, key=f"{key_prefix}_chk")

        choices = ["Tutte"] + [
            f"{phase}{i:02d}" if phase in ("A", "R") else f"{phase}{i}"
            for i in range(1, max_n + 1)
        ]

        sel_list = st.sidebar.multiselect(
            label=f"{label} (selezione)",
            options=choices,
            default=["Tutte"],
            key=f"{key_prefix}_sel",
            disabled=(not checked),
            label_visibility="collapsed",
        )

        if not checked:
            return None

        # Se "Tutte" è selezionato (o non c'è selezione), prendo tutta la fase
        if (not sel_list) or ("Tutte" in sel_list):
            rounds = list(range(1, max_n + 1))
            return (phase, rounds)

        rounds = []
        for sel in sel_list:
            if sel == "Tutte":
                continue
            if phase in ("A", "R"):
                rounds.append(int(sel[-2:]))
            else:
                rounds.append(int(re.sub(r"\D+", "", sel)))
        rounds = sorted(set(rounds))
        return (phase, rounds)

    st.sidebar.caption("Spunta una fase → scegli 1+ partite (oppure 'Tutte').")
    a_sel = phase_box("ANDATA", "A", 11, "mf_a")
    r_sel = phase_box("RITORNO", "R", 11, "mf_r")
    poq_sel = phase_box("PO-QUARTI", "POQ", 5, "mf_poq")
    pos_sel = phase_box("PO-SEMIFINALI", "POS", 5, "mf_pos")
    pof_sel = phase_box("PO-FINALI", "POF", 5, "mf_pof")

    selections = [s for s in [a_sel, r_sel, poq_sel, pos_sel, pof_sel] if s is not None]
    if not selections:
        from_ord = int(df["order"].min())
        to_ord = int(df["order"].max())
        label = f"Tutte ({from_ord}→{to_ord})"
    else:
        from_ord = min(match_order_value(ph, min(rnds)) for ph, rnds in selections if rnds)
        to_ord = max(match_order_value(ph, max(rnds)) for ph, rnds in selections if rnds)

        parts = []
        for ph, rnds in selections:
            if not rnds:
                continue
            if len(rnds) == 1:
                lo = rnds[0]
                parts.append(f"{ph}{lo:02d}" if ph in ("A", "R") else f"{ph}{lo}")
            else:
                lo, hi = min(rnds), max(rnds)
                if ph in ("A", "R"):
                    parts.append(f"{ph}{lo:02d}–{ph}{hi:02d}")
                else:
                    parts.append(f"{ph}{lo}–{ph}{hi}")
        label = " • ".join(parts)

    st.session_state["mf_from"] = from_ord
    st.session_state["mf_to"] = to_ord
    st.session_state["mf_label"] = label

    with st.sidebar.container(border=True):
        st.write("🗓️ **Range attivo**")
        st.write(label)

    return from_ord, to_ord, label


def get_selected_match_range() -> tuple[int, int, str]:
    if "mf_from" not in st.session_state or "mf_to" not in st.session_state:
        sidebar_match_filters()
    return int(st.session_state.get("mf_from", 0)), int(st.session_state.get("mf_to", 99999)), str(st.session_state.get("mf_label", ""))

    m = re.search(r"([AR]\d{2})", filename)
    if m:
        code = m.group(1)
        return code[0], int(code[1:3])

    raise ValueError("Codice giornata non trovato nel filename (atteso A01 / R06).")


def parse_dvw_minimal(dvw_text: str) -> dict:
    season = None
    competition = None
    teams: list[str] = []

    lines = dvw_text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()

        if line == "[3MATCH]":
            if i + 1 < len(lines):
                parts = lines[i + 1].split(";")
                if len(parts) >= 4:
                    season = parts[2].strip() or None
                    competition = parts[3].strip() or None
            i += 1

        if line == "[3TEAMS]":
            j = i + 1
            while j < len(lines):
                row = lines[j].strip()
                if not row or row.startswith("["):
                    break
                parts = row.split(";")
                if len(parts) >= 2:
                    teams.append(parts[1].strip())
                j += 1
            i = j
            continue

        i += 1

    team_a = teams[0] if len(teams) >= 1 else ""
    team_b = teams[1] if len(teams) >= 2 else ""

    return {"season": season, "competition": competition, "team_a": team_a, "team_b": team_b}


def extract_scout_lines(dvw_text: str) -> list[str]:
    lines = dvw_text.splitlines()
    start = None
    for idx, line in enumerate(lines):
        if line.strip() == "[3SCOUT]":
            start = idx + 1
            break
    if start is None:
        return []

    out = []
    for line in lines[start:]:
        if line.strip().startswith("["):
            break
        s = line.strip()
        if not s:
            continue

        k = 0
        while k < len(s) and ord(s[k]) < 32:
            k += 1
        s2 = s[k:].lstrip()
        if not s2:
            continue

        if s2[0] in ("*", "a"):
            out.append(s2)

    return out


def code6(line: str) -> str:
    if not line:
        return ""
    i = 0
    while i < len(line) and ord(line[i]) < 32:
        i += 1
    s = line[i:].lstrip()
    return s[:6]


def is_home_rece(c6: str) -> bool:
    return len(c6) >= 5 and c6[0] == "*" and c6[3:5] in ("RQ", "RM")


def is_away_rece(c6: str) -> bool:
    return len(c6) >= 5 and c6[0] == "a" and c6[3:5] in ("RQ", "RM")


def is_home_spin(c6: str) -> bool:
    return len(c6) >= 5 and c6[0] == "*" and c6[3:5] == "RQ"


def is_away_spin(c6: str) -> bool:
    return len(c6) >= 5 and c6[0] == "a" and c6[3:5] == "RQ"


def is_home_float(c6: str) -> bool:
    return len(c6) >= 5 and c6[0] == "*" and c6[3:5] == "RM"


def is_away_float(c6: str) -> bool:
    return len(c6) >= 5 and c6[0] == "a" and c6[3:5] == "RM"


def is_serve(c6: str) -> bool:
    # include SQ/SM (battuta)
    return len(c6) >= 5 and c6[0] in ("*", "a") and c6[3:5] in ("SQ", "SM")


def is_home_point(c6: str) -> bool:
    return c6.startswith("*p")


def is_away_point(c6: str) -> bool:
    return c6.startswith("ap")


def is_attack(c6: str, prefix: str) -> bool:
    return len(c6) >= 6 and c6[0] == prefix and c6[3] == "A"


def first_attack_after_reception_is_winner(rally: list[str], prefix: str) -> bool:
    rece_idx = None
    for i, c in enumerate(rally):
        if len(c) >= 5 and c[0] == prefix and c[3:5] in ("RQ", "RM"):
            rece_idx = i
            break
    if rece_idx is None:
        return False

    for c in rally[rece_idx + 1 :]:
        if is_attack(c, prefix):
            return len(c) >= 6 and c[5] == "#"
    return False


# =========================
# SIDE OUT filters
# =========================
PLAYABLE_RECV = {"#", "+", "!", "-"}  # giocabili: scarta '=' e '/'
GOOD_RECV = {"#", "+"}               # buone
EXC_RECV = {"!"}                     # esclamativa
NEG_RECV = {"-"}                     # negativa

def is_home_rece_playable(c6: str) -> bool:
    return len(c6) >= 6 and c6[0] == "*" and c6[3:5] in ("RQ", "RM") and c6[5] in PLAYABLE_RECV

def is_away_rece_playable(c6: str) -> bool:
    return len(c6) >= 6 and c6[0] == "a" and c6[3:5] in ("RQ", "RM") and c6[5] in PLAYABLE_RECV

def is_home_rece_good(c6: str) -> bool:
    return len(c6) >= 6 and c6[0] == "*" and c6[3:5] in ("RQ", "RM") and c6[5] in GOOD_RECV

def is_away_rece_good(c6: str) -> bool:
    return len(c6) >= 6 and c6[0] == "a" and c6[3:5] in ("RQ", "RM") and c6[5] in GOOD_RECV

def is_home_rece_exc(c6: str) -> bool:
    return len(c6) >= 6 and c6[0] == "*" and c6[3:5] in ("RQ", "RM") and c6[5] in EXC_RECV

def is_away_rece_exc(c6: str) -> bool:
    return len(c6) >= 6 and c6[0] == "a" and c6[3:5] in ("RQ", "RM") and c6[5] in EXC_RECV

def is_home_rece_neg(c6: str) -> bool:
    return len(c6) >= 6 and c6[0] == "*" and c6[3:5] in ("RQ", "RM") and c6[5] in NEG_RECV

def is_away_rece_neg(c6: str) -> bool:
    return len(c6) >= 6 and c6[0] == "a" and c6[3:5] in ("RQ", "RM") and c6[5] in NEG_RECV


def pct(wins: int, attempts: int) -> float:
    return (wins / attempts * 100.0) if attempts else 0.0

# =========================
# ROSTER / RUOLI HELPERS
# =========================
def fix_team_name(name: str) -> str:
    """
    Normalizza nome squadra per match con roster.
    (stesse regole usate nelle pagine Break/Confronto)
    """
    n = " ".join((name or "").split())
    nl = n.lower()
    if nl.startswith("gas sales bluenergy p"):
        return "Gas Sales Bluenergy Piacenza"
    if "grottazzolina" in nl:
        return "Yuasa Battery Grottazzolina"
    return n


def team_norm(name: str) -> str:
    """Chiave stabile (lower + pulizia) per join DB."""
    n = fix_team_name(name).strip().lower()
    n = re.sub(r"[^a-z0-9\s]", " ", n)
    return " ".join(n.split())


def serve_player_number(c6: str) -> int | None:
    """
    Estrae numero maglia dal code6 della battuta:
    es: *06SQ- -> 6 ; a08SM+ -> 8
    """
    if not c6 or len(c6) < 3:
        return None
    if c6[0] not in ("*", "a"):
        return None
    digits = c6[1:3]
    if not digits.isdigit():
        return None
    return int(digits)


def serve_sign(c6: str) -> str:
    """Valutazione battuta: 6° carattere ( -, +, !, /, #, = )"""
    return c6[5] if c6 and len(c6) >= 6 else ""



def compute_counts_from_scout(scout_lines: list[str]) -> dict:
    # --- helper BT (break tendency) dal testo SERVIZIO ---
    def detect_bt(raw_line: str) -> str | None:
        if not raw_line:
            return None
        s = raw_line.strip()

        # Priorità: simboli tra parentesi quadre
        if "[-]" in s:
            return "NEG"
        if "[+]" in s:
            return "POS"
        if "[!]" in s:
            return "EXC"
        if "[½]" in s:
            return "HALF"

        # Varianti non-bracketed
        if "½" in s or "1/2" in s or "0.5" in s:
            return "HALF"

        tail = s[-6:]  # spesso il segno è vicino alla fine
        if "!" in tail:
            return "EXC"
        if "+" in tail:
            return "POS"
        if "-" in tail:
            return "NEG"

        return None

    # --- costruzione rallies: teniamo (c6, raw) per non perdere i segni BT ---
    rallies: list[list[tuple[str, str]]] = []
    current: list[tuple[str, str]] = []

    for raw in scout_lines:
        c = code6(raw)
        if not c:
            continue

        if is_serve(c):
            if current:
                rallies.append(current)
            current = [(c, raw)]
            continue

        if not current:
            continue

        current.append((c, raw))

        if is_home_point(c) or is_away_point(c):
            rallies.append(current)
            current = []

    # =========================
    # SIDE OUT counters
    # =========================
    so_home_attempts = so_home_wins = 0
    so_away_attempts = so_away_wins = 0

    bp_home_attempts = bp_home_wins = 0
    bp_away_attempts = bp_away_wins = 0

    so_spin_home_attempts = so_spin_home_wins = 0
    so_spin_away_attempts = so_spin_away_wins = 0

    so_float_home_attempts = so_float_home_wins = 0
    so_float_away_attempts = so_float_away_wins = 0

    so_dir_home_wins = 0
    so_dir_away_wins = 0

    so_play_home_attempts = so_play_home_wins = 0
    so_play_away_attempts = so_play_away_wins = 0

    so_good_home_attempts = so_good_home_wins = 0
    so_good_away_attempts = so_good_away_wins = 0

    so_exc_home_attempts = so_exc_home_wins = 0
    so_exc_away_attempts = so_exc_away_wins = 0

    so_neg_home_attempts = so_neg_home_wins = 0
    so_neg_away_attempts = so_neg_away_wins = 0

    # =========================
    # BREAK "GIOCATO" + BT counters
    # =========================
    bp_play_home_attempts = bp_play_home_wins = 0
    bp_play_away_attempts = bp_play_away_wins = 0

    bt_neg_home = bt_pos_home = bt_exc_home = bt_half_home = 0
    bt_neg_away = bt_pos_away = bt_exc_away = bt_half_away = 0

    for r in rallies:
        first_c6, first_raw = r[0]
        home_served = first_c6.startswith("*")
        away_served = first_c6.startswith("a")

        home_point = any(is_home_point(c6) for c6, _ in r)
        away_point = any(is_away_point(c6) for c6, _ in r)

        home_rece = any(is_home_rece(c6) for c6, _ in r)
        away_rece = any(is_away_rece(c6) for c6, _ in r)

        home_spin = any(is_home_spin(c6) for c6, _ in r)
        away_spin = any(is_away_spin(c6) for c6, _ in r)

        home_float = any(is_home_float(c6) for c6, _ in r)
        away_float = any(is_away_float(c6) for c6, _ in r)

        # SideOut totale
        if home_rece:
            so_home_attempts += 1
            if home_point:
                so_home_wins += 1

        if away_rece:
            so_away_attempts += 1
            if away_point:
                so_away_wins += 1

        # SPIN
        if home_spin:
            so_spin_home_attempts += 1
            if home_point:
                so_spin_home_wins += 1

        if away_spin:
            so_spin_away_attempts += 1
            if away_point:
                so_spin_away_wins += 1

        # FLOAT
        if home_float:
            so_float_home_attempts += 1
            if home_point:
                so_float_home_wins += 1

        if away_float:
            so_float_away_attempts += 1
            if away_point:
                so_float_away_wins += 1

        # DIRETTO
        rally_c6 = [c6 for c6, _ in r]
        if home_rece and home_point and first_attack_after_reception_is_winner(rally_c6, "*"):
            so_dir_home_wins += 1

        if away_rece and away_point and first_attack_after_reception_is_winner(rally_c6, "a"):
            so_dir_away_wins += 1

        # GIOCATO (# + ! -)
        home_play = any(is_home_rece_playable(c6) for c6, _ in r)
        away_play = any(is_away_rece_playable(c6) for c6, _ in r)

        if home_play:
            so_play_home_attempts += 1
            if home_point:
                so_play_home_wins += 1

        if away_play:
            so_play_away_attempts += 1
            if away_point:
                so_play_away_wins += 1

        # BUONA (#,+)
        home_good = any(is_home_rece_good(c6) for c6, _ in r)
        away_good = any(is_away_rece_good(c6) for c6, _ in r)

        if home_good:
            so_good_home_attempts += 1
            if home_point:
                so_good_home_wins += 1

        if away_good:
            so_good_away_attempts += 1
            if away_point:
                so_good_away_wins += 1

        # ESCLAMATIVA (!)
        home_exc = any(is_home_rece_exc(c6) for c6, _ in r)
        away_exc = any(is_away_rece_exc(c6) for c6, _ in r)

        if home_exc:
            so_exc_home_attempts += 1
            if home_point:
                so_exc_home_wins += 1

        if away_exc:
            so_exc_away_attempts += 1
            if away_point:
                so_exc_away_wins += 1

        # NEGATIVA (-)
        home_neg = any(is_home_rece_neg(c6) for c6, _ in r)
        away_neg = any(is_away_rece_neg(c6) for c6, _ in r)

        if home_neg:
            so_neg_home_attempts += 1
            if home_point:
                so_neg_home_wins += 1

        if away_neg:
            so_neg_away_attempts += 1
            if away_point:
                so_neg_away_wins += 1

        # Break totale
        if home_served:
            bp_home_attempts += 1
            if home_point:
                bp_home_wins += 1

        if away_served:
            bp_away_attempts += 1
            if away_point:
                bp_away_wins += 1

        # Break giocato + BT
        bt = detect_bt(first_raw)
        if bt is not None:
            if home_served:
                bp_play_home_attempts += 1
                if home_point:
                    bp_play_home_wins += 1

                if bt == "NEG":
                    bt_neg_home += 1
                elif bt == "POS":
                    bt_pos_home += 1
                elif bt == "EXC":
                    bt_exc_home += 1
                elif bt == "HALF":
                    bt_half_home += 1

            if away_served:
                bp_play_away_attempts += 1
                if away_point:
                    bp_play_away_wins += 1

                if bt == "NEG":
                    bt_neg_away += 1
                elif bt == "POS":
                    bt_pos_away += 1
                elif bt == "EXC":
                    bt_exc_away += 1
                elif bt == "HALF":
                    bt_half_away += 1

    return {
        "so_home_attempts": so_home_attempts,
        "so_home_wins": so_home_wins,
        "so_away_attempts": so_away_attempts,
        "so_away_wins": so_away_wins,
        "sideout_home_pct": pct(so_home_wins, so_home_attempts),
        "sideout_away_pct": pct(so_away_wins, so_away_attempts),

        "bp_home_attempts": bp_home_attempts,
        "bp_home_wins": bp_home_wins,
        "bp_away_attempts": bp_away_attempts,
        "bp_away_wins": bp_away_wins,
        "break_home_pct": pct(bp_home_wins, bp_home_attempts),
        "break_away_pct": pct(bp_away_wins, bp_away_attempts),

        "so_spin_home_attempts": so_spin_home_attempts,
        "so_spin_home_wins": so_spin_home_wins,
        "so_spin_away_attempts": so_spin_away_attempts,
        "so_spin_away_wins": so_spin_away_wins,

        "so_float_home_attempts": so_float_home_attempts,
        "so_float_home_wins": so_float_home_wins,
        "so_float_away_attempts": so_float_away_attempts,
        "so_float_away_wins": so_float_away_wins,

        "so_dir_home_wins": so_dir_home_wins,
        "so_dir_away_wins": so_dir_away_wins,

        "so_play_home_attempts": so_play_home_attempts,
        "so_play_home_wins": so_play_home_wins,
        "so_play_away_attempts": so_play_away_attempts,
        "so_play_away_wins": so_play_away_wins,

        "so_good_home_attempts": so_good_home_attempts,
        "so_good_home_wins": so_good_home_wins,
        "so_good_away_attempts": so_good_away_attempts,
        "so_good_away_wins": so_good_away_wins,

        "so_exc_home_attempts": so_exc_home_attempts,
        "so_exc_home_wins": so_exc_home_wins,
        "so_exc_away_attempts": so_exc_away_attempts,
        "so_exc_away_wins": so_exc_away_wins,

        "so_neg_home_attempts": so_neg_home_attempts,
        "so_neg_home_wins": so_neg_home_wins,
        "so_neg_away_attempts": so_neg_away_attempts,
        "so_neg_away_wins": so_neg_away_wins,

        # NEW: Break giocato + BT
        "bp_play_home_attempts": bp_play_home_attempts,
        "bp_play_home_wins": bp_play_home_wins,
        "bp_play_away_attempts": bp_play_away_attempts,
        "bp_play_away_wins": bp_play_away_wins,

        "bt_neg_home": bt_neg_home,
        "bt_pos_home": bt_pos_home,
        "bt_exc_home": bt_exc_home,
        "bt_half_home": bt_half_home,

        "bt_neg_away": bt_neg_away,
        "bt_pos_away": bt_pos_away,
        "bt_exc_away": bt_exc_away,
        "bt_half_away": bt_half_away,
    }


# =========================
# DB INIT + MIGRATION
# =========================
def init_db():
    try:
        """Create required tables on Postgres (Supabase)."""
        with engine.begin() as conn:
            # =========================
            # MATCHES
            # =========================
            conn.execute(text("""
                CREATE TABLE IF NOT EXISTS matches (
                    id BIGSERIAL PRIMARY KEY,
                    filename TEXT,
                    phase TEXT,
                    round_number INTEGER,
                    season TEXT,
                    competition TEXT,
                    team_a TEXT,
                    team_b TEXT,
                    n_azioni INTEGER,
                    preview TEXT,
                    scout_text TEXT,
                    match_key TEXT UNIQUE,
                    created_at TEXT,

                    so_home_attempts INTEGER,
                    so_home_wins INTEGER,
                    so_away_attempts INTEGER,
                    so_away_wins INTEGER,
                    sideout_home_pct DOUBLE PRECISION,
                    sideout_away_pct DOUBLE PRECISION,

                    bp_home_attempts INTEGER,
                    bp_home_wins INTEGER,
                    bp_away_attempts INTEGER,
                    bp_away_wins INTEGER,
                    break_home_pct DOUBLE PRECISION,
                    break_away_pct DOUBLE PRECISION,

                    so_spin_home_attempts INTEGER,
                    so_spin_home_wins INTEGER,
                    so_spin_away_attempts INTEGER,
                    so_spin_away_wins INTEGER,

                    so_float_home_attempts INTEGER,
                    so_float_home_wins INTEGER,
                    so_float_away_attempts INTEGER,
                    so_float_away_wins INTEGER,

                    so_dir_home_wins INTEGER,
                    so_dir_away_wins INTEGER,

                    so_play_home_attempts INTEGER,
                    so_play_home_wins INTEGER,
                    so_play_away_attempts INTEGER,
                    so_play_away_wins INTEGER,

                    so_good_home_attempts INTEGER,
                    so_good_home_wins INTEGER,
                    so_good_away_attempts INTEGER,
                    so_good_away_wins INTEGER,

                    so_exc_home_attempts INTEGER,
                    so_exc_home_wins INTEGER,
                    so_exc_away_attempts INTEGER,
                    so_exc_away_wins INTEGER,

                    so_neg_home_attempts INTEGER,
                    so_neg_home_wins INTEGER,
                    so_neg_away_attempts INTEGER,
                    so_neg_away_wins INTEGER,

                    -- Break “giocato” + BT
                    bp_play_home_attempts INTEGER,
                    bp_play_home_wins INTEGER,
                    bp_play_away_attempts INTEGER,
                    bp_play_away_wins INTEGER,

                    bt_neg_home INTEGER,
                    bt_pos_home INTEGER,
                    bt_exc_home INTEGER,
                    bt_half_home INTEGER,

                    bt_neg_away INTEGER,
                    bt_pos_away INTEGER,
                    bt_exc_away INTEGER,
                    bt_half_away INTEGER
                );
            """))

            # =========================
            # ROSTER
            # =========================
            conn.execute(text("""
                CREATE TABLE IF NOT EXISTS roster (
                    id BIGSERIAL PRIMARY KEY,
                    season TEXT,
                    team_raw TEXT,
                    team_norm TEXT,
                    jersey_number INTEGER,
                    player_name TEXT,
                    role TEXT,
                    created_at TEXT,
                    UNIQUE(season, team_norm, jersey_number)
                );
            """))

            # =========================
            # INDICI
            # =========================
            conn.execute(text("CREATE INDEX IF NOT EXISTS idx_matches_round ON matches(round_number);"))
            conn.execute(text("CREATE INDEX IF NOT EXISTS idx_matches_key ON matches(match_key);"))
            conn.execute(text("CREATE INDEX IF NOT EXISTS idx_roster_season_team ON roster(season, team_norm);"))

    except Exception as e:
        st.error("Errore connessione/inizializzazione DB (Supabase).")
        st.code(str(e))
        st.stop()

def render_import(admin_mode: bool):
    st.header("Import multiplo DVW (settimana)")

    if not admin_mode:
        st.warning("Accesso riservato allo staff (admin).")
        return

    uploaded_files = st.file_uploader(
        "Carica uno o più file .dvw",
        type=["dvw"],
        accept_multiple_files=True
    )

    st.divider()
    st.subheader("Elimina un import (dal database)")
    with engine.begin() as conn:
        del_rows = conn.execute(text("""
            SELECT id, filename, team_a, team_b, phase, round_number, created_at
            FROM matches
            ORDER BY id DESC
            LIMIT 200
        """)).mappings().all()

    if del_rows:
        def label(r):
            rn = int(r.get("round_number") or 0)
            ph = r.get("phase") or ""
            return f"[id {r['id']}] {r['filename']} — {r.get('team_a','')} vs {r.get('team_b','')} ({ph}{rn:02d})"

        selected = st.selectbox("Seleziona il match da eliminare", del_rows, format_func=label)
        confirm = st.checkbox("Confermo: voglio eliminare questo match dal DB", value=False)

        if st.button("Elimina selezionato", disabled=not confirm):
            with engine.begin() as conn:
                conn.execute(text("DELETE FROM matches WHERE id = :id"), {"id": selected["id"]})
            st.success("Eliminato dal DB.")
            st.rerun()
    else:
        st.info("Nessun match da eliminare.")

    st.divider()

    if not uploaded_files:
        st.info("Seleziona uno o più file DVW per importare.")
        return

    st.write(f"File selezionati: {len(uploaded_files)}")

    if st.button("Importa tutti (con dedup/upssert)"):
        saved = 0
        errors = 0
        details = []

        sql_upsert = """
        INSERT INTO matches
        (filename, phase, round_number, season, competition, team_a, team_b,
         n_azioni, preview, scout_text, match_key, created_at,

         so_home_attempts, so_home_wins, so_away_attempts, so_away_wins,
         sideout_home_pct, sideout_away_pct,

         bp_home_attempts, bp_home_wins, bp_away_attempts, bp_away_wins,
         break_home_pct, break_away_pct,

         so_spin_home_attempts, so_spin_home_wins, so_spin_away_attempts, so_spin_away_wins,
         so_float_home_attempts, so_float_home_wins, so_float_away_attempts, so_float_away_wins,
         so_dir_home_wins, so_dir_away_wins,

         so_play_home_attempts, so_play_home_wins, so_play_away_attempts, so_play_away_wins,
         so_good_home_attempts, so_good_home_wins, so_good_away_attempts, so_good_away_wins,
         so_exc_home_attempts, so_exc_home_wins, so_exc_away_attempts, so_exc_away_wins,
         so_neg_home_attempts, so_neg_home_wins, so_neg_away_attempts, so_neg_away_wins,

         bp_play_home_attempts, bp_play_home_wins, bp_play_away_attempts, bp_play_away_wins,
         bt_neg_home, bt_pos_home, bt_exc_home, bt_half_home,
         bt_neg_away, bt_pos_away, bt_exc_away, bt_half_away
        )
        VALUES
        (:filename, :phase, :round_number, :season, :competition, :team_a, :team_b,
         :n_azioni, :preview, :scout_text, :match_key, :created_at,

         :so_home_attempts, :so_home_wins, :so_away_attempts, :so_away_wins,
         :sideout_home_pct, :sideout_away_pct,

         :bp_home_attempts, :bp_home_wins, :bp_away_attempts, :bp_away_wins,
         :break_home_pct, :break_away_pct,

         :so_spin_home_attempts, :so_spin_home_wins, :so_spin_away_attempts, :so_spin_away_wins,
         :so_float_home_attempts, :so_float_home_wins, :so_float_away_attempts, :so_float_away_wins,
         :so_dir_home_wins, :so_dir_away_wins,

         :so_play_home_attempts, :so_play_home_wins, :so_play_away_attempts, :so_play_away_wins,
         :so_good_home_attempts, :so_good_home_wins, :so_good_away_attempts, :so_good_away_wins,
         :so_exc_home_attempts, :so_exc_home_wins, :so_exc_away_attempts, :so_exc_away_wins,
         :so_neg_home_attempts, :so_neg_home_wins, :so_neg_away_attempts, :so_neg_away_wins,

         :bp_play_home_attempts, :bp_play_home_wins, :bp_play_away_attempts, :bp_play_away_wins,
         :bt_neg_home, :bt_pos_home, :bt_exc_home, :bt_half_home,
         :bt_neg_away, :bt_pos_away, :bt_exc_away, :bt_half_away
        )
        ON CONFLICT(match_key) DO UPDATE SET
            filename = excluded.filename,
            phase = excluded.phase,
            round_number = excluded.round_number,
            season = excluded.season,
            competition = excluded.competition,
            team_a = excluded.team_a,
            team_b = excluded.team_b,
            n_azioni = excluded.n_azioni,
            preview = excluded.preview,
            scout_text = excluded.scout_text,
            created_at = excluded.created_at,

            so_home_attempts = excluded.so_home_attempts,
            so_home_wins = excluded.so_home_wins,
            so_away_attempts = excluded.so_away_attempts,
            so_away_wins = excluded.so_away_wins,
            sideout_home_pct = excluded.sideout_home_pct,
            sideout_away_pct = excluded.sideout_away_pct,

            bp_home_attempts = excluded.bp_home_attempts,
            bp_home_wins = excluded.bp_home_wins,
            bp_away_attempts = excluded.bp_away_attempts,
            bp_away_wins = excluded.bp_away_wins,
            break_home_pct = excluded.break_home_pct,
            break_away_pct = excluded.break_away_pct,

            so_spin_home_attempts = excluded.so_spin_home_attempts,
            so_spin_home_wins = excluded.so_spin_home_wins,
            so_spin_away_attempts = excluded.so_spin_away_attempts,
            so_spin_away_wins = excluded.so_spin_away_wins,

            so_float_home_attempts = excluded.so_float_home_attempts,
            so_float_home_wins = excluded.so_float_home_wins,
            so_float_away_attempts = excluded.so_float_away_attempts,
            so_float_away_wins = excluded.so_float_away_wins,

            so_dir_home_wins = excluded.so_dir_home_wins,
            so_dir_away_wins = excluded.so_dir_away_wins,

            so_play_home_attempts = excluded.so_play_home_attempts,
            so_play_home_wins = excluded.so_play_home_wins,
            so_play_away_attempts = excluded.so_play_away_attempts,
            so_play_away_wins = excluded.so_play_away_wins,

            so_good_home_attempts = excluded.so_good_home_attempts,
            so_good_home_wins = excluded.so_good_home_wins,
            so_good_away_attempts = excluded.so_good_away_attempts,
            so_good_away_wins = excluded.so_good_away_wins,

            so_exc_home_attempts = excluded.so_exc_home_attempts,
            so_exc_home_wins = excluded.so_exc_home_wins,
            so_exc_away_attempts = excluded.so_exc_away_attempts,
            so_exc_away_wins = excluded.so_exc_away_wins,

            so_neg_home_attempts = excluded.so_neg_home_attempts,
            so_neg_home_wins = excluded.so_neg_home_wins,
            so_neg_away_attempts = excluded.so_neg_away_attempts,
            so_neg_away_wins = excluded.so_neg_away_wins,

            bp_play_home_attempts = excluded.bp_play_home_attempts,
            bp_play_home_wins = excluded.bp_play_home_wins,
            bp_play_away_attempts = excluded.bp_play_away_attempts,
            bp_play_away_wins = excluded.bp_play_away_wins,

            bt_neg_home = excluded.bt_neg_home,
            bt_pos_home = excluded.bt_pos_home,
            bt_exc_home = excluded.bt_exc_home,
            bt_half_home = excluded.bt_half_home,

            bt_neg_away = excluded.bt_neg_away,
            bt_pos_away = excluded.bt_pos_away,
            bt_exc_away = excluded.bt_exc_away,
            bt_half_away = excluded.bt_half_away
        """

        with engine.begin() as conn:
            for uf in uploaded_files:
                filename = uf.name
                try:
                    dvw_text = uf.getvalue().decode("utf-8", errors="ignore")
                    phase, round_number = extract_round_code(filename)
                    parsed = parse_dvw_minimal(dvw_text)
                    scout_lines = extract_scout_lines(dvw_text)
                    counts = compute_counts_from_scout(scout_lines)

                    match_key = build_match_key(
                        parsed.get("team_a", ""),
                        parsed.get("team_b", ""),
                        parsed.get("competition", ""),
                        phase,
                        round_number,
                    )
                    preview = " | ".join([code6(x) for x in scout_lines[:3]])

                    params = {
                        "filename": filename,
                        "phase": phase,
                        "round_number": int(round_number),
                        "season": parsed.get("season"),
                        "competition": parsed.get("competition"),
                        "team_a": parsed.get("team_a"),
                        "team_b": parsed.get("team_b"),
                        "n_azioni": int(len(scout_lines)),
                        "preview": preview,
                        "scout_text": "\n".join(scout_lines),
                        "match_key": match_key,
                        "created_at": datetime.now(timezone.utc).isoformat(),
                        **counts,
                    }

                    conn.execute(text(sql_upsert), params)
                    saved += 1
                    details.append({"file": filename, "esito": "OK"})
                except Exception as e:
                    errors += 1
                    details.append({"file": filename, "esito": f"ERRORE: {e}"})

        st.success(f"Fatto. Importati/aggiornati: {saved} | Errori: {errors}")
        st.dataframe(pd.DataFrame(details), width="stretch", hide_index=True)
        st.rerun()


# =========================
# UI: SIDEOUT TEAM (già completa)
# =========================
def render_sideout_team():
    st.header("Indici Side Out - Squadre")

    voce = st.radio(
        "Seleziona indice",
        [
            "Side Out TOTALE",
            "Side Out SPIN",
            "Side Out FLOAT",
            "Side Out DIRETTO",
            "Side Out GIOCATO",
            "Side Out con RICE BUONA",
            "Side Out con RICE ESCALAMATIVA",
            "Side Out con RICE NEGATIVA",
        ],
        index=0,
    )

    with engine.begin() as conn:
        bounds = conn.execute(text("""
            SELECT MIN(round_number) AS min_r, MAX(round_number) AS max_r
            FROM matches
            WHERE round_number IS NOT NULL
        """)).mappings().first()

    min_r = int((bounds["min_r"] or 1))
    max_r = int((bounds["max_r"] or 1))

    from_round, to_round, _range_label = get_selected_match_range()
    st.caption(f"📌 Range: {_range_label}")

    if from_round > to_round:
        st.error("Range non valido: 'Da giornata' deve essere <= 'A giornata'.")
        st.stop()

    def highlight_perugia(row):
        is_perugia = "perugia" in str(row["squadra"]).lower()
        style = "background-color: #fff3cd; font-weight: 800;" if is_perugia else ""
        return [style] * len(row)

    def show_table(df: pd.DataFrame, fmt: dict):
        if df.empty:
            st.info("Nessun dato nel range selezionato.")
            return
        styled = (
            df.style
              .apply(highlight_perugia, axis=1)
              .format(fmt)
              .set_properties(subset=[c for c in df.columns if ('%' in str(c))], **{'background-color': '#e7f5ff', 'font-weight': '800'})
              .set_table_styles([
                  {"selector": "th", "props": [("font-size", "24px"), ("text-align", "left"), ("padding", "10px 12px")]},
                  {"selector": "td", "props": [("font-size", "23px"), ("padding", "10px 12px")]},
              ])
        )
        render_any_table(styled)

    # --- TOTALE ---
    if voce == "Side Out TOTALE":
        with engine.begin() as conn:
            rows = conn.execute(
                text("""
                    SELECT
                        squadra,
                        SUM(n_ricezioni) AS n_ricezioni,
                        SUM(n_sideout)   AS n_sideout,
                        COALESCE(ROUND(100.0 * SUM(n_sideout) / NULLIF(SUM(n_ricezioni), 0), 1), 0.0) AS so_pct
                    FROM (
                        SELECT team_a AS squadra,
                               COALESCE(so_home_attempts, 0) AS n_ricezioni,
                               COALESCE(so_home_wins, 0)     AS n_sideout
                        FROM matches
                        WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                        UNION ALL
                        SELECT team_b AS squadra,
                               COALESCE(so_away_attempts, 0) AS n_ricezioni,
                               COALESCE(so_away_wins, 0)     AS n_sideout
                        FROM matches
                        WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                    )
                    GROUP BY squadra
                    ORDER BY so_pct DESC, n_ricezioni DESC
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        df = pd.DataFrame(rows).rename(columns={
            "squadra": "squadra",
            "so_pct": "% S.O.",
            "n_ricezioni": "n° ricezioni",
            "n_sideout": "n° Side Out",
        })
        # Canonicalizza nomi squadra (sponsor/varianti) e raggruppa a 12 squadre
        df["TEAM_CANON"] = df["squadra"].apply(canonical_team)
        df = (df.groupby("TEAM_CANON", as_index=False)
                .agg({"n° ricezioni": "sum", "n° Side Out": "sum"}))
        df["% S.O."] = (100.0 * df["n° Side Out"] / df["n° ricezioni"].replace({0: pd.NA})).round(1).fillna(0.0)
        df = df.rename(columns={"TEAM_CANON": "squadra"})
        # Ordina per % S.O. (colonna guida) e poi per volume
        df = df.sort_values(["% S.O.", "n° ricezioni"], ascending=[False, False])
        # Ranking 1..12 come prima colonna
        df = df.head(12).reset_index(drop=True)
        df.insert(0, "Ranking", range(1, len(df) + 1))
        df = df[["Ranking", "squadra", "% S.O.", "n° ricezioni", "n° Side Out"]].copy()
        show_table(df, {"Ranking": "{:.0f}", "% S.O.": "{:.1f}", "n° ricezioni": "{:.0f}", "n° Side Out": "{:.0f}"})

    # --- SPIN ---
    elif voce == "Side Out SPIN":
        with engine.begin() as conn:
            rows = conn.execute(
                text("""
                    SELECT
                        squadra,
                        SUM(spin_att) AS spin_att,
                        SUM(spin_win) AS spin_win,
                        SUM(tot_att) AS tot_att,
                        COALESCE(ROUND(100.0 * SUM(spin_win) / NULLIF(SUM(spin_att), 0), 1), 0.0) AS so_spin_pct,
                        COALESCE(ROUND(100.0 * SUM(spin_att) / NULLIF(SUM(tot_att), 0), 1), 0.0) AS spin_share_pct
                    FROM (
                        SELECT team_a AS squadra,
                               COALESCE(so_spin_home_attempts, 0) AS spin_att,
                               COALESCE(so_spin_home_wins, 0)     AS spin_win,
                               COALESCE(so_home_attempts, 0)      AS tot_att
                        FROM matches
                        WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                        UNION ALL
                        SELECT team_b AS squadra,
                               COALESCE(so_spin_away_attempts, 0) AS spin_att,
                               COALESCE(so_spin_away_wins, 0)     AS spin_win,
                               COALESCE(so_away_attempts, 0)      AS tot_att
                        FROM matches
                        WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                    )
                    GROUP BY squadra
                    ORDER BY so_spin_pct DESC, spin_att DESC
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        df_spin = pd.DataFrame(rows).rename(columns={
            "squadra": "squadra",
            "so_spin_pct": "% S.O. SPIN",
            "spin_att": "n° ricezioni SPIN",
            "spin_win": "n° Side Out SPIN",
            "spin_share_pct": "% SPIN su TOT",
            "tot_att": "n° ricezioni TOT",
        })
        # Canonicalizza e raggruppa squadre (sponsor/varianti -> 12)
        df_spin["TEAM_CANON"] = df_spin["squadra"].apply(canonical_team)
        df_spin = (df_spin.groupby("TEAM_CANON", as_index=False)
                          .agg({"n° ricezioni SPIN": "sum", "n° Side Out SPIN": "sum", "n° ricezioni TOT": "sum"}))
        df_spin = df_spin.rename(columns={"TEAM_CANON": "squadra"})
        # Ricalcolo percentuali dopo raggruppamento
        df_spin["% S.O. SPIN"] = (100.0 * df_spin["n° Side Out SPIN"] / df_spin["n° ricezioni SPIN"].replace({0: pd.NA})).round(1).fillna(0.0)
        df_spin["% SPIN su TOT"] = (100.0 * df_spin["n° ricezioni SPIN"] / df_spin["n° ricezioni TOT"].replace({0: pd.NA})).round(1).fillna(0.0)
        df_spin = df_spin.sort_values(['% S.O. SPIN', 'n° ricezioni SPIN'], ascending=[False, False])
        df_spin = df_spin.head(12).reset_index(drop=True)
        df_spin.insert(0, "Ranking", range(1, len(df_spin) + 1))
        df_spin = df_spin[['Ranking', 'squadra', '% S.O. SPIN', 'n° ricezioni SPIN', 'n° Side Out SPIN', '% SPIN su TOT']].copy()
        show_table(df_spin, {"Ranking": "{:.0f}", "% S.O. SPIN": "{:.1f}", "n° ricezioni SPIN": "{:.0f}",
                             "n° Side Out SPIN": "{:.0f}", "% SPIN su TOT": "{:.1f}"})

    # --- FLOAT ---
    elif voce == "Side Out FLOAT":
        with engine.begin() as conn:
            rows = conn.execute(
                text("""
                    SELECT
                        squadra,
                        SUM(float_att) AS float_att,
                        SUM(float_win) AS float_win,
                        SUM(tot_att) AS tot_att,
                        COALESCE(ROUND(100.0 * SUM(float_win) / NULLIF(SUM(float_att), 0), 1), 0.0) AS so_float_pct,
                        COALESCE(ROUND(100.0 * SUM(float_att) / NULLIF(SUM(tot_att), 0), 1), 0.0) AS float_share_pct
                    FROM (
                        SELECT team_a AS squadra,
                               COALESCE(so_float_home_attempts, 0) AS float_att,
                               COALESCE(so_float_home_wins, 0)     AS float_win,
                               COALESCE(so_home_attempts, 0)       AS tot_att
                        FROM matches
                        WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                        UNION ALL
                        SELECT team_b AS squadra,
                               COALESCE(so_float_away_attempts, 0) AS float_att,
                               COALESCE(so_float_away_wins, 0)     AS float_win,
                               COALESCE(so_away_attempts, 0)       AS tot_att
                        FROM matches
                        WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                    )
                    GROUP BY squadra
                    ORDER BY so_float_pct DESC, float_att DESC
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        df_float = pd.DataFrame(rows).rename(columns={
            "squadra": "squadra",
            "so_float_pct": "% S.O. FLOAT",
            "float_att": "n° ricezioni FLOAT",
            "float_win": "n° Side Out FLOAT",
            "float_share_pct": "% FLOAT su TOT",
            "tot_att": "n° ricezioni TOT",
        })
        # Canonicalizza e raggruppa squadre (sponsor/varianti -> 12)
        df_float["TEAM_CANON"] = df_float["squadra"].apply(canonical_team)
        df_float = (df_float.groupby("TEAM_CANON", as_index=False)
                          .agg({"n° ricezioni FLOAT": "sum", "n° Side Out FLOAT": "sum", "n° ricezioni TOT": "sum"}))
        df_float = df_float.rename(columns={"TEAM_CANON": "squadra"})
        # Ricalcolo percentuali dopo raggruppamento
        df_float["% S.O. FLOAT"] = (100.0 * df_float["n° Side Out FLOAT"] / df_float["n° ricezioni FLOAT"].replace({0: pd.NA})).round(1).fillna(0.0)
        df_float["% FLOAT su TOT"] = (100.0 * df_float["n° ricezioni FLOAT"] / df_float["n° ricezioni TOT"].replace({0: pd.NA})).round(1).fillna(0.0)
        df_float = df_float.sort_values(['% S.O. FLOAT', 'n° ricezioni FLOAT'], ascending=[False, False])
        df_float = df_float.head(12).reset_index(drop=True)
        df_float.insert(0, "Ranking", range(1, len(df_float) + 1))
        df_float = df_float[['Ranking', 'squadra', '% S.O. FLOAT', 'n° ricezioni FLOAT', 'n° Side Out FLOAT', '% FLOAT su TOT']].copy()
        show_table(df_float, {"Ranking": "{:.0f}", "% S.O. FLOAT": "{:.1f}", "n° ricezioni FLOAT": "{:.0f}",
                              "n° Side Out FLOAT": "{:.0f}", "% FLOAT su TOT": "{:.1f}"})

    # --- DIRETTO ---
    elif voce == "Side Out DIRETTO":
        with engine.begin() as conn:
            rows = conn.execute(
                text("""
                    SELECT
                        squadra,
                        SUM(tot_att) AS n_ricezioni,
                        SUM(dir_win) AS n_sideout_dir,
                        COALESCE(ROUND(100.0 * SUM(dir_win) / NULLIF(SUM(tot_att), 0), 1), 0.0) AS so_dir_pct
                    FROM (
                        SELECT team_a AS squadra,
                               COALESCE(so_home_attempts, 0) AS tot_att,
                               COALESCE(so_dir_home_wins, 0) AS dir_win
                        FROM matches
                        WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                        UNION ALL
                        SELECT team_b AS squadra,
                               COALESCE(so_away_attempts, 0) AS tot_att,
                               COALESCE(so_dir_away_wins, 0) AS dir_win
                        FROM matches
                        WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                    )
                    GROUP BY squadra
                    ORDER BY so_dir_pct DESC, n_ricezioni DESC
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        df_dir = pd.DataFrame(rows).rename(columns={
            "squadra": "squadra",
            "so_dir_pct": "% S.O. DIR",
            "n_ricezioni": "n° ricezioni",
            "n_sideout_dir": "n° Side Out DIR",
        })
        # Canonicalizza e raggruppa squadre (sponsor/varianti -> 12)
        df_dir["TEAM_CANON"] = df_dir["squadra"].apply(canonical_team)
        df_dir = (df_dir.groupby("TEAM_CANON", as_index=False)
                          .agg({"n° ricezioni": "sum", "n° Side Out DIR": "sum"}))
        df_dir = df_dir.rename(columns={"TEAM_CANON": "squadra"})
        # Ricalcolo percentuali dopo raggruppamento
        df_dir["% S.O. DIR"] = (100.0 * df_dir["n° Side Out DIR"] / df_dir["n° ricezioni"].replace({0: pd.NA})).round(1).fillna(0.0)
        df_dir = df_dir.sort_values(['% S.O. DIR', 'n° ricezioni'], ascending=[False, False])
        df_dir = df_dir.head(12).reset_index(drop=True)
        df_dir.insert(0, "Ranking", range(1, len(df_dir) + 1))
        df_dir = df_dir[['Ranking', 'squadra', '% S.O. DIR', 'n° ricezioni', 'n° Side Out DIR']].copy()
        show_table(df_dir, {"Ranking": "{:.0f}", "% S.O. DIR": "{:.1f}", "n° ricezioni": "{:.0f}", "n° Side Out DIR": "{:.0f}"})

    # --- GIOCATO ---
    elif voce == "Side Out GIOCATO":
        with engine.begin() as conn:
            rows = conn.execute(
                text("""
                    SELECT
                        squadra,
                        SUM(att) AS n_ricezioni_giocato,
                        SUM(win) AS n_sideout_giocato,
                        COALESCE(ROUND(100.0 * SUM(win) / NULLIF(SUM(att), 0), 1), 0.0) AS so_giocato_pct
                    FROM (
                        SELECT team_a AS squadra,
                               COALESCE(so_play_home_attempts, 0) AS att,
                               COALESCE(so_play_home_wins, 0)     AS win
                        FROM matches
                        WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                        UNION ALL
                        SELECT team_b AS squadra,
                               COALESCE(so_play_away_attempts, 0) AS att,
                               COALESCE(so_play_away_wins, 0)     AS win
                        FROM matches
                        WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                    )
                    GROUP BY squadra
                    ORDER BY so_giocato_pct DESC, n_ricezioni_giocato DESC
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        df_g = pd.DataFrame(rows).rename(columns={
            "squadra": "squadra",
            "so_giocato_pct": "% S.O. GIOCATO",
            "n_ricezioni_giocato": "n° ricezioni (giocabili)",
            "n_sideout_giocato": "n° Side Out",
        })
        # Canonicalizza e raggruppa squadre (sponsor/varianti -> 12)
        df_g["TEAM_CANON"] = df_g["squadra"].apply(canonical_team)
        df_g = (df_g.groupby("TEAM_CANON", as_index=False)
                          .agg({"n° ricezioni (giocabili)": "sum", "n° Side Out": "sum"}))
        df_g = df_g.rename(columns={"TEAM_CANON": "squadra"})
        # Ricalcolo percentuali dopo raggruppamento
        df_g["% S.O. GIOCATO"] = (100.0 * df_g["n° Side Out"] / df_g["n° ricezioni (giocabili)"].replace({0: pd.NA})).round(1).fillna(0.0)
        df_g = df_g.sort_values(['% S.O. GIOCATO', 'n° ricezioni (giocabili)'], ascending=[False, False])
        df_g = df_g.head(12).reset_index(drop=True)
        df_g.insert(0, "Ranking", range(1, len(df_g) + 1))
        df_g = df_g[['Ranking', 'squadra', '% S.O. GIOCATO', 'n° ricezioni (giocabili)', 'n° Side Out']].copy()
        show_table(df_g, {"Ranking": "{:.0f}", "% S.O. GIOCATO": "{:.1f}", "n° ricezioni (giocabili)": "{:.0f}", "n° Side Out": "{:.0f}"})

    # --- BUONA ---
    elif voce == "Side Out con RICE BUONA":
        with engine.begin() as conn:
            rows = conn.execute(
                text("""
                    SELECT
                        squadra,
                        SUM(att) AS n_ricezioni_buone,
                        SUM(win) AS n_sideout_buone,
                        COALESCE(ROUND(100.0 * SUM(win) / NULLIF(SUM(att), 0), 1), 0.0) AS so_buona_pct
                    FROM (
                        SELECT team_a AS squadra,
                               COALESCE(so_good_home_attempts, 0) AS att,
                               COALESCE(so_good_home_wins, 0)     AS win
                        FROM matches
                        WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                        UNION ALL
                        SELECT team_b AS squadra,
                               COALESCE(so_good_away_attempts, 0) AS att,
                               COALESCE(so_good_away_wins, 0)     AS win
                        FROM matches
                        WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                    )
                    GROUP BY squadra
                    ORDER BY so_buona_pct DESC, n_ricezioni_buone DESC
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        df_b = pd.DataFrame(rows).rename(columns={
            "squadra": "squadra",
            "so_buona_pct": "% S.O. RICE BUONA",
            "n_ricezioni_buone": "n° ricezioni (#,+)",
            "n_sideout_buone": "n° Side Out",
        })
        # Canonicalizza e raggruppa squadre (sponsor/varianti -> 12)
        df_b["TEAM_CANON"] = df_b["squadra"].apply(canonical_team)
        df_b = (df_b.groupby("TEAM_CANON", as_index=False)
                          .agg({"n° ricezioni (#,+)": "sum", "n° Side Out": "sum"}))
        df_b = df_b.rename(columns={"TEAM_CANON": "squadra"})
        # Ricalcolo percentuali dopo raggruppamento
        df_b["% S.O. RICE BUONA"] = (100.0 * df_b["n° Side Out"] / df_b["n° ricezioni (#,+)"].replace({0: pd.NA})).round(1).fillna(0.0)
        df_b = df_b.sort_values(['% S.O. RICE BUONA', 'n° ricezioni (#,+)'], ascending=[False, False])
        df_b = df_b.head(12).reset_index(drop=True)
        df_b.insert(0, "Ranking", range(1, len(df_b) + 1))
        df_b = df_b[['Ranking', 'squadra', '% S.O. RICE BUONA', 'n° ricezioni (#,+)', 'n° Side Out']].copy()
        show_table(df_b, {"Ranking": "{:.0f}", "% S.O. RICE BUONA": "{:.1f}", "n° ricezioni (#,+)": "{:.0f}", "n° Side Out": "{:.0f}"})

    # --- ESCLAMATIVA ---
    elif voce == "Side Out con RICE ESCALAMATIVA":
        with engine.begin() as conn:
            rows = conn.execute(
                text("""
                    SELECT
                        squadra,
                        SUM(att) AS n_ricezioni_exc,
                        SUM(win) AS n_sideout_exc,
                        COALESCE(ROUND(100.0 * SUM(win) / NULLIF(SUM(att), 0), 1), 0.0) AS so_exc_pct
                    FROM (
                        SELECT team_a AS squadra,
                               COALESCE(so_exc_home_attempts, 0) AS att,
                               COALESCE(so_exc_home_wins, 0)     AS win
                        FROM matches
                        WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                        UNION ALL
                        SELECT team_b AS squadra,
                               COALESCE(so_exc_away_attempts, 0) AS att,
                               COALESCE(so_exc_away_wins, 0)     AS win
                        FROM matches
                        WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                    )
                    GROUP BY squadra
                    ORDER BY so_exc_pct DESC, n_ricezioni_exc DESC
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        df_e = pd.DataFrame(rows).rename(columns={
            "squadra": "squadra",
            "so_exc_pct": "% S.O. RICE !",
            "n_ricezioni_exc": "n° ricezioni (!)",
            "n_sideout_exc": "n° Side Out",
        })
        # Canonicalizza e raggruppa squadre (sponsor/varianti -> 12)
        df_e["TEAM_CANON"] = df_e["squadra"].apply(canonical_team)
        df_e = (df_e.groupby("TEAM_CANON", as_index=False)
                          .agg({"n° ricezioni (!)": "sum", "n° Side Out": "sum"}))
        df_e = df_e.rename(columns={"TEAM_CANON": "squadra"})
        # Ricalcolo percentuali dopo raggruppamento
        df_e["% S.O. RICE !"] = (100.0 * df_e["n° Side Out"] / df_e["n° ricezioni (!)"].replace({0: pd.NA})).round(1).fillna(0.0)
        df_e = df_e.sort_values(['% S.O. RICE !', 'n° ricezioni (!)'], ascending=[False, False])
        df_e = df_e.head(12).reset_index(drop=True)
        df_e.insert(0, "Ranking", range(1, len(df_e) + 1))
        df_e = df_e[['Ranking', 'squadra', '% S.O. RICE !', 'n° ricezioni (!)', 'n° Side Out']].copy()
        show_table(df_e, {"Ranking": "{:.0f}", "% S.O. RICE !": "{:.1f}", "n° ricezioni (!)": "{:.0f}", "n° Side Out": "{:.0f}"})

    # --- NEGATIVA ---
    elif voce == "Side Out con RICE NEGATIVA":
        with engine.begin() as conn:
            rows = conn.execute(
                text("""
                    SELECT
                        squadra,
                        SUM(att) AS n_ricezioni_neg,
                        SUM(win) AS n_sideout_neg,
                        COALESCE(ROUND(100.0 * SUM(win) / NULLIF(SUM(att), 0), 1), 0.0) AS so_neg_pct
                    FROM (
                        SELECT team_a AS squadra,
                               COALESCE(so_neg_home_attempts, 0) AS att,
                               COALESCE(so_neg_home_wins, 0)     AS win
                        FROM matches
                        WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                        UNION ALL
                        SELECT team_b AS squadra,
                               COALESCE(so_neg_away_attempts, 0) AS att,
                               COALESCE(so_neg_away_wins, 0)     AS win
                        FROM matches
                        WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                    )
                    GROUP BY squadra
                    ORDER BY so_neg_pct DESC, n_ricezioni_neg DESC
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        df_n = pd.DataFrame(rows).rename(columns={
            "squadra": "squadra",
            "so_neg_pct": "% S.O. RICE -",
            "n_ricezioni_neg": "n° ricezioni (-)",
            "n_sideout_neg": "n° Side Out",
        })
        # Canonicalizza e raggruppa squadre (sponsor/varianti -> 12)
        df_n["TEAM_CANON"] = df_n["squadra"].apply(canonical_team)
        df_n = (df_n.groupby("TEAM_CANON", as_index=False)
                          .agg({"n° ricezioni (-)": "sum", "n° Side Out": "sum"}))
        df_n = df_n.rename(columns={"TEAM_CANON": "squadra"})
        # Ricalcolo percentuali dopo raggruppamento
        df_n["% S.O. RICE -"] = (100.0 * df_n["n° Side Out"] / df_n["n° ricezioni (-)"].replace({0: pd.NA})).round(1).fillna(0.0)
        df_n = df_n.sort_values(['% S.O. RICE -', 'n° ricezioni (-)'], ascending=[False, False])
        df_n = df_n.head(12).reset_index(drop=True)
        df_n.insert(0, "Ranking", range(1, len(df_n) + 1))
        df_n = df_n[['Ranking', 'squadra', '% S.O. RICE -', 'n° ricezioni (-)', 'n° Side Out']].copy()
        show_table(df_n, {"Ranking": "{:.0f}", "% S.O. RICE -": "{:.1f}", "n° ricezioni (-)": "{:.0f}", "n° Side Out": "{:.0f}"})


# =========================
# UI: BREAK TEAM
# =========================
def render_break_team():
    st.header("Indici Fase Break – Squadre")

    voce = st.radio(
        "Seleziona indice Break",
        [
            "BREAK TOTALE",
            "BREAK GIOCATO",
            "BREAK con BT. NEGATIVA",
            "BREAK con BT. ESCLAMATIVA",
            "BREAK con BT. POSITIVA",
            "BREAK con BT. 1/2 PUNTO",
            "BT punto/errore/ratio",            "Confronto TEAM",
            "GRAFICI",
        ],
        index=0,
    )

    # ===== FILTRO RANGE GIORNATE =====
    with engine.begin() as conn:
        bounds = conn.execute(text("""
            SELECT MIN(round_number) AS min_r, MAX(round_number) AS max_r
            FROM matches
            WHERE round_number IS NOT NULL
        """)).mappings().first()

    min_r = int(bounds["min_r"] or 1)
    max_r = int(bounds["max_r"] or 1)

    from_round, to_round, _range_label = get_selected_match_range()
    st.caption(f"📌 Range: {_range_label}")

    if from_round > to_round:
        st.error("Range non valido: 'Da giornata' deve essere <= 'A giornata'.")
        st.stop()

    def fix_team(name: str) -> str:
        # Canonicalizza sempre (sponsor/varianti -> 12 squadre)
        return canonical_team(name)

    def highlight_perugia(row):
        is_perugia = "perugia" in str(row["squadra"]).lower()
        style = "background-color: #fff3cd; font-weight: 800;" if is_perugia else ""
        return [style] * len(row)

    def show_table(df: pd.DataFrame, fmt: dict):
        if df.empty:
            st.info("Nessun dato nel range selezionato.")
            return
        # applica mapping richiesto
        if "squadra" in df.columns:
            df = df.copy()
            df["squadra"] = df["squadra"].apply(fix_team)

        styled = (
            df.style
              .apply(highlight_perugia, axis=1)
              .format(fmt)
              .set_properties(subset=[c for c in df.columns if ('%' in str(c))], **{'background-color': '#e7f5ff', 'font-weight': '800'})
              .set_table_styles([
                  {"selector": "th", "props": [("font-size", "24px"), ("text-align", "left"), ("padding", "10px 12px")]},
                  {"selector": "td", "props": [("font-size", "23px"), ("padding", "10px 12px")]},
              ])
        )
        render_any_table(styled)

    # ===== helper: calcola BT giocato/segni direttamente da scout_text (NON usa colonne bp_play_*/bt_* nel DB) =====
    def compute_break_bt_agg() -> pd.DataFrame:
        with engine.begin() as conn:
            rows = conn.execute(
                text("""
                    SELECT
                        team_a, team_b,
                        scout_text,
                        COALESCE(bp_home_attempts, 0) AS bp_home_attempts,
                        COALESCE(bp_away_attempts, 0) AS bp_away_attempts,
                        COALESCE(bp_home_wins, 0)     AS bp_home_wins,
                        COALESCE(bp_away_wins, 0)     AS bp_away_wins
                    FROM matches
                    WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        if not rows:
            return pd.DataFrame()

        agg = {}

        def ensure(team: str):
            team = fix_team(team)
            if team not in agg:
                agg[team] = {
                    "squadra": team,
                    "bt_att": 0,      # battute con valutazione (- + !)
                    "bt_bp": 0,       # break point su quelle battute
                    "bt_neg": 0,
                    "bt_pos": 0,
                    "bt_exc": 0,
                    "serves_tot": 0,  # battute totali (da DB)
                    "bp_tot": 0,      # break point totali (da DB)
                }
            return team

        def iter_rallies_from_scout_text(scout_text: str):
            if not scout_text:
                return []
            lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
            scout_lines = [ln for ln in lines if ln and ln[0] in ("*", "a")]

            rallies = []
            current = []
            for raw in scout_lines:
                c = code6(raw)
                if not c:
                    continue
                if is_serve(c):
                    if current:
                        rallies.append(current)
                    current = [c]
                    continue
                if not current:
                    continue
                current.append(c)
                if is_home_point(c) or is_away_point(c):
                    rallies.append(current)
                    current = []
            return rallies

        for r in rows:
            home_team = ensure(r.get("team_a") or "")
            away_team = ensure(r.get("team_b") or "")

            # totali dal DB
            agg[home_team]["serves_tot"] += int(r.get("bp_home_attempts") or 0)
            agg[away_team]["serves_tot"] += int(r.get("bp_away_attempts") or 0)
            agg[home_team]["bp_tot"] += int(r.get("bp_home_wins") or 0)
            agg[away_team]["bp_tot"] += int(r.get("bp_away_wins") or 0)

            rallies = iter_rallies_from_scout_text(r.get("scout_text") or "")
            for rally in rallies:
                first = rally[0]
                if len(first) < 6:
                    continue
                sign = first[5]  # '-' '+' '!' ecc.
                if sign not in ("-", "+", "!"):
                    continue  # non "giocato" secondo tua logica BT
                home_served = first.startswith("*")
                away_served = first.startswith("a")

                home_point = any(is_home_point(x) for x in rally)
                away_point = any(is_away_point(x) for x in rally)

                if home_served:
                    agg[home_team]["bt_att"] += 1
                    if home_point:
                        agg[home_team]["bt_bp"] += 1
                    if sign == "-":
                        agg[home_team]["bt_neg"] += 1
                    elif sign == "+":
                        agg[home_team]["bt_pos"] += 1
                    elif sign == "!":
                        agg[home_team]["bt_exc"] += 1

                if away_served:
                    agg[away_team]["bt_att"] += 1
                    if away_point:
                        agg[away_team]["bt_bp"] += 1
                    if sign == "-":
                        agg[away_team]["bt_neg"] += 1
                    elif sign == "+":
                        agg[away_team]["bt_pos"] += 1
                    elif sign == "!":
                        agg[away_team]["bt_exc"] += 1

        return pd.DataFrame(list(agg.values()))

    # ===== BREAK TOTALE (come prima) =====
    if voce == "BREAK TOTALE":
        with engine.begin() as conn:
            rows = conn.execute(text("""
                SELECT
                    squadra,
                    SUM(att) AS n_battute,
                    SUM(win) AS n_bpoint,
                    COALESCE(ROUND(100.0 * SUM(win) / NULLIF(SUM(att), 0), 1), 0.0) AS bp_pct
                FROM (
                    SELECT team_a AS squadra,
                           COALESCE(bp_home_attempts, 0) AS att,
                           COALESCE(bp_home_wins, 0)     AS win
                    FROM matches
                    WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                    UNION ALL
                    SELECT team_b AS squadra,
                           COALESCE(bp_away_attempts, 0) AS att,
                           COALESCE(bp_away_wins, 0)     AS win
                    FROM matches
                    WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                )
                GROUP BY squadra
                ORDER BY bp_pct DESC, n_battute DESC
            """), {"from_round": int(from_round), "to_round": int(to_round)}).mappings().all()

        df = pd.DataFrame(rows)
        if df.empty:
            st.info("Nessun dato nel range selezionato.")
            return

        df["squadra"] = df["squadra"].apply(fix_team)
        df = df.groupby("squadra", as_index=False).sum(numeric_only=True)
        df["% B.Point"] = (100.0 * df["n_bpoint"] / df["n_battute"].replace({0: pd.NA})).fillna(0.0)

        df = df.rename(columns={
            "n_battute": "n° Battute",
            "n_bpoint": "n° B.Point",
        })

        df = df.sort_values(by=["% B.Point", "n° Battute"], ascending=[False, False]).reset_index(drop=True)
        df.insert(0, "Ranking", range(1, len(df) + 1))
        df = df[["Ranking", "squadra", "% B.Point", "n° Battute", "n° B.Point"]].copy()

        show_table(df, {"Ranking": "{:.0f}", "% B.Point": "{:.1f}", "n° Battute": "{:.0f}", "n° B.Point": "{:.0f}"})

    # ===== BREAK GIOCATO (BT = '-' '+' '!') =====
    elif voce == "BREAK GIOCATO":
        base = compute_break_bt_agg()
        if base.empty:
            st.info("Nessun dato nel range selezionato.")
            return

        df = base.copy()
        df["% B.Point (Giocato)"] = (100.0 * df["bt_bp"] / df["bt_att"].replace({0: pd.NA})).fillna(0.0)

        df = df.rename(columns={
            "bt_att": "n° Battute (BT)",
            "bt_bp": "n° B.Point",
        })

        df = df.sort_values(by=["% B.Point (Giocato)", "n° Battute (BT)"], ascending=[False, False]).reset_index(drop=True)
        df.insert(0, "Ranking", range(1, len(df) + 1))

        df = df[["Ranking", "squadra", "% B.Point (Giocato)", "n° Battute (BT)", "n° B.Point"]].copy()
        show_table(df, {"Ranking": "{:.0f}", "% B.Point (Giocato)": "{:.1f}", "n° Battute (BT)": "{:.0f}", "n° B.Point": "{:.0f}"})

    # ===== BT NEGATIVA / POSITIVA / ESCLAMATIVA =====
    elif voce == "BREAK con BT. NEGATIVA":
        # ==========================================================
        # TABELLA RIDEFINITA (richiesta):
        # A Ranking (per % B.Point con Bt-)
        # B Team (12)
        # C % B.Point con Bt-  = BP_su_battuta_negativa / battute_negative
        #    BP: fine azione *p (casa) / ap (ospite) A FAVORE DEL BATTITORE
        # D % Bt-/Bt Tot = battute_negative / battute_totali
        # ==========================================================
        with engine.begin() as conn:
            rows = conn.execute(
                text("""
                    SELECT team_a, team_b, scout_text
                    FROM matches
                    WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        if not rows:
            st.info("Nessun dato nel range selezionato.")
            return

        def fix_team(name: str) -> str:
            return canonical_team(name)

        agg = {}
        def ensure(team: str):
            if team not in agg:
                agg[team] = {"Team": team, "neg_serves": 0, "neg_bp": 0, "tot_serves": 0}

        for r in rows:
            ta = fix_team(r.get("team_a") or "")
            tb = fix_team(r.get("team_b") or "")
            ensure(ta)
            ensure(tb)

            scout_text = r.get("scout_text") or ""
            lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
            scout_lines = [ln for ln in lines if ln[0] in ("*", "a")]

            rallies = []
            current = []

            for raw in scout_lines:
                c = code6(raw)
                if not c:
                    continue

                if is_serve(c):
                    if current:
                        rallies.append(current)
                    current = [c]
                    continue

                if not current:
                    continue

                current.append(c)

                if is_home_point(c) or is_away_point(c):
                    rallies.append(current)
                    current = []

            for rally in rallies:
                first = rally[0]
                home_served = first.startswith("*")
                away_served = first.startswith("a")

                home_point = any(is_home_point(x) for x in rally)
                away_point = any(is_away_point(x) for x in rally)

                if home_served:
                    agg[ta]["tot_serves"] += 1
                if away_served:
                    agg[tb]["tot_serves"] += 1

                is_neg = (len(first) >= 6 and first[5] == "-")
                if not is_neg:
                    continue

                if home_served:
                    agg[ta]["neg_serves"] += 1
                    if home_point:
                        agg[ta]["neg_bp"] += 1

                if away_served:
                    agg[tb]["neg_serves"] += 1
                    if away_point:
                        agg[tb]["neg_bp"] += 1

        df = pd.DataFrame(list(agg.values()))
        if df.empty:
            st.info("Nessun dato nel range selezionato.")
            return

        df["% B.Point con Bt-"] = (100.0 * df["neg_bp"] / df["neg_serves"].replace({0: pd.NA})).fillna(0.0)
        df["% Bt-/Bt Tot"] = (100.0 * df["neg_serves"] / df["tot_serves"].replace({0: pd.NA})).fillna(0.0)

        df = df.sort_values(by=["% B.Point con Bt-", "% Bt-/Bt Tot", "Team"], ascending=[False, False, True]).reset_index(drop=True)
        df.insert(0, "Ranking", range(1, len(df) + 1))

        out = df[["Ranking", "Team", "% B.Point con Bt-", "% Bt-/Bt Tot"]].copy()

        def highlight_perugia_row(row):
            is_perugia = "perugia" in str(row["Team"]).lower()
            style = "background-color: #fff3cd; font-weight: 800;" if is_perugia else ""
            return [style] * len(row)

        styled = (
            out.style
              .apply(highlight_perugia_row, axis=1)
              .format({"Ranking": "{:.0f}", "% B.Point con Bt-": "{:.1f}", "% Bt-/Bt Tot": "{:.1f}"})
              
              .set_properties(subset=[c for c in ["% B.Point con Bt-"] if c in out.columns], **{'background-color': '#e7f5ff', 'font-weight': '800'}).set_table_styles([
                  {"selector": "th", "props": [("font-size", "24px"), ("text-align", "left"), ("padding", "10px 12px")]},
                  {"selector": "td", "props": [("font-size", "23px"), ("padding", "10px 12px")]},
              ])
        )
        render_any_table(styled)
    elif voce == "BREAK con BT. POSITIVA":
        # ==========================================================
        # TABELLA POSITIVA (stesso modello della NEGATIVA):
        # A Ranking (per % B.Point con Bt+)
        # B Team (12)
        # C % B.Point con Bt+  = BP_su_battuta_positiva / battute_positive
        #    BP: fine azione *p (casa) / ap (ospite) A FAVORE DEL BATTITORE
        # D % Bt+/Bt Tot = battute_positive / battute_totali
        # ==========================================================
        with engine.begin() as conn:
            rows = conn.execute(
                text("""
                    SELECT team_a, team_b, scout_text
                    FROM matches
                    WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        if not rows:
            st.info("Nessun dato nel range selezionato.")
            return

        def fix_team(name: str) -> str:
            return canonical_team(name)

        agg = {}
        def ensure(team: str):
            if team not in agg:
                agg[team] = {"Team": team, "pos_serves": 0, "pos_bp": 0, "tot_serves": 0}

        for r in rows:
            ta = fix_team(r.get("team_a") or "")
            tb = fix_team(r.get("team_b") or "")
            ensure(ta)
            ensure(tb)

            scout_text = r.get("scout_text") or ""
            lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
            scout_lines = [ln for ln in lines if ln[0] in ("*", "a")]

            rallies = []
            current = []

            for raw in scout_lines:
                c = code6(raw)
                if not c:
                    continue

                if is_serve(c):
                    if current:
                        rallies.append(current)
                    current = [c]
                    continue

                if not current:
                    continue

                current.append(c)

                if is_home_point(c) or is_away_point(c):
                    rallies.append(current)
                    current = []

            for rally in rallies:
                first = rally[0]
                home_served = first.startswith("*")
                away_served = first.startswith("a")

                home_point = any(is_home_point(x) for x in rally)
                away_point = any(is_away_point(x) for x in rally)

                # total serves
                if home_served:
                    agg[ta]["tot_serves"] += 1
                if away_served:
                    agg[tb]["tot_serves"] += 1

                # positiva se 6° char del servizio è '+'
                is_pos = (len(first) >= 6 and first[5] == "+")
                if not is_pos:
                    continue

                if home_served:
                    agg[ta]["pos_serves"] += 1
                    # BP su battuta positiva: punto a favore del battitore (casa) -> *p
                    if home_point:
                        agg[ta]["pos_bp"] += 1

                if away_served:
                    agg[tb]["pos_serves"] += 1
                    # BP su battuta positiva: punto a favore del battitore (ospite) -> ap
                    if away_point:
                        agg[tb]["pos_bp"] += 1

        df = pd.DataFrame(list(agg.values()))
        if df.empty:
            st.info("Nessun dato nel range selezionato.")
            return

        df["% B.Point con Bt+"] = (100.0 * df["pos_bp"] / df["pos_serves"].replace({0: pd.NA})).fillna(0.0)
        df["% Bt+/Bt Tot"] = (100.0 * df["pos_serves"] / df["tot_serves"].replace({0: pd.NA})).fillna(0.0)

        df = df.sort_values(by=["% B.Point con Bt+", "% Bt+/Bt Tot", "Team"], ascending=[False, False, True]).reset_index(drop=True)
        df.insert(0, "Ranking", range(1, len(df) + 1))

        out = df[["Ranking", "Team", "% B.Point con Bt+", "% Bt+/Bt Tot"]].copy()

        def highlight_perugia_row(row):
            is_perugia = "perugia" in str(row["Team"]).lower()
            style = "background-color: #fff3cd; font-weight: 800;" if is_perugia else ""
            return [style] * len(row)

        styled = (
            out.style
              .apply(highlight_perugia_row, axis=1)
              .format({"Ranking": "{:.0f}", "% B.Point con Bt+": "{:.1f}", "% Bt+/Bt Tot": "{:.1f}"})
              
              .set_properties(subset=[c for c in ["% B.Point con Bt+"] if c in out.columns], **{'background-color': '#e7f5ff', 'font-weight': '800'}).set_table_styles([
                  {"selector": "th", "props": [("font-size", "24px"), ("text-align", "left"), ("padding", "10px 12px")]},
                  {"selector": "td", "props": [("font-size", "23px"), ("padding", "10px 12px")]},
              ])
        )
        render_any_table(styled)
    elif voce == "BREAK con BT. ESCLAMATIVA":
        # ==========================================================
        # TABELLA ESCLAMATIVA (solo '!'):
        # A Ranking (per % B.Point con Bt!)
        # B Team (12)
        # C % B.Point con Bt! = BP_su_battuta_esclamativa / battute_esclamative
        #    BP: fine azione *p (casa) / ap (ospite) a favore del battitore
        # D % Bt!/Bt Tot = battute_esclamative / battute_totali
        # ==========================================================
        with engine.begin() as conn:
            rows = conn.execute(
                text("""
                    SELECT team_a, team_b, scout_text
                    FROM matches
                    WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        if not rows:
            st.info("Nessun dato nel range selezionato.")
            return

        def fix_team(name: str) -> str:
            return canonical_team(name)

        agg = {}
        def ensure(team: str):
            if team not in agg:
                agg[team] = {"Team": team, "exc_serves": 0, "exc_bp": 0, "tot_serves": 0}

        for r in rows:
            ta = fix_team(r.get("team_a") or "")
            tb = fix_team(r.get("team_b") or "")
            ensure(ta)
            ensure(tb)

            scout_text = r.get("scout_text") or ""
            lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
            scout_lines = [ln for ln in lines if ln[0] in ("*", "a")]

            rallies = []
            current = []

            for raw in scout_lines:
                c = code6(raw)
                if not c:
                    continue

                if is_serve(c):
                    if current:
                        rallies.append(current)
                    current = [c]
                    continue

                if not current:
                    continue

                current.append(c)

                if is_home_point(c) or is_away_point(c):
                    rallies.append(current)
                    current = []

            for rally in rallies:
                first = rally[0]
                home_served = first.startswith("*")
                away_served = first.startswith("a")

                home_point = any(is_home_point(x) for x in rally)
                away_point = any(is_away_point(x) for x in rally)

                if home_served:
                    agg[ta]["tot_serves"] += 1
                if away_served:
                    agg[tb]["tot_serves"] += 1

                # esclamativa se 6° char del servizio è '!'
                is_exc = (len(first) >= 6 and first[5] == "!")
                if not is_exc:
                    continue

                if home_served:
                    agg[ta]["exc_serves"] += 1
                    if home_point:
                        agg[ta]["exc_bp"] += 1

                if away_served:
                    agg[tb]["exc_serves"] += 1
                    if away_point:
                        agg[tb]["exc_bp"] += 1

        df = pd.DataFrame(list(agg.values()))
        if df.empty:
            st.info("Nessun dato nel range selezionato.")
            return

        df["% B.Point con Bt!"] = (100.0 * df["exc_bp"] / df["exc_serves"].replace({0: pd.NA})).fillna(0.0)
        df["% Bt!/Bt Tot"] = (100.0 * df["exc_serves"] / df["tot_serves"].replace({0: pd.NA})).fillna(0.0)

        df = df.sort_values(by=["% B.Point con Bt!", "% Bt!/Bt Tot", "Team"], ascending=[False, False, True]).reset_index(drop=True)
        df.insert(0, "Ranking", range(1, len(df) + 1))

        out = df[["Ranking", "Team", "% B.Point con Bt!", "% Bt!/Bt Tot"]].copy()

        def highlight_perugia_row(row):
            is_perugia = "perugia" in str(row["Team"]).lower()
            style = "background-color: #fff3cd; font-weight: 800;" if is_perugia else ""
            return [style] * len(row)

        styled = (
            out.style
              .apply(highlight_perugia_row, axis=1)
              .format({"Ranking": "{:.0f}", "% B.Point con Bt!": "{:.1f}", "% Bt!/Bt Tot": "{:.1f}"})
              
              .set_properties(subset=[c for c in ["% B.Point con Bt!"] if c in out.columns], **{'background-color': '#e7f5ff', 'font-weight': '800'}).set_table_styles([
                  {"selector": "th", "props": [("font-size", "24px"), ("text-align", "left"), ("padding", "10px 12px")]},
                  {"selector": "td", "props": [("font-size", "23px"), ("padding", "10px 12px")]},
              ])
        )
        render_any_table(styled)


    elif voce == "BREAK con BT. 1/2 PUNTO":
        # ==========================================================
        # TABELLA 1/2 PUNTO (stesso modello di NEG / POS / !):
        # A Ranking (per % B.Point con Bt½)
        # B Team (12)
        # C % B.Point con Bt½ = BP_su_battuta_½ / battute_½
        #    BP: fine azione *p (casa) / ap (ospite) a favore del battitore
        # D % Bt½/Bt Tot = battute_½ / battute_totali
        # ==========================================================
        with engine.begin() as conn:
            rows = conn.execute(
                text("""
                    SELECT team_a, team_b, scout_text
                    FROM matches
                    WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        if not rows:
            st.info("Nessun dato nel range selezionato.")
            return

        def fix_team(name: str) -> str:
            return canonical_team(name)

        agg = {}
        def ensure(team: str):
            if team not in agg:
                agg[team] = {"Team": team, "half_serves": 0, "half_bp": 0, "tot_serves": 0}

        for r in rows:
            ta = fix_team(r.get("team_a") or "")
            tb = fix_team(r.get("team_b") or "")
            ensure(ta)
            ensure(tb)

            scout_text = r.get("scout_text") or ""
            lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
            scout_lines = [ln for ln in lines if ln[0] in ("*", "a")]

            rallies = []
            current = []

            for raw in scout_lines:
                c = code6(raw)
                if not c:
                    continue

                if is_serve(c):
                    if current:
                        rallies.append(current)
                    current = [c]
                    continue

                if not current:
                    continue

                current.append(c)

                if is_home_point(c) or is_away_point(c):
                    rallies.append(current)
                    current = []

            for rally in rallies:
                first = rally[0]
                home_served = first.startswith("*")
                away_served = first.startswith("a")

                home_point = any(is_home_point(x) for x in rally)
                away_point = any(is_away_point(x) for x in rally)

                if home_served:
                    agg[ta]["tot_serves"] += 1
                if away_served:
                    agg[tb]["tot_serves"] += 1

                # 1/2 punto: nel tuo scout è '/' nel 6° carattere (es: *06SQ/).
                # Nessun fallback: usiamo '/' come da codifica.
                is_half = (len(first) >= 6 and first[5] == "/")
                if not is_half:
                    continue

                if home_served:
                    agg[ta]["half_serves"] += 1
                    if home_point:
                        agg[ta]["half_bp"] += 1

                if away_served:
                    agg[tb]["half_serves"] += 1
                    if away_point:
                        agg[tb]["half_bp"] += 1

        df = pd.DataFrame(list(agg.values()))
        if df.empty:
            st.info("Nessun dato nel range selezionato.")
            return

        df["% B.Point con Bt½"] = (100.0 * df["half_bp"] / df["half_serves"].replace({0: pd.NA})).fillna(0.0)
        df["% Bt½/Bt Tot"] = (100.0 * df["half_serves"] / df["tot_serves"].replace({0: pd.NA})).fillna(0.0)

        df = df.sort_values(by=["% B.Point con Bt½", "% Bt½/Bt Tot", "Team"], ascending=[False, False, True]).reset_index(drop=True)
        df.insert(0, "Ranking", range(1, len(df) + 1))

        out = df[["Ranking", "Team", "% B.Point con Bt½", "% Bt½/Bt Tot"]].copy()

        def highlight_perugia_row(row):
            is_perugia = "perugia" in str(row["Team"]).lower()
            style = "background-color: #fff3cd; font-weight: 800;" if is_perugia else ""
            return [style] * len(row)

        styled = (
            out.style
              .apply(highlight_perugia_row, axis=1)
              .format({"Ranking": "{:.0f}", "% B.Point con Bt½": "{:.1f}", "% Bt½/Bt Tot": "{:.1f}"})
              
              .set_properties(subset=[c for c in ["% B.Point con Bt½", "% B.Point con Bt1/2"] if c in out.columns], **{'background-color': '#e7f5ff', 'font-weight': '800'}).set_table_styles([
                  {"selector": "th", "props": [("font-size", "24px"), ("text-align", "left"), ("padding", "10px 12px")]},
                  {"selector": "td", "props": [("font-size", "23px"), ("padding", "10px 12px")]},
              ])
        )
        render_any_table(styled)


    elif voce == "BT punto/errore/ratio":
        # ==========================================================
        # BT punto/errore/ratio (richiesta):
        # COL A: Teams
        # COL B: % Bt Punto (#) su Tot Battute
        # COL C: % Bt Errore (=) su Tot Battute
        # COL D: Ratio Errore/Punto = Tot '=' / Tot '#'
        #
        # Definizioni (sul servizio code6):
        # - Battuta = SQ o SM (is_serve)
        # - Punto = 6° carattere '#': *06SQ#
        # - Errore = 6° carattere '=': *06SQ=
        # ==========================================================
        with engine.begin() as conn:
            rows = conn.execute(
                text("""
                    SELECT team_a, team_b, scout_text
                    FROM matches
                    WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        if not rows:
            st.info("Nessun dato nel range selezionato.")
            return

        def fix_team(name: str) -> str:
            return canonical_team(name)

        agg = {}
        def ensure(team: str):
            if team not in agg:
                agg[team] = {"Teams": team, "tot_serves": 0, "bt_punto": 0, "bt_errore": 0}

        for r in rows:
            ta = fix_team(r.get("team_a") or "")
            tb = fix_team(r.get("team_b") or "")
            ensure(ta); ensure(tb)

            scout_text = r.get("scout_text") or ""
            lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
            scout_lines = [ln for ln in lines if ln[0] in ("*", "a")]

            rallies = []
            current = []

            for raw in scout_lines:
                c = code6(raw)
                if not c:
                    continue

                if is_serve(c):
                    if current:
                        rallies.append(current)
                    current = [c]
                    continue

                if not current:
                    continue

                current.append(c)

                if is_home_point(c) or is_away_point(c):
                    rallies.append(current)
                    current = []

            for rally in rallies:
                first = rally[0]
                home_served = first.startswith("*")
                away_served = first.startswith("a")

                if home_served:
                    agg[ta]["tot_serves"] += 1
                    if len(first) >= 6 and first[5] == "#":
                        agg[ta]["bt_punto"] += 1
                    elif len(first) >= 6 and first[5] == "=":
                        agg[ta]["bt_errore"] += 1

                if away_served:
                    agg[tb]["tot_serves"] += 1
                    if len(first) >= 6 and first[5] == "#":
                        agg[tb]["bt_punto"] += 1
                    elif len(first) >= 6 and first[5] == "=":
                        agg[tb]["bt_errore"] += 1

        df = pd.DataFrame(list(agg.values()))
        if df.empty:
            st.info("Nessun dato nel range selezionato.")
            return

        df["% Bt Punto/Tot Bt"] = (100.0 * df["bt_punto"] / df["tot_serves"].replace({0: pd.NA})).fillna(0.0)
        df["% Bt Errore/Tot Bt"] = (100.0 * df["bt_errore"] / df["tot_serves"].replace({0: pd.NA})).fillna(0.0)
        df["Ratio Errore/Punto"] = (df["bt_errore"] / df["bt_punto"].replace({0: pd.NA})).fillna(0.0)
        show_details = st.checkbox("Mostra colonne di controllo (Tot Bt, Bt#, Bt=)", value=False)

        df["Tot Bt"] = df["tot_serves"]
        df["Bt#"] = df["bt_punto"]
        df["Bt="] = df["bt_errore"]

        base_cols = ["Teams", "% Bt Punto/Tot Bt", "% Bt Errore/Tot Bt", "Ratio Errore/Punto"]
        detail_cols = ["Tot Bt", "Bt#", "Bt="] if show_details else []

        out = df[base_cols + detail_cols].copy()
        out = out.sort_values(by=["% Bt Punto/Tot Bt", "Teams"], ascending=[False, True]).reset_index(drop=True)


        def highlight_perugia_row(row):
            is_perugia = "perugia" in str(row["Teams"]).lower()
            style = "background-color: #fff3cd; font-weight: 800;" if is_perugia else ""
            return [style] * len(row)

        styled = (
            out.style
              .apply(highlight_perugia_row, axis=1)
              .format({
                  "% Bt Punto/Tot Bt": "{:.1f}",
                  "% Bt Errore/Tot Bt": "{:.1f}",
                  "Ratio Errore/Punto": "{:.2f}",
              })
              
              .set_properties(subset=[c for c in ["% Bt Punto/Tot Bt"] if c in out.columns], **{'background-color': '#e7f5ff', 'font-weight': '800'}).set_table_styles([
                  {"selector": "th", "props": [("font-size", "24px"), ("text-align", "left"), ("padding", "10px 12px")]},
                  {"selector": "td", "props": [("font-size", "23px"), ("padding", "10px 12px")]},
              ])
        )
        render_any_table(styled)


    elif voce == "Confronto TEAM":
        # ==========================================================
        # Confronto TEAM (max 4 squadre) nel range giornate selezionato
        # Colonne:
        # TEAM | % BREAK TOTALE | BREAK GIOCATO | Bt- | Bt! | Bt+ | 1/2 | BT punto/errore/ratio
        # Definizioni (sul servizio code6):
        # - Battuta = is_serve(code6) -> SQ o SM
        # - Break totale % = punti a favore del battitore / battute totali
        # - Break giocato % = punti a favore del battitore / battute con segno tra (-,+,!,/)
        # - Bt-: % BP con Bt- = punti a favore del battitore / battute con '-' (6° char)
        # - Bt!: % BP con Bt! = punti a favore del battitore / battute con '!' (6° char)
        # - Bt+: % BP con Bt+ = punti a favore del battitore / battute con '+' (6° char)
        # - 1/2: % BP con Bt/ = punti a favore del battitore / battute con '/' (6° char)
        # - BT punto/errore/ratio: "P% / E% / R" su totale battute (P=#, E==, R=E/P)
        # ==========================================================

        with engine.begin() as conn:
            teams_raw = conn.execute(text("""
                SELECT DISTINCT squadra FROM (
                    SELECT team_a AS squadra FROM matches
                    WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                    UNION
                    SELECT team_b AS squadra FROM matches
                    WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                )
                WHERE squadra IS NOT NULL AND TRIM(squadra) <> ''
                ORDER BY squadra
            """), {"from_round": int(from_round), "to_round": int(to_round)}).mappings().all()

        def fix_team(name: str) -> str:
            n = " ".join((name or "").split())
            if n.lower().startswith("gas sales bluenergy p"):
                return "Gas Sales Bluenergy Piacenza"
            if "grottazzolina" in n.lower():
                return "Yuasa Battery Grottazzolina"
            return n

        teams = sorted({fix_team(r["squadra"]) for r in teams_raw if r.get("squadra")})

        selected = st.multiselect(
            "Seleziona fino a 4 squadre",
            options=teams,
            default=[],
            max_selections=4,
        )

        if not selected:
            st.info("Seleziona 1–4 squadre per vedere il confronto.")
            return

        # carica match nel range
        with engine.begin() as conn:
            matches = conn.execute(
                text("""
                    SELECT team_a, team_b, scout_text
                    FROM matches
                    WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        if not matches:
            st.info("Nessun match nel range selezionato.")
            return

        # aggregatori per team
        agg = {}
        def ensure(team: str):
            if team not in agg:
                agg[team] = {
                    "TEAM": team,
                    "tot_serves": 0, "tot_bp": 0,
                    "g_att": 0, "g_bp": 0,     # giocato (- + ! /)
                    "neg_att": 0, "neg_bp": 0,
                    "exc_att": 0, "exc_bp": 0,
                    "pos_att": 0, "pos_bp": 0,
                    "half_att": 0, "half_bp": 0,
                    "pt_cnt": 0, "err_cnt": 0,  # # and =
                }

        # helper per parse rally da scout_text
        def rallies_from_scout_text(scout_text: str):
            if not scout_text:
                return []
            lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
            scout_lines = [ln for ln in lines if ln[0] in ("*", "a")]

            rallies = []
            current = []
            for raw in scout_lines:
                c = code6(raw)
                if not c:
                    continue
                if is_serve(c):
                    if current:
                        rallies.append(current)
                    current = [c]
                    continue
                if not current:
                    continue
                current.append(c)
                if is_home_point(c) or is_away_point(c):
                    rallies.append(current)
                    current = []
            return rallies

        def serve_sign(c6: str) -> str:
            return c6[5] if len(c6) >= 6 else ""

        for m in matches:
            ta = fix_team(m.get("team_a") or "")
            tb = fix_team(m.get("team_b") or "")
            # assicurati in agg solo se selezionate (ottimizziamo)
            if ta not in selected and tb not in selected:
                continue
            ensure(ta); ensure(tb)

            for rally in rallies_from_scout_text(m.get("scout_text") or ""):
                first = rally[0]
                if not is_serve(first):
                    continue

                home_served = first.startswith("*")
                away_served = first.startswith("a")

                home_point = any(is_home_point(x) for x in rally)
                away_point = any(is_away_point(x) for x in rally)

                # team che batte e "bp" (punto a favore del battitore)
                if home_served:
                    team = ta
                    bp = 1 if home_point else 0
                elif away_served:
                    team = tb
                    bp = 1 if away_point else 0
                else:
                    continue

                if team not in selected:
                    continue

                sgn = serve_sign(first)

                agg[team]["tot_serves"] += 1
                agg[team]["tot_bp"] += bp

                # Break giocato: solo - + ! /
                if sgn in ("-", "+", "!", "/"):
                    agg[team]["g_att"] += 1
                    agg[team]["g_bp"] += bp

                # per segno
                if sgn == "-":
                    agg[team]["neg_att"] += 1
                    agg[team]["neg_bp"] += bp
                elif sgn == "!":
                    agg[team]["exc_att"] += 1
                    agg[team]["exc_bp"] += bp
                elif sgn == "+":
                    agg[team]["pos_att"] += 1
                    agg[team]["pos_bp"] += bp
                elif sgn == "/":
                    agg[team]["half_att"] += 1
                    agg[team]["half_bp"] += bp

                # BT punto/errore/ratio
                if sgn == "#":
                    agg[team]["pt_cnt"] += 1
                elif sgn == "=":
                    agg[team]["err_cnt"] += 1

        df = pd.DataFrame([agg[t] for t in selected])
        if df.empty:
            st.info("Nessun dato per le squadre selezionate nel range.")
            return

        def safe_pct(num, den):
            return float(num) / float(den) * 100.0 if den else 0.0

        # calcoli finali
        df["% BREAK TOTALE"] = df.apply(lambda r: safe_pct(r["tot_bp"], r["tot_serves"]), axis=1)
        df["BREAK GIOCATO"] = df.apply(lambda r: safe_pct(r["g_bp"], r["g_att"]), axis=1)

        df["BREAK con BT. NEGATIVA"] = df.apply(lambda r: safe_pct(r["neg_bp"], r["neg_att"]), axis=1)
        df["BREAK con BT. ESCLAMATIVA"] = df.apply(lambda r: safe_pct(r["exc_bp"], r["exc_att"]), axis=1)
        df["BREAK con BT. POSITIVA"] = df.apply(lambda r: safe_pct(r["pos_bp"], r["pos_att"]), axis=1)
        df["1/2 PUNTO"] = df.apply(lambda r: safe_pct(r["half_bp"], r["half_att"]), axis=1)

        df["BT Punto %"] = df.apply(lambda r: safe_pct(r["pt_cnt"], r["tot_serves"]), axis=1)
        df["BT Errore %"] = df.apply(lambda r: safe_pct(r["err_cnt"], r["tot_serves"]), axis=1)
        df["BT Ratio (=/#)"] = df.apply(lambda r: (float(r["err_cnt"]) / float(r["pt_cnt"]) if r["pt_cnt"] else 0.0), axis=1)


        out = df[[
            "TEAM",
            "% BREAK TOTALE",
            "BREAK GIOCATO",
            "BREAK con BT. NEGATIVA",
            "BREAK con BT. ESCLAMATIVA",
            "BREAK con BT. POSITIVA",
            "1/2 PUNTO",
            "BT Punto %",
            "BT Errore %",
            "BT Ratio (=/#)",
        ]].copy()

        # formatting + highlight Perugia
        def highlight_perugia_row(row):
            is_perugia = "perugia" in str(row["TEAM"]).lower()
            style = "background-color: #fff3cd; font-weight: 800;" if is_perugia else ""
            return [style] * len(row)

        styled = (
            out.style
              .apply(highlight_perugia_row, axis=1)
              .format({
                  "% BREAK TOTALE": "{:.1f}",
                  "BREAK GIOCATO": "{:.1f}",
                  "BREAK con BT. NEGATIVA": "{:.1f}",
                  "BREAK con BT. ESCLAMATIVA": "{:.1f}",
                  "BREAK con BT. POSITIVA": "{:.1f}",
                  "1/2 PUNTO": "{:.1f}",
                                "BT Punto %": "{:.1f}",
                  "BT Errore %": "{:.1f}",
                  "BT Ratio (=/#)": "{:.2f}",
})
              .set_table_styles([
                  {"selector": "th", "props": [("font-size", "24px"), ("text-align", "left"), ("padding", "10px 12px")]},
                  {"selector": "td", "props": [("font-size", "23px"), ("padding", "10px 12px")]},
              ])
        )
        render_any_table(styled)

    elif voce == "GRAFICI":
        # ==========================================================
        # GRAFICI (max 4 squadre) nel range giornate selezionato
        # Opzioni:
        # 1) Distribuzione BP per tipo di Battuta
        # 2) Distribuzione tipo di battuta (su Tot battute)
        #
        # Codifica (6° carattere del servizio):
        #   '-' negativa, '!' esclamativa, '+' positiva, '/' mezzo punto,
        #   '#' punto, '=' errore
        # ==========================================================
        import matplotlib.pyplot as plt

        with engine.begin() as conn:
            teams_raw = conn.execute(text("""
                SELECT DISTINCT squadra FROM (
                    SELECT team_a AS squadra FROM matches
                    WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                    UNION
                    SELECT team_b AS squadra FROM matches
                    WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                )
                WHERE squadra IS NOT NULL AND TRIM(squadra) <> ''
                ORDER BY squadra
            """), {"from_round": int(from_round), "to_round": int(to_round)}).mappings().all()

        def fix_team(name: str) -> str:
            n = " ".join((name or "").split())
            if n.lower().startswith("gas sales bluenergy p"):
                return "Gas Sales Bluenergy Piacenza"
            if "grottazzolina" in n.lower():
                return "Yuasa Battery Grottazzolina"
            return n

        teams = sorted({fix_team(r["squadra"]) for r in teams_raw if r.get("squadra")})

        selected = st.multiselect(
            "Seleziona fino a 4 squadre",
            options=teams,
            default=[],
            max_selections=4,
            key="grafici_teams",
        )

        option = st.radio(
            "Seleziona grafico",
            [
                "Vedi distribuzione BP per tipo di Battuta",
                "Vedi distribuzione tipo di battuta",
            ],
            index=0,
            key="grafici_option",
        )

        # --- grafici compatti (2 colonne) ---
        FIGSIZE = (3.4, 3.4)   # ancora più compatto
        DPI = 120
        LABEL_FONTSIZE = 8
        TITLE_FONTSIZE = 10

        if not selected:
            st.info("Seleziona 1–4 squadre per vedere i grafici.")
            return

        with engine.begin() as conn:
            matches = conn.execute(
                text("""
                    SELECT team_a, team_b, scout_text
                    FROM matches
                    WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        if not matches:
            st.info("Nessun match nel range selezionato.")
            return

        def rallies_from_scout_text(scout_text: str):
            if not scout_text:
                return []
            lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
            scout_lines = [ln for ln in lines if ln[0] in ("*", "a")]

            rallies = []
            current = []
            for raw in scout_lines:
                c = code6(raw)
                if not c:
                    continue
                if is_serve(c):
                    if current:
                        rallies.append(current)
                    current = [c]
                    continue
                if not current:
                    continue
                current.append(c)
                if is_home_point(c) or is_away_point(c):
                    rallies.append(current)
                    current = []
            return rallies

        def serve_sign(c6: str) -> str:
            return c6[5] if len(c6) >= 6 else ""

        stats = {}
        for t in selected:
            stats[t] = {
                "tot_serves": 0,
                "serve_counts": {"-": 0, "!": 0, "+": 0, "/": 0, "#": 0, "=": 0},
                "bp_counts": {"-": 0, "!": 0, "+": 0, "/": 0, "#": 0},
            }

        for m in matches:
            ta = fix_team(m.get("team_a") or "")
            tb = fix_team(m.get("team_b") or "")
            if ta not in selected and tb not in selected:
                continue

            for rally in rallies_from_scout_text(m.get("scout_text") or ""):
                first = rally[0]
                if not is_serve(first):
                    continue

                home_served = first.startswith("*")
                away_served = first.startswith("a")

                home_point = any(is_home_point(x) for x in rally)
                away_point = any(is_away_point(x) for x in rally)

                if home_served:
                    team = ta
                    bp = 1 if home_point else 0
                elif away_served:
                    team = tb
                    bp = 1 if away_point else 0
                else:
                    continue

                if team not in selected:
                    continue

                sgn = serve_sign(first)
                stats[team]["tot_serves"] += 1
                if sgn in stats[team]["serve_counts"]:
                    stats[team]["serve_counts"][sgn] += 1
                if sgn in stats[team]["bp_counts"]:
                    stats[team]["bp_counts"][sgn] += bp

        label_map_bp = {"-": "Neg", "!": "!", "+": "+", "/": "½", "#": "#"}
        label_map_all = {"-": "Neg", "!": "!", "+": "+", "/": "½", "#": "#", "=": "="}

        def counts_df(keys, labels, values):
            return pd.DataFrame({"Tipo": labels, "Conteggio": values})

        cols = st.columns(2)
        for i, t in enumerate(selected):
            col = cols[i % 2]
            with col:
                st.markdown(f"### {t}")

                if option == "Vedi distribuzione BP per tipo di Battuta":
                    keys = ["-", "!", "+", "/", "#"]
                    labels = [label_map_bp[k] for k in keys]
                    values = [stats[t]["bp_counts"].get(k, 0) for k in keys]
                    tot_bp = sum(values)

                    if tot_bp == 0:
                        st.info("Nessun Break Point nel range selezionato.")
                    else:
                        fig, ax = plt.subplots(figsize=FIGSIZE, dpi=DPI)
                        ax.pie(values, labels=labels, autopct="%1.1f%%", textprops={"fontsize": LABEL_FONTSIZE})
                        ax.set_title("BP per battuta", fontsize=TITLE_FONTSIZE)
                        st.pyplot(fig, clear_figure=True)

                        # Chicca: totale + dettaglio conteggi in expander
                        st.caption(f"Tot BP: **{tot_bp}** | Tot battute: **{stats[t]['tot_serves']}**")
                        with st.expander("Dettaglio conteggi BP", expanded=False):
                            st.dataframe(counts_df(keys, labels, values), width="stretch", hide_index=True)

                else:
                    keys = ["-", "!", "+", "/", "#", "="]
                    labels = [label_map_all[k] for k in keys]
                    values = [stats[t]["serve_counts"].get(k, 0) for k in keys]
                    tot_serves = sum(values)

                    if tot_serves == 0:
                        st.info("Nessuna battuta nel range selezionato.")
                    else:
                        fig, ax = plt.subplots(figsize=FIGSIZE, dpi=DPI)
                        ax.pie(values, labels=labels, autopct="%1.1f%%", textprops={"fontsize": LABEL_FONTSIZE})
                        ax.set_title("Distribuzione battute", fontsize=TITLE_FONTSIZE)
                        st.pyplot(fig, clear_figure=True)

                        # Chicca: totale + dettaglio conteggi in expander
                        st.caption(f"Tot battute: **{tot_serves}**")
                        with st.expander("Dettaglio conteggi battute", expanded=False):
                            st.dataframe(counts_df(keys, labels, values), width="stretch", hide_index=True)

            if i % 2 == 1 and i < len(selected) - 1:
                st.divider()
                cols = st.columns(2)



# =========================
# UI: GRAFICI 4 QUADRANTI (Trend per giornata, X/Y selezionabili)
# =========================
def render_grafici_4_quadranti():
    st.header("GRAFICI 4 Quadranti – Trend")

    from_round, to_round, _range_label = get_selected_match_range()
    st.caption(f"📌 Range: {_range_label}")

    if from_round > to_round:
        st.error("Range non valido: 'Da giornata' deve essere <= 'A giornata'.")
        st.stop()

    # ---------------------------------------------------------
    # Metriche disponibili (per match / per squadra)
    # Ogni metrica definisce (num_col, den_col) per HOME e AWAY.
    # ---------------------------------------------------------

    METRICS = {
        # ------------------------
        # SIDE OUT (percentuali)
        # ------------------------
        "Side Out TOTALE": {
            "home_num": "so_home_wins", "home_den": "so_home_attempts",
            "away_num": "so_away_wins", "away_den": "so_away_attempts",
            "kind": "pct",
        },
        "Side Out SPIN": {
            "home_num": "so_spin_home_wins", "home_den": "so_spin_home_attempts",
            "away_num": "so_spin_away_wins", "away_den": "so_spin_away_attempts",
            "kind": "pct",
        },
        "Side Out FLOAT": {
            "home_num": "so_float_home_wins", "home_den": "so_float_home_attempts",
            "away_num": "so_float_away_wins", "away_den": "so_float_away_attempts",
            "kind": "pct",
        },
        # quota di Side Out diretto sul totale Side Out (share)
        "Side Out DIRETTO": {
            "home_num": "so_dir_home_wins", "home_den": "so_home_wins",
            "away_num": "so_dir_away_wins", "away_den": "so_away_wins",
            "kind": "pct",
        },
        "Side Out GIOCATO": {
            "home_num": "so_play_home_wins", "home_den": "so_play_home_attempts",
            "away_num": "so_play_away_wins", "away_den": "so_play_away_attempts",
            "kind": "pct",
        },
        "Side Out con RICE BUONA": {
            "home_num": "so_good_home_wins", "home_den": "so_good_home_attempts",
            "away_num": "so_good_away_wins", "away_den": "so_good_away_attempts",
            "kind": "pct",
        },
        "Side Out con RICE ESCLAMATIVA": {
            "home_num": "so_exc_home_wins", "home_den": "so_exc_home_attempts",
            "away_num": "so_exc_away_wins", "away_den": "so_exc_away_attempts",
            "kind": "pct",
        },
        "Side Out con RICE NEGATIVA": {
            "home_num": "so_neg_home_wins", "home_den": "so_neg_home_attempts",
            "away_num": "so_neg_away_wins", "away_den": "so_neg_away_attempts",
            "kind": "pct",
        },

        # ------------------------
        # BREAK (percentuali)
        # ------------------------
        "BREAK TOTALE": {
            "home_num": "bp_home_wins", "home_den": "bp_home_attempts",
            "away_num": "bp_away_wins", "away_den": "bp_away_attempts",
            "kind": "pct",
        },
        "BREAK GIOCATO": {
            "home_num": "bp_play_home_wins", "home_den": "bp_play_home_attempts",
            "away_num": "bp_play_away_wins", "away_den": "bp_play_away_attempts",
            "kind": "pct",
        },
        "BREAK con BT. NEGATIVA": {
            "home_num": "bp_neg_home_wins", "home_den": "bp_neg_home_attempts",
            "away_num": "bp_neg_away_wins", "away_den": "bp_neg_away_attempts",
            "kind": "pct",
        },
        "BREAK con BT. ESCLAMATIVA": {
            "home_num": "bp_exc_home_wins", "home_den": "bp_exc_home_attempts",
            "away_num": "bp_exc_away_wins", "away_den": "bp_exc_away_attempts",
            "kind": "pct",
        },
        "BREAK con BT. POSITIVA": {
            "home_num": "bp_pos_home_wins", "home_den": "bp_pos_home_attempts",
            "away_num": "bp_pos_away_wins", "away_den": "bp_pos_away_attempts",
            "kind": "pct",
        },
        "BREAK con BT. 1/2 PUNTO": {
            "home_num": "bp_half_home_wins", "home_den": "bp_half_home_attempts",
            "away_num": "bp_half_away_wins", "away_den": "bp_half_away_attempts",
            "kind": "pct",
        },

        # ------------------------
        # BATTUTA – punti/errori/ratio (per battuta)
        # ------------------------
        "BT punto": {
            "home_num": "bt_home_ace", "home_den": "bp_home_attempts",
            "away_num": "bt_away_ace", "away_den": "bp_away_attempts",
            "kind": "pct",
        },
        "Bt errore": {
            "home_num": "bt_home_err", "home_den": "bp_home_attempts",
            "away_num": "bt_away_err", "away_den": "bp_away_attempts",
            "kind": "pct",
        },
        "Bt errore/punto": {
            "home_num": "bt_home_err", "home_den": "bt_home_ace",
            "away_num": "bt_away_err", "away_den": "bt_away_ace",
            "kind": "ratio",
        },

        # ------------------------
        # EFFICIENZE FONDAMENTALI
        # (per ora come placeholder: le agganciamo nella prossima iterazione
        #  a contatori per match o a parsing scout_text per giornata)
        # ------------------------
        "Eff. Battuta": {"kind": "eff_battuta"},
        "Eff. Ricezione": {"kind": "eff_ricezione"},
        "Eff. Attacco": {"kind": "eff_attacco"},
        "Eff. Muro": {"kind": "eff_muro"},
        "Eff. Difesa": {"kind": "eff_difesa"},
    }


    col1, col2 = st.columns(2)
    with col1:
        x_metric = st.selectbox("Ascisse (X)", list(METRICS.keys()), index=list(METRICS.keys()).index("Side Out TOTALE"), key="g4q_x")
    with col2:
        y_metric = st.selectbox("Ordinate (Y)", list(METRICS.keys()), index=list(METRICS.keys()).index("BREAK TOTALE"), key="g4q_y")
        # Selezione Efficienze (fondamentali): non ancora disponibili nel grafico per giornata
        if METRICS.get(x_metric, {}).get("kind") == "todo" or METRICS.get(y_metric, {}).get("kind") == "todo":
            st.warning("Le voci Eff. (fondamentali) saranno disponibili a breve nel grafico. Per ora scegli un indice Side Out / Break / BT.")
            return

    # Carico match nel range
    with engine.begin() as conn:
        # Colonne presenti nel DB (evita crash se mancano colonne in alcune installazioni)
        cols = conn.execute(text("PRAGMA table_info(matches)")).mappings().all()
        present = {c["name"] for c in cols}

        def _c(col: str, alias: str, numeric: bool = True):
            """Return a safe SELECT expression for a possibly-missing column."""
            if col in present:
                if numeric:
                    return f"COALESCE({col},0) AS {alias}"
                return f"COALESCE({col},'') AS {alias}"
            return ("0" if numeric else "''") + f" AS {alias}"

        select_cols = [
            "phase",
            "round_number",
            "team_a",
            "team_b",
            _c("scout_text", "scout_text", numeric=False),

            _c("so_home_attempts", "so_home_attempts"),
            _c("so_home_wins", "so_home_wins"),
            _c("so_away_attempts", "so_away_attempts"),
            _c("so_away_wins", "so_away_wins"),

            _c("bp_home_attempts", "bp_home_attempts"),
            _c("bp_home_wins", "bp_home_wins"),
            _c("bp_away_attempts", "bp_away_attempts"),
            _c("bp_away_wins", "bp_away_wins"),

            _c("so_spin_home_attempts", "so_spin_home_attempts"),
            _c("so_spin_home_wins", "so_spin_home_wins"),
            _c("so_spin_away_attempts", "so_spin_away_attempts"),
            _c("so_spin_away_wins", "so_spin_away_wins"),

            _c("so_float_home_attempts", "so_float_home_attempts"),
            _c("so_float_home_wins", "so_float_home_wins"),
            _c("so_float_away_attempts", "so_float_away_attempts"),
            _c("so_float_away_wins", "so_float_away_wins"),

            _c("bp_play_home_attempts", "bp_play_home_attempts"),
            _c("bp_play_home_wins", "bp_play_home_wins"),
            _c("bp_play_away_attempts", "bp_play_away_attempts"),
            _c("bp_play_away_wins", "bp_play_away_wins"),

            _c("so_play_home_attempts", "so_play_home_attempts"),
            _c("so_play_home_wins", "so_play_home_wins"),
            _c("so_play_away_attempts", "so_play_away_attempts"),
            _c("so_play_away_wins", "so_play_away_wins"),

            _c("so_dir_home_attempts", "so_dir_home_attempts"),
            _c("so_dir_home_wins", "so_dir_home_wins"),
            _c("so_dir_away_attempts", "so_dir_away_attempts"),
            _c("so_dir_away_wins", "so_dir_away_wins"),

            _c("so_good_home_attempts", "so_good_home_attempts"),
            _c("so_good_home_wins", "so_good_home_wins"),
            _c("so_good_away_attempts", "so_good_away_attempts"),
            _c("so_good_away_wins", "so_good_away_wins"),

            _c("so_exc_home_attempts", "so_exc_home_attempts"),
            _c("so_exc_home_wins", "so_exc_home_wins"),
            _c("so_exc_away_attempts", "so_exc_away_attempts"),
            _c("so_exc_away_wins", "so_exc_away_wins"),

            _c("so_neg_home_attempts", "so_neg_home_attempts"),
            _c("so_neg_home_wins", "so_neg_home_wins"),
            _c("so_neg_away_attempts", "so_neg_away_attempts"),
            _c("so_neg_away_wins", "so_neg_away_wins"),

            _c("bp_neg_home_attempts", "bp_neg_home_attempts"),
            _c("bp_neg_home_wins", "bp_neg_home_wins"),
            _c("bp_neg_away_attempts", "bp_neg_away_attempts"),
            _c("bp_neg_away_wins", "bp_neg_away_wins"),

            _c("bp_exc_home_attempts", "bp_exc_home_attempts"),
            _c("bp_exc_home_wins", "bp_exc_home_wins"),
            _c("bp_exc_away_attempts", "bp_exc_away_attempts"),
            _c("bp_exc_away_wins", "bp_exc_away_wins"),

            _c("bp_pos_home_attempts", "bp_pos_home_attempts"),
            _c("bp_pos_home_wins", "bp_pos_home_wins"),
            _c("bp_pos_away_attempts", "bp_pos_away_attempts"),
            _c("bp_pos_away_wins", "bp_pos_away_wins"),

            _c("bp_half_home_attempts", "bp_half_home_attempts"),
            _c("bp_half_home_wins", "bp_half_home_wins"),
            _c("bp_half_away_attempts", "bp_half_away_attempts"),
            _c("bp_half_away_wins", "bp_half_away_wins"),

            _c("bt_home_ace", "bt_home_ace"),
            _c("bt_home_err", "bt_home_err"),
            _c("bt_away_ace", "bt_away_ace"),
            _c("bt_away_err", "bt_away_err"),
        ]

        sql = f"""
            SELECT {', '.join(select_cols)}
            FROM matches
            WHERE {match_order_sql('matches')} BETWEEN :from_round AND :to_round
            ORDER BY {match_order_sql('matches')} ASC
        """
        rows = conn.execute(text(sql), {"from_round": int(from_round), "to_round": int(to_round)}).mappings().all()

    if not rows:
        st.info("Nessun match nel range selezionato.")
        return

    dfm = pd.DataFrame(rows)
    dfm["team_a_city"] = dfm["team_a"].apply(lambda x: canonical_team(fix_team_name(x)))
    dfm["team_b_city"] = dfm["team_b"].apply(lambda x: canonical_team(fix_team_name(x)))

    all_teams = sorted(set(dfm["team_a_city"].dropna().astype(str)) | set(dfm["team_b_city"].dropna().astype(str)))
    default_sel = [t for t in ["PERUGIA", "TRENTO"] if t in all_teams] or all_teams[:2]

    selected = st.multiselect(
        "Squadre (seleziona 1–6)",
        options=all_teams,
        default=default_sel,
        max_selections=6,
        key="g4q_teams",
    )
    if len(selected) < 1:
        st.info("Seleziona almeno 1 squadra.")
        return

    def _round_label(ph, rn):
        ph = str(ph or "")
        rn = int(rn or 0)
        if ph in ("A", "R"):
            return f"{ph}{rn:02d}"
        return f"{ph}{rn}"

    def _pct(num, den):
        return (100.0 * float(num) / float(den)) if den else 0.0

    def build_points(metric_key: str) -> pd.DataFrame:
        mdef = METRICS[metric_key]
        kind = mdef.get("kind")

        def _round_label(ph, rn):
            ph = str(ph or "")
            rn = int(rn or 0)
            if ph in ("A", "R"):
                return f"{ph}{rn:02d}"
            return f"{ph}{rn}"

        def _pct(num, den):
            return (100.0 * float(num) / float(den)) if den else 0.0

        def parse_rallies(scout_text: str):
            if not scout_text:
                return []
            lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
            scout_lines = [ln for ln in lines if ln[0] in ("*", "a")]
            rallies = []
            current = []
            for raw in scout_lines:
                c = code6(raw)
                if not c:
                    continue
                if is_serve(c):
                    if current:
                        rallies.append(current)
                    current = [c]
                    continue
                if not current:
                    continue
                current.append(c)
                if is_home_point(c) or is_away_point(c):
                    rallies.append(current)
                    current = []
            return rallies

        # -------------------------------------------------
        # Efficienze fondamentali per giornata (per match)
        # -------------------------------------------------
        if kind in ("eff_battuta", "eff_ricezione", "eff_attacco", "eff_muro", "eff_difesa"):
            pts = []
            for _, r in dfm.iterrows():
                lab = _round_label(r["phase"], r["round_number"])
                ordv = match_order_value(str(r["phase"]), int(r["round_number"]))

                ta = canonical_team(fix_team_name(r["team_a"]))
                tb = canonical_team(fix_team_name(r["team_b"]))
                rallies = parse_rallies(r.get("scout_text", ""))

                if kind == "eff_battuta":
                    # Eff Battuta: (Punti + Half*0.8 + Pos*0.45 + Esc*0.3 + Neg*0.15 - Err) / Tot * 100
                    counts = {ta: {"Tot":0,"Punti":0,"Half":0,"Pos":0,"Esc":0,"Neg":0,"Err":0},
                              tb: {"Tot":0,"Punti":0,"Half":0,"Pos":0,"Esc":0,"Neg":0,"Err":0}}
                    def serve_type(c6: str) -> str:
                        return c6[3:5] if c6 and len(c6) >= 5 else ""
                    def serve_sign(c6: str) -> str:
                        return c6[5] if c6 and len(c6) >= 6 else ""
                    for rally in rallies:
                        if not rally or not is_serve(rally[0]):
                            continue
                        first = rally[0]
                        stype = serve_type(first)
                        if stype not in ("SQ","SM"):  # includi tutto
                            continue
                        team = ta if first[0] == "*" else tb
                        rec = counts[team]
                        rec["Tot"] += 1
                        s = serve_sign(first)
                        if s == "#":
                            rec["Punti"] += 1
                        elif s == "/":
                            rec["Half"] += 1
                        elif s == "+":
                            rec["Pos"] += 1
                        elif s == "!":
                            rec["Esc"] += 1
                        elif s == "-":
                            rec["Neg"] += 1
                        elif s == "=":
                            rec["Err"] += 1
                    for team, c in counts.items():
                        tot = c["Tot"]
                        eff = ((c["Punti"] + c["Half"]*0.8 + c["Pos"]*0.45 + c["Esc"]*0.3 + c["Neg"]*0.15 - c["Err"]) / tot * 100.0) if tot else 0.0
                        pts.append({"Team": team, "Round": lab, "order": ordv, "value": eff})

                elif kind == "eff_ricezione":
                    # Eff Ricezione: (Ok*0.77 + Escl*0.55 + Neg*0.38 - Mez*0.8 - Err) / Tot * 100
                    counts = {ta: {"Tot":0,"Perf":0,"Pos":0,"Escl":0,"Neg":0,"Mez":0,"Err":0},
                              tb: {"Tot":0,"Perf":0,"Pos":0,"Escl":0,"Neg":0,"Mez":0,"Err":0}}
                    def serve_type(c6: str) -> str:
                        return c6[3:5] if c6 and len(c6) >= 5 else ""
                    def first_reception(rally: list[str], prefix: str):
                        for c in rally:
                            if len(c) >= 6 and c[0] == prefix and c[3:5] in ("RQ","RM"):
                                return c
                        return None
                    for rally in rallies:
                        if not rally or not is_serve(rally[0]):
                            continue
                        first = rally[0]
                        stype = serve_type(first)
                        if stype not in ("SQ","SM"):
                            continue
                        if first[0] == "*":  # home served -> away receives
                            recv_team = tb
                            recv_prefix = "a"
                        else:
                            recv_team = ta
                            recv_prefix = "*"
                        rece = first_reception(rally, recv_prefix)
                        if not rece:
                            continue
                        s = rece[5]
                        rec = counts[recv_team]
                        rec["Tot"] += 1
                        if s == "#":
                            rec["Perf"] += 1
                        elif s == "+":
                            rec["Pos"] += 1
                        elif s == "!":
                            rec["Escl"] += 1
                        elif s == "-":
                            rec["Neg"] += 1
                        elif s == "/":
                            rec["Mez"] += 1
                        elif s == "=":
                            rec["Err"] += 1
                    for team, c in counts.items():
                        tot = c["Tot"]
                        ok_cnt = c["Perf"] + c["Pos"]
                        eff = ((ok_cnt*0.77 + c["Escl"]*0.55 + c["Neg"]*0.38 - c["Mez"]*0.8 - c["Err"]) / tot * 100.0) if tot else 0.0
                        pts.append({"Team": team, "Round": lab, "order": ordv, "value": eff})

                elif kind == "eff_attacco":
                    # Eff Attacco: (Punti - Murate - Errori) / Tot * 100 (includi primo attacco + transizione)
                    counts = {ta: {"Tot":0,"Punti":0,"Mur":0,"Err":0},
                              tb: {"Tot":0,"Punti":0,"Mur":0,"Err":0}}
                    def is_attack_code(c: str) -> bool:
                        return len(c) >= 6 and c[3] == "A" and c[0] in ("*", "a")
                    for rally in rallies:
                        for c in rally:
                            if not is_attack_code(c):
                                continue
                            team = ta if c[0] == "*" else tb
                            rec = counts[team]
                            rec["Tot"] += 1
                            s = c[5]
                            if s == "#":
                                rec["Punti"] += 1
                            elif s == "/":
                                rec["Mur"] += 1
                            elif s == "=":
                                rec["Err"] += 1
                    for team, c in counts.items():
                        tot = c["Tot"]
                        eff = ((c["Punti"] - c["Mur"] - c["Err"]) / tot * 100.0) if tot else 0.0
                        pts.append({"Team": team, "Round": lab, "order": ordv, "value": eff})

                elif kind == "eff_muro":
                    # Eff Muro: (Vincenti*2 + Pos*0.7 + Neg*0.07 + Cop*0.15 - Inv - Err) / Tot * 100
                    counts = {ta: {"Tot":0,"Perf":0,"Pos":0,"Neg":0,"Cop":0,"Inv":0,"Err":0},
                              tb: {"Tot":0,"Perf":0,"Pos":0,"Neg":0,"Cop":0,"Inv":0,"Err":0}}
                    def is_block(c: str) -> bool:
                        return len(c) >= 6 and c[3] == "B" and c[0] in ("*", "a")
                    for rally in rallies:
                        for c in rally:
                            if not is_block(c):
                                continue
                            team = ta if c[0] == "*" else tb
                            rec = counts[team]
                            rec["Tot"] += 1
                            s = c[5]
                            if s == "#":
                                rec["Perf"] += 1
                            elif s == "+":
                                rec["Pos"] += 1
                            elif s == "-":
                                rec["Neg"] += 1
                            elif s == "!":
                                rec["Cop"] += 1
                            elif s == "/":
                                rec["Inv"] += 1
                            elif s == "=":
                                rec["Err"] += 1
                    for team, c in counts.items():
                        tot = c["Tot"]
                        eff = ((c["Perf"]*2.0 + c["Pos"]*0.7 + c["Neg"]*0.07 + c["Cop"]*0.15 - c["Inv"] - c["Err"]) / tot * 100.0) if tot else 0.0
                        pts.append({"Team": team, "Round": lab, "order": ordv, "value": eff})

                elif kind == "eff_difesa":
                    # Eff Difesa: ((Perf+Cop) - (Neg+Over+Err)) / Tot * 100
                    counts = {ta: {"Tot":0,"Perf":0,"Cop":0,"Neg":0,"Over":0,"Err":0},
                              tb: {"Tot":0,"Perf":0,"Cop":0,"Neg":0,"Over":0,"Err":0}}
                    def is_def(c: str) -> bool:
                        return len(c) >= 6 and c[3] == "D" and c[0] in ("*", "a")
                    for rally in rallies:
                        for c in rally:
                            if not is_def(c):
                                continue
                            team = ta if c[0] == "*" else tb
                            rec = counts[team]
                            rec["Tot"] += 1
                            s = c[5]
                            if s == "+":
                                rec["Perf"] += 1
                            elif s == "!":
                                rec["Cop"] += 1
                            elif s == "-":
                                rec["Neg"] += 1
                            elif s == "/":
                                rec["Over"] += 1
                            elif s == "=":
                                rec["Err"] += 1
                    for team, c in counts.items():
                        tot = c["Tot"]
                        eff = (((c["Perf"] + c["Cop"]) - (c["Neg"] + c["Over"] + c["Err"])) / tot * 100.0) if tot else 0.0
                        pts.append({"Team": team, "Round": lab, "order": ordv, "value": eff})

            dfp = pd.DataFrame(pts)
            if dfp.empty:
                return dfp
            # per Team+Round: media ponderata su Tot? qui value già per match. Usiamo media semplice se ci sono più match nello stesso Round (caso raro)
            dfp = dfp.groupby(["Team", "Round", "order"], as_index=False).agg({"value": "mean"})
            return dfp

        # -------------------------------------------------
        # Indici num/den (SideOut/Break/BT) – default
        # -------------------------------------------------
        pts = []
        for _, r in dfm.iterrows():
            lab = _round_label(r["phase"], r["round_number"])
            ordv = match_order_value(str(r["phase"]), int(r["round_number"]))
            ta = r["team_a_city"]
            tb = r["team_b_city"]
            if ta:
                pts.append({"Team": ta, "Round": lab, "order": ordv, "num": float(r[mdef["home_num"]]), "den": float(r[mdef["home_den"]])})
            if tb:
                pts.append({"Team": tb, "Round": lab, "order": ordv, "num": float(r[mdef["away_num"]]), "den": float(r[mdef["away_den"]])})
        dfp = pd.DataFrame(pts)
        if dfp.empty:
            return dfp
        dfp = dfp.groupby(["Team", "Round", "order"], as_index=False).agg({"num": "sum", "den": "sum"})
        dfp["value"] = dfp.apply(lambda rr: _pct(rr["num"], rr["den"]), axis=1)
        return dfp

    dfx = build_points(x_metric).rename(columns={"value": "X"})
    dfy = build_points(y_metric).rename(columns={"value": "Y"})

    required = {"Team","Round","order","X"}
    if not required.issubset(set(dfx.columns)) or not {"Team","Round","order","Y"}.issubset(set(dfy.columns)):
        st.warning("Dati insufficienti per costruire il grafico con le metriche selezionate (per ora supportati: Side Out / Break / BT).")
        return

    df = dfx.merge(dfy[["Team", "Round", "order", "Y"]], on=["Team", "Round", "order"], how="inner")
    df = df[df["Team"].isin(selected)].copy()
    if df.empty:
        st.info("Nessun punto per le squadre selezionate nel range.")
        return
    df = df.sort_values(["Team", "order"]).reset_index(drop=True)

    import matplotlib.pyplot as plt
    fig, ax = plt.subplots()

    palette = ["tab:red", "tab:blue", "tab:green", "tab:purple", "tab:orange", "tab:brown"]

    for i, team in enumerate(selected):
        sub = df[df["Team"] == team]
        if sub.empty:
            continue
        color = palette[i % len(palette)]
        ax.scatter(sub["X"], sub["Y"], s=45, label=team, color=color)
        for _, rr in sub.iterrows():
            ax.text(rr["X"], rr["Y"], rr["Round"], fontsize=8, color=color)
    ax.set_xlabel(x_metric)
    ax.set_ylabel(y_metric)
    ax.set_title("4 Quadranti – punti per giornata (per squadra)")
    ax.grid(True, linewidth=0.3)
    ax.legend(loc="best")

    st.pyplot(fig, clear_figure=True)

    with st.expander("Tabella punti (debug / export)", expanded=False):
        out = df[["Team", "Round", "X", "Y"]].copy()
        render_any_table(out, fmt={"X": "{:.1f}", "Y": "{:.1f}"}, highlight_cols=["X", "Y"])



# =========================
# UI: IMPORT RUOLI (ROSTER)
# =========================
def render_import_ruoli(admin_mode: bool):
    st.header("Import Ruoli (Roster)")

    if not admin_mode:
        st.warning("Accesso riservato allo staff (admin).")
        return

    # --- Import XLSX ---
    up = st.file_uploader("Carica file Ruoli (.xlsx)", type=["xlsx"])
    st.info("Attese colonne: Team | Nome | Ruolo | N° (ID facoltativo).")

    season = st.text_input("Stagione (obbligatoria, es. 2025-26)", value="2025-26")

    if up is not None:
        df = pd.read_excel(up)
        df.columns = [str(c).strip() for c in df.columns]

        required = ["Team", "Nome", "Ruolo", "N°"]
        missing = [c for c in required if c not in df.columns]
        if missing:
            st.error(f"Mancano colonne: {missing}")
            st.stop()

        df = df.copy()
        df["Team"] = df["Team"].astype(str).apply(fix_team_name)
        df["team_norm"] = df["Team"].astype(str).apply(team_norm)
        df["Nome"] = df["Nome"].astype(str).str.strip()
        df["Ruolo"] = df["Ruolo"].astype(str).str.strip()
        df["N°"] = pd.to_numeric(df["N°"], errors="coerce").astype("Int64")

        df = df.dropna(subset=["N°", "team_norm"])
        df["N°"] = df["N°"].astype(int)

        st.subheader("Preview import")
        st.dataframe(df[["Team", "Nome", "Ruolo", "N°"]].head(80), width="stretch", hide_index=True)

        if st.button("Importa/aggiorna roster nel DB", key="btn_roster_import"):
            sql = """
            INSERT INTO roster (season, team_raw, team_norm, jersey_number, player_name, role, created_at)
            VALUES (:season, :team_raw, :team_norm, :jersey_number, :player_name, :role, :created_at)
            ON CONFLICT(season, team_norm, jersey_number) DO UPDATE SET
                team_raw = excluded.team_raw,
                player_name = excluded.player_name,
                role = excluded.role,
                created_at = excluded.created_at
            """
            now = datetime.now(timezone.utc).isoformat()
            with engine.begin() as conn:
                for _, r in df.iterrows():
                    conn.execute(
                        text(sql),
                        {
                            "season": season,
                            "team_raw": str(r["Team"]),
                            "team_norm": str(r["team_norm"]),
                            "jersey_number": int(r["N°"]),
                            "player_name": str(r["Nome"]),
                            "role": str(r["Ruolo"]),
                            "created_at": now,
                        }
                    )
            st.success("Roster importato/aggiornato.")
            st.rerun()

    st.divider()
    st.subheader("Correggi / Elimina record")

    # --- Carica roster dal DB per la stagione ---
    with engine.begin() as conn:
        rows = conn.execute(
            text("""
                SELECT id, season, team_raw, team_norm, jersey_number, player_name, role
                FROM roster
                WHERE season = :season
                ORDER BY team_raw, jersey_number
            """),
            {"season": season}
        ).mappings().all()

    if not rows:
        st.info("Nessun record in roster per questa stagione. Importa un file .xlsx sopra.")
        return

    df_db = pd.DataFrame(rows)

    # ===== Selezione: Team -> Giocatore (con filtro nome) =====
    teams = sorted(df_db["team_raw"].dropna().unique().tolist())
    team_sel = st.selectbox("Team", teams, index=0, key="edit_team")

    filtro_nome = st.text_input("Filtro nome (opzionale)", value="", key="edit_name_filter").strip().lower()

    df_team = df_db[df_db["team_raw"] == team_sel].copy()
    if filtro_nome:
        df_team = df_team[df_team["player_name"].fillna("").str.lower().str.contains(filtro_nome)]

    if df_team.empty:
        st.warning("Nessun giocatore trovato con questi filtri.")
        return

    df_team = df_team.sort_values(by=["jersey_number", "player_name"], ascending=[True, True])

    # opzioni: usiamo id come chiave stabile
    options = df_team["id"].tolist()

    def fmt_player(rid: int) -> str:
        r = df_team[df_team["id"] == rid].iloc[0]
        num = int(r["jersey_number"]) if pd.notna(r["jersey_number"]) else 0
        name = str(r["player_name"] or "").strip()
        role = str(r["role"] or "").strip()
        return f"{num:02d} — {name}  ({role})"

    player_id = st.selectbox("Giocatore", options, format_func=fmt_player, key="edit_player_id")

    rec = df_team[df_team["id"] == player_id].iloc[0].to_dict()
    st.caption(f"Record selezionato: id={rec['id']} | season={rec['season']} | team_norm={rec['team_norm']}")

    # campi editabili
    c3, c4 = st.columns(2)
    with c3:
        new_name = st.text_input("Nome giocatore", value=str(rec.get("player_name") or ""), key="edit_name")
        new_team_raw = st.text_input("Team (testo)", value=str(rec.get("team_raw") or ""), key="edit_team_raw")
    with c4:
        role_options = ["Alzatore", "Opposto", "Centrale", "Schiacciatore", "Libero"]
        current_role = str(rec.get("role") or "").strip()
        if current_role and current_role not in role_options:
            role_options = [current_role] + role_options
        new_role = st.selectbox("Ruolo", role_options, index=role_options.index(current_role) if current_role in role_options else 0, key="edit_role")
        new_num = st.number_input("Numero maglia", min_value=0, max_value=99, value=int(rec.get("jersey_number") or 0), step=1, key="edit_jersey")

    c5, c6 = st.columns(2)
    with c5:
        if st.button("Salva correzione", key="btn_roster_save"):
            team_raw_fixed = fix_team_name(new_team_raw)
            team_norm_fixed = team_norm(team_raw_fixed)
            now = datetime.now(timezone.utc).isoformat()

            with engine.begin() as conn:
                conn.execute(
                    text("""
                        UPDATE roster
                        SET team_raw = :team_raw,
                            team_norm = :team_norm,
                            jersey_number = :jersey_number,
                            player_name = :player_name,
                            role = :role,
                            created_at = :created_at
                        WHERE id = :id
                    """),
                    {
                        "team_raw": team_raw_fixed,
                        "team_norm": team_norm_fixed,
                        "jersey_number": int(new_num),
                        "player_name": str(new_name).strip(),
                        "role": str(new_role).strip(),
                        "created_at": now,
                        "id": int(rec["id"]),
                    }
                )
            st.success("Record aggiornato.")
            st.rerun()

    with c6:
        confirm_del = st.checkbox("Confermo: elimina questo record", value=False, key="confirm_del_roster")
        if st.button("Elimina record", disabled=not confirm_del, key="btn_roster_delete"):
            with engine.begin() as conn:
                conn.execute(text("DELETE FROM roster WHERE id = :id"), {"id": int(rec["id"])})
            st.success("Record eliminato.")
            st.rerun()


def render_sideout_players_by_role():
    st.header("Indici Side Out - Giocatori (per ruolo)")

    voce = st.radio(
        "Seleziona indice",
        [
            "SIDE OUT TOTALE",
            "SIDE OUT SPIN",
            "SIDE OUT FLOAT",
        ],
        index=0,
        key="so_pl_voce",
    )

    # ===== RANGE GIORNATE =====
    with engine.begin() as conn:
        bounds = conn.execute(text("""
            SELECT MIN(round_number) AS min_r, MAX(round_number) AS max_r
            FROM matches
            WHERE round_number IS NOT NULL
        """)).mappings().first()

    min_r = int((bounds["min_r"] or 1))
    max_r = int((bounds["max_r"] or 1))

    from_round, to_round, _range_label = get_selected_match_range()
    st.caption(f"📌 Range: {_range_label}")

    if from_round > to_round:
        st.error("Range non valido: 'Da giornata' deve essere <= 'A giornata'.")
        st.stop()

    # ===== FILTRO RUOLI + MIN RICEZIONI =====
    season = st.text_input("Stagione roster", value="2025-26", key="so_pl_season")

    with engine.begin() as conn:
        roster_rows = conn.execute(text("""
            SELECT team_raw, team_norm, jersey_number, player_name, role, created_at
            FROM roster
            WHERE season = :season
        """), {"season": season}).mappings().all()

    if not roster_rows:
        st.warning("Roster vuoto per questa stagione: importa prima i ruoli (pagina Import Ruoli).")
        return

    df_roster = pd.DataFrame(roster_rows)
    df_roster["created_at"] = df_roster["created_at"].fillna("")
    df_roster = (
        df_roster.sort_values(by=["team_norm", "jersey_number", "created_at"])
                 .drop_duplicates(subset=["team_norm", "jersey_number"], keep="last")
    )

    roles_all = sorted(df_roster["role"].dropna().unique().tolist())
    roles_sel = st.multiselect(
        "Filtra per ruolo (selezione multipla)",
        options=roles_all,
        default=roles_all,
        key="so_pl_roles",
    )

    min_recv = st.number_input(
        "Numero minimo di ricezioni (sotto questo valore il giocatore non viene mostrato)",
        min_value=0,
        value=10,
        step=1,
        key="so_pl_minrecv",
    )

    st.info(
        "Si precisa che il ranking rappresenta la percentuale di side out ottenuta durante la ricezione da parte del giocatore indicato. "
        "Selezionando il titolo di ciascuna colonna, i nominativi saranno ordinati in base al parametro corrispondente."
    )

    # ===== MATCHES =====
    with engine.begin() as conn:
        matches = conn.execute(text("""
            SELECT team_a, team_b, scout_text
            FROM matches
            WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
        """), {"from_round": int(from_round), "to_round": int(to_round)}).mappings().all()

    if not matches:
        st.info("Nessun match nel range selezionato.")
        return

    SET_RE = re.compile(r"\*\*(\d)set\b", re.IGNORECASE)

    def parse_rallies(scout_text: str):
        if not scout_text:
            return []
        lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
        current_set = None
        rallies = []
        current = []
        for ln in lines:
            mm = SET_RE.search(ln)
            if mm:
                current_set = int(mm.group(1))
                continue
            if current_set is None:
                continue
            c6 = code6(ln)
            if not c6 or c6[0] not in ("*", "a"):
                continue
            if is_serve(c6):
                if current:
                    rallies.append(current)
                current = [c6]
                continue
            if not current:
                continue
            current.append(c6)
            if is_home_point(c6) or is_away_point(c6):
                rallies.append(current)
                current = []
        if current:
            rallies.append(current)
        return rallies

    def player_num(c6: str):
        mm = re.match(r"^[\*a](\d{2})", c6)
        return int(mm.group(1)) if mm else None

    def is_rece(c6: str) -> bool:
        return len(c6) >= 6 and c6[0] in ("*", "a") and c6[3:5] in ("RQ", "RM")

    def is_spin_serve(c6: str) -> bool:
        return len(c6) >= 5 and c6[3:5] == "SQ"

    def is_float_serve(c6: str) -> bool:
        return len(c6) >= 5 and c6[3:5] == "SM"

    def is_attack(c6: str) -> bool:
        return len(c6) >= 6 and c6[0] in ("*", "a") and c6[3] == "A"

    def rece_player_and_sign(rally, rece_prefix):
        for x in rally:
            if is_rece(x) and x[0] == rece_prefix:
                return player_num(x), x[5]
        return None, None

    def first_attack_winner_after_rece(rally, rece_prefix):
        seen_rece = False
        for x in rally:
            if not seen_rece:
                if is_rece(x) and x[0] == rece_prefix:
                    seen_rece = True
                continue
            if is_attack(x) and x[0] == rece_prefix:
                return (len(x) >= 6 and x[5] == "#")
        return False

    P = {}
    team_recv_tot = {}

    def ensure_player_rec(team_raw: str, num: int):
        tn = team_norm(team_raw)
        key = (tn, int(num))
        if key not in P:
            P[key] = {"team_norm": tn, "Squadra": canonical_team(team_raw), "N°": int(num), "recv": 0, "so": 0, "sod": 0}
        return P[key]

    def add_team_recv(team_raw: str, n: int):
        tn = team_norm(team_raw)
        team_recv_tot[tn] = team_recv_tot.get(tn, 0) + int(n)

    for m in matches:
        ta_raw = (m.get("team_a") or "")
        tb_raw = (m.get("team_b") or "")
        ta = ta_raw
        tb = tb_raw
        rallies = parse_rallies(m.get("scout_text") or "")

        for r in rallies:
            if not r or not is_serve(r[0]):
                continue
            serve = r[0]

            if voce == "SIDE OUT SPIN" and not is_spin_serve(serve):
                continue
            if voce == "SIDE OUT FLOAT" and not is_float_serve(serve):
                continue

            s_prefix = serve[0]
            rece_prefix = "a" if s_prefix == "*" else "*"
            rece_team = ta if rece_prefix == "*" else tb

            pnum, _ = rece_player_and_sign(r, rece_prefix)
            if pnum is None:
                continue

            home_won = any(is_home_point(x) for x in r)
            away_won = any(is_away_point(x) for x in r)
            rece_team_won = (home_won and rece_prefix == "*") or (away_won and rece_prefix == "a")

            rec = ensure_player_rec(rece_team, pnum)
            rec["recv"] += 1
            add_team_recv(rece_team, 1)

            if rece_team_won:
                rec["so"] += 1
                if first_attack_winner_after_rece(r, rece_prefix):
                    rec["sod"] += 1

    df = pd.DataFrame(list(P.values()))
    if df.empty:
        st.info("Nessun dato trovato nel range.")
        return

    df = df.merge(
        df_roster[["team_norm", "jersey_number", "player_name", "role", "team_raw"]],
        left_on=["team_norm", "N°"],
        right_on=["team_norm", "jersey_number"],
        how="left",
    ).drop(columns=["jersey_number"])

    df.rename(columns={"player_name": "Nome giocatore", "role": "Ruolo"}, inplace=True)
    df["Nome giocatore"] = df["Nome giocatore"].fillna(df["N°"].apply(lambda x: f"N°{int(x):02d}"))
    df["Ruolo"] = df["Ruolo"].fillna("(non in roster)")
    df["Squadra"] = df["team_raw"].apply(canonical_team).fillna(df["Squadra"])
    df = df.drop(columns=["team_raw"])

    # Canonicalizza sempre la colonna Squadra (solo città)
    df["Squadra"] = df["Squadra"].apply(canonical_team)


    if roles_sel:
        df = df[df["Ruolo"].isin(roles_sel)].copy()

    df = df[df["recv"] >= int(min_recv)].copy()
    if df.empty:
        st.info("Nessun giocatore sopra il minimo ricezioni nel filtro selezionato.")
        return

    df["% di SO"] = df.apply(lambda r: (100.0 * r["so"] / r["recv"]) if r["recv"] else 0.0, axis=1)
    df["% di SO-d"] = df.apply(lambda r: (100.0 * r["sod"] / r["recv"]) if r["recv"] else 0.0, axis=1)
    df["% Ply/Team"] = df.apply(lambda r: (100.0 * r["recv"] / team_recv_tot.get(r["team_norm"], 0)) if team_recv_tot.get(r["team_norm"], 0) else 0.0, axis=1)

    df_rank = df.sort_values(by=["% di SO", "recv"], ascending=[False, False]).reset_index(drop=True)
    df_rank.insert(0, "Ranking", range(1, len(df_rank) + 1))
    # Canonicalizza sempre la colonna Squadra (solo città)
    df_rank["Squadra"] = df_rank["Squadra"].apply(canonical_team)


    out = df_rank[["Ranking", "Nome giocatore", "Squadra", "recv", "% Ply/Team", "% di SO", "% di SO-d"]].rename(columns={
        "recv": "N° ricezioni fatte",
    }).copy()

    def highlight_perugia(row):
        is_perugia = "perugia" in str(row.get("Squadra", "")).lower()
        style = "background-color: #fff3cd; font-weight: 800;" if is_perugia else ""
        return [style] * len(row)

    def _fmt1(x):
        try:
            v = float(x)
            s = f"{v:.1f}"
            return s.rstrip("0").rstrip(".")
        except Exception:
            return x

    def _highlight_col(df_):
        styles = pd.DataFrame("", index=df_.index, columns=df_.columns)
        col = "% di SO"
        if col in df_.columns:
            styles[col] = "background-color: #e7f5ff; font-weight: 900;"
        return styles

    styled = (
        out.style
          .apply(highlight_perugia, axis=1)
          .apply(_highlight_col, axis=None)
          .format({
              "% di SO": _fmt1,
              "% di SO-d": _fmt1,
              "% Ply/Team": _fmt1,
              "N° ricezioni fatte": "{:.0f}",
              "Ranking": "{:.0f}",
          })
          .set_table_styles([
              {"selector": "th", "props": [("font-size", "22px"), ("text-align", "left"), ("padding", "8px 10px")]},
              {"selector": "td", "props": [("font-size", "21px"), ("padding", "8px 10px")]},
          ])
    )
    render_focus_4_players(out, key_base=f"so_{voce}_{from_round}_{to_round}")


    render_any_table(styled)


# =========================
# UI: BREAK POINT - GIOCATORI (per ruolo)
# =========================
def render_break_players_by_role():
    st.header("Indici Break Point - Giocatori (per ruolo)")

    # ===== RANGE GIORNATE =====
    with engine.begin() as conn:
        bounds = conn.execute(text("""
            SELECT MIN(round_number) AS min_r, MAX(round_number) AS max_r
            FROM matches
            WHERE round_number IS NOT NULL
        """)).mappings().first()

    min_r = int((bounds["min_r"] or 1))
    max_r = int((bounds["max_r"] or 1))

    from_round, to_round, _range_label = get_selected_match_range()
    st.caption(f"📌 Range: {_range_label}")

    if from_round > to_round:
        st.error("Range non valido.")
        st.stop()

    # ===== ROSTER (ruoli) =====
    season = st.text_input("Stagione roster (deve coincidere con quella importata)", value="2025-26", key="bppl_season")

    with engine.begin() as conn:
        roster_rows = conn.execute(text("""
            SELECT team_raw, team_norm, jersey_number, player_name, role, created_at
            FROM roster
            WHERE season = :season
        """), {"season": season}).mappings().all()

    if not roster_rows:
        st.warning("Roster vuoto per questa stagione: importa prima i ruoli (pagina Import Ruoli).")
        return

    df_roster = pd.DataFrame(roster_rows)
    df_roster["created_at"] = df_roster["created_at"].fillna("")
    df_roster = (
        df_roster.sort_values(by=["team_norm", "jersey_number", "created_at"])
                 .drop_duplicates(subset=["team_norm", "jersey_number"], keep="last")
    )

    roles_all = sorted(df_roster["role"].dropna().unique().tolist())
    roles_sel = st.multiselect("Filtro ruoli (puoi selezionarne quanti vuoi)", options=roles_all, default=roles_all, key="bppl_roles")

    min_serves = st.number_input(
        "Numero minimo di battute (sotto questo numero il giocatore non appare)",
        min_value=0, max_value=500, value=10, step=1, key="bppl_min_serves"
    )

    st.info(
        "Si precisa che il ranking rappresenta la percentuale di Break Point ottenuta durante la battuta da parte del giocatore indicato. "
        "Selezionando il titolo di ciascuna colonna, i nominativi saranno ordinati in base al parametro corrispondente."
    )

    # ===== MATCHES NEL RANGE =====
    with engine.begin() as conn:
        matches = conn.execute(text("""
            SELECT team_a, team_b, scout_text
            FROM matches
            WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
        """), {"from_round": int(from_round), "to_round": int(to_round)}).mappings().all()

    if not matches:
        st.info("Nessun match nel range.")
        return

    def parse_rallies(scout_text: str):
        if not scout_text:
            return []
        lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
        scout_lines = [ln for ln in lines if ln[0] in ("*", "a")]

        rallies = []
        current = []
        for raw in scout_lines:
            c = code6(raw)
            if not c:
                continue
            if is_serve(c):
                if current:
                    rallies.append(current)
                current = [c]
                continue
            if not current:
                continue
            current.append(c)
            if is_home_point(c) or is_away_point(c):
                rallies.append(current)
                current = []
        return rallies

    agg = {}
    team_agg = {}

    def ensure_player(team_raw: str, num: int):
        tnorm = team_norm(team_raw)
        key = (tnorm, num)
        if key not in agg:
            agg[key] = {
                "team_norm": tnorm,
                "Squadra": canonical_team(team_raw),
                "N°": num,
                "serves": 0,
                "bp_win": 0,
                "aces": 0,
                "errors": 0,
                "played_serves": 0,
                "played_bp_win": 0,
            }
        return agg[key]

    def ensure_team(team_raw: str):
        tnorm = team_norm(team_raw)
        if tnorm not in team_agg:
            team_agg[tnorm] = {"team_norm": tnorm, "Squadra": canonical_team(team_raw), "serves": 0, "bp_win": 0}
        return team_agg[tnorm]

    for m in matches:
        team_a = (m.get("team_a") or "")
        team_b = (m.get("team_b") or "")

        rallies = parse_rallies(m.get("scout_text") or "")
        for rally in rallies:
            if not rally:
                continue
            first = rally[0]
            if not is_serve(first):
                continue

            home_served = first.startswith("*")
            away_served = first.startswith("a")

            home_point = any(is_home_point(x) for x in rally)
            away_point = any(is_away_point(x) for x in rally)

            if home_served:
                serving_team = team_a
                bp = 1 if home_point else 0
            elif away_served:
                serving_team = team_b
                bp = 1 if away_point else 0
            else:
                continue

            num = serve_player_number(first)
            if num is None:
                continue

            rec = ensure_player(serving_team, num)
            rec["serves"] += 1
            rec["bp_win"] += bp

            sgn = serve_sign(first)
            if sgn == "#":
                rec["aces"] += 1
            elif sgn == "=":
                rec["errors"] += 1

            if sgn not in ("#", "="):
                rec["played_serves"] += 1
                rec["played_bp_win"] += bp

            t = ensure_team(serving_team)
            t["serves"] += 1
            t["bp_win"] += bp

    df = pd.DataFrame(list(agg.values()))
    if df.empty:
        st.info("Nessun dato (battute) nel range selezionato.")
        return

    df = df.merge(
        df_roster[["team_norm", "jersey_number", "player_name", "role"]],
        left_on=["team_norm", "N°"],
        right_on=["team_norm", "jersey_number"],
        how="left",
    ).drop(columns=["jersey_number"])

    df.rename(columns={"player_name": "Nome giocatore", "role": "Ruolo"}, inplace=True)
    df["Nome giocatore"] = df["Nome giocatore"].fillna("(non in roster)")
    df["Ruolo"] = df["Ruolo"].fillna("(non in roster)")

    if roles_sel:
        df = df[df["Ruolo"].isin(roles_sel)].copy()

    df = df[df["serves"] >= int(min_serves)].copy()
    if df.empty:
        st.info("Nessun giocatore supera il filtro del numero minimo di battute.")
        return

    df["% di BP"] = (100.0 * df["bp_win"] / df["serves"].replace({0: pd.NA})).fillna(0.0)
    df["% di BP giocato"] = (100.0 * df["played_bp_win"] / df["played_serves"].replace({0: pd.NA})).fillna(0.0)

    df_team = pd.DataFrame(list(team_agg.values()))
    df_team["% Team BP"] = (100.0 * df_team["bp_win"] / df_team["serves"].replace({0: pd.NA})).fillna(0.0)
    df = df.merge(df_team[["team_norm", "% Team BP"]], on="team_norm", how="left")
    df["Diff. Team"] = df["% di BP"].fillna(0.0) - df["% Team BP"].fillna(0.0)

    def safe_ratio(err, ace):
        try:
            err = float(err or 0)
            ace = float(ace or 0)
        except Exception:
            return 0.0
        if ace == 0:
            return float("inf") if err > 0 else 0.0
        return err / ace

    df["Ratio err/p.to"] = df.apply(lambda r: safe_ratio(r["errors"], r["aces"]), axis=1)

    df_rank = df.sort_values(by=["% di BP", "serves"], ascending=[False, False]).reset_index(drop=True)
    df_rank.insert(0, "Ranking", range(1, len(df_rank) + 1))

    out = df_rank[[
        "Ranking",
        "Nome giocatore",
        "Squadra",
        "serves",
        "% di BP",
        "Diff. Team",
        "Ratio err/p.to",
        "% di BP giocato",
    ]].rename(columns={"serves": "Battute", "Ratio err/p.to": "Err/P.ti"}).copy()

    # Ratio come stringa (1 decimale) per render uniforme in Streamlit
    def ratio_str(x):
        try:
            if x == float("inf"):
                return "∞"
            return f"{float(x):.1f}"
        except Exception:
            return ""

    out["Err/P.ti"] = out["Err/P.ti"].apply(ratio_str)


    def highlight_perugia(row):
        is_perugia = "perugia" in str(row["Squadra"]).lower()
        style = "background-color: #fff3cd; font-weight: 800;" if is_perugia else ""
        return [style] * len(row)

    styled = (
        out.style
          .apply(highlight_perugia, axis=1)
          .set_properties(subset=["% di BP"], **{"background-color": "#e7f5ff", "font-weight": "900"})
          .format({
              "Ranking": "{:.0f}",
              "N° di Battute Fatte": "{:.0f}",
              "% di BP": "{:.1f}",
              "Diff. Team": "{:.1f}",
              "% di BP giocato": "{:.1f}",
          })
    )
    styled = styled.set_table_styles([
        {"selector": "th", "props": [("font-size", "22px"), ("text-align", "left"), ("padding", "8px 10px")]},
        {"selector": "td", "props": [("font-size", "21px"), ("padding", "8px 10px")]},
    ])
    render_focus_4_players(out, key_base=f"bp_{from_round}_{to_round}")


    render_any_table(styled)


# =========================
# UI: CLASSIFICHE FONDAMENTALI - SQUADRE
# =========================
def render_fondamentali_team():
    st.header("Classifiche Fondamentali - Squadre")

    voce = st.radio(
        "Seleziona fondamentale",
        [
            "Battuta",
            "Ricezione",
            "Attacco",
            "Muro",
            "Difesa",
        ],
        index=0,
        key="fund_team_voce",
    )

    # ===== FILTRO RANGE GIORNATE =====
    with engine.begin() as conn:
        bounds = conn.execute(text("""
            SELECT MIN(round_number) AS min_r, MAX(round_number) AS max_r
            FROM matches
            WHERE round_number IS NOT NULL
        """)).mappings().first()

    min_r = int((bounds["min_r"] or 1))
    max_r = int((bounds["max_r"] or 1))

    from_round, to_round, _range_label = get_selected_match_range()
    st.caption(f"📌 Range: {_range_label}")

    if from_round > to_round:
        st.error("Range non valido: 'Da giornata' deve essere <= 'A giornata'.")
        st.stop()


    # =======================
    # DIFESA (SQUADRE) — come tabella giocatori (Efficienza difesa)
    # =======================
    if voce == "Difesa":
        st.subheader("Difesa")

        # stato iniziale (toggle identico alla tabella giocatori)
        if "fund_tm_def_total" not in st.session_state:
            st.session_state.fund_tm_def_total = True
        for k in ("dt", "dq", "dm", "dh"):
            kk = f"fund_tm_def_{k}"
            if kk not in st.session_state:
                st.session_state[kk] = False

        def _toggle_total_tm():
            if st.session_state.fund_tm_def_total:
                st.session_state.fund_tm_def_dt = False
                st.session_state.fund_tm_def_dq = False
                st.session_state.fund_tm_def_dm = False
                st.session_state.fund_tm_def_dh = False

        def _toggle_specific_tm():
            if (st.session_state.fund_tm_def_dt or st.session_state.fund_tm_def_dq or
                st.session_state.fund_tm_def_dm or st.session_state.fund_tm_def_dh):
                st.session_state.fund_tm_def_total = False
            if (not st.session_state.fund_tm_def_total and
                not st.session_state.fund_tm_def_dt and not st.session_state.fund_tm_def_dq and
                not st.session_state.fund_tm_def_dm and not st.session_state.fund_tm_def_dh):
                st.session_state.fund_tm_def_total = True

        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            st.checkbox("Difesa su palla Spinta", key="fund_tm_def_dt", on_change=_toggle_specific_tm)
        with c2:
            st.checkbox("Difesa su 1°tempo", key="fund_tm_def_dq", on_change=_toggle_specific_tm)
        with c3:
            st.checkbox("Difesa su PIPE", key="fund_tm_def_dm", on_change=_toggle_specific_tm)
        with c4:
            st.checkbox("Difesa su H-ball", key="fund_tm_def_dh", on_change=_toggle_specific_tm)
        with c5:
            st.checkbox("Difesa TOTALE", key="fund_tm_def_total", on_change=_toggle_total_tm)

        st.caption(
            "L’efficienza della difesa è calcolata in questo modo: "
            "(Buone*2 + Coperture*0,5 + Negative*0,4 + OverTheNet*0,3 – Errori) / Tot * 100."
        )

        # ===== MATCHES NEL RANGE =====
        with engine.begin() as conn:
            rows = conn.execute(
                text("""
                    SELECT team_a, team_b, scout_text
                    FROM matches
                    WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        if not rows:
            st.info("Nessun match nel range selezionato.")
            return

        def parse_rallies(scout_text: str):
            if not scout_text:
                return []
            _lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
            scout_lines = [ln for ln in _lines if ln[0] in ("*", "a")]
            rallies = []
            current = []
            for raw in scout_lines:
                c6 = code6(raw)
                if not c6:
                    continue
                if is_serve(c6):
                    if current:
                        rallies.append(current)
                    current = [c6]
                    continue
                if not current:
                    continue
                current.append(c6)
                if is_home_point(c6) or is_away_point(c6):
                    rallies.append(current)
                    current = []
            return rallies

        def is_defense(c6: str) -> bool:
            return len(c6) >= 6 and c6[3] == "D" and c6[0] in ("*", "a")

        def def_type(c6: str) -> str:
            return c6[3:5] if c6 and len(c6) >= 5 else ""

        def def_sign(c6: str) -> str:
            return c6[5] if c6 and len(c6) >= 6 else ""

        def defense_selected(c6: str) -> bool:
            if st.session_state.fund_tm_def_total:
                return True
            t = def_type(c6)
            s = def_sign(c6)
            if st.session_state.fund_tm_def_dt and t == "DT":
                return s != "!"
            if st.session_state.fund_tm_def_dq and t == "DQ":
                return s != "!"
            if st.session_state.fund_tm_def_dm and t == "DM":
                return s != "!"
            if st.session_state.fund_tm_def_dh and t == "DH":
                return s != "!"
            return False

        agg = {}  # team -> counts
        def ensure(team_raw: str):
            team = canonical_team(team_raw)
            if team not in agg:
                agg[team] = {"Squadra": team, "Tot": 0, "Buone": 0, "Cop": 0, "Neg": 0, "Over": 0, "Err": 0}
            return agg[team]

        for m in rows:
            ta_raw = (m.get("team_a") or "")
            tb_raw = (m.get("team_b") or "")
            rallies = parse_rallies(m.get("scout_text") or "")
            for rally in rallies:
                if not rally:
                    continue
                for c in rally[1:]:
                    if not is_defense(c):
                        continue
                    if not defense_selected(c):
                        continue
                    team_raw = ta_raw if c[0] == "*" else tb_raw
                    rec = ensure(team_raw)
                    rec["Tot"] += 1
                    s = def_sign(c)
                    if s == "+":
                        rec["Buone"] += 1
                    elif s == "!":
                        rec["Cop"] += 1
                    elif s == "-":
                        rec["Neg"] += 1
                    elif s == "/":
                        rec["Over"] += 1
                    elif s == "=":
                        rec["Err"] += 1

        df = pd.DataFrame(list(agg.values()))
        if df.empty:
            st.info("Nessuna difesa trovata nel range selezionato.")
            return

        def pct(num, den):
            return (100.0 * num / den) if den else 0.0

        df["Buone%"] = df.apply(lambda r: pct(r["Buone"], r["Tot"]), axis=1)
        df["Cop%"]   = df.apply(lambda r: pct(r["Cop"],   r["Tot"]), axis=1)
        df["Neg%"]   = df.apply(lambda r: pct(r["Neg"],   r["Tot"]), axis=1)
        df["Over%"]  = df.apply(lambda r: pct(r["Over"],  r["Tot"]), axis=1)
        df["Err%"]   = df.apply(lambda r: pct(r["Err"],   r["Tot"]), axis=1)

        df["EFF"] = df.apply(
            lambda r: ((r["Buone"] * 2.0 + r["Cop"] * 0.5 + r["Neg"] * 0.4 + r["Over"] * 0.3 - r["Err"]) / r["Tot"] * 100.0) if r["Tot"] else 0.0,
            axis=1
        )

        df = df.sort_values(by=["EFF", "Tot"], ascending=[False, False]).reset_index(drop=True)
        df = df.head(12).reset_index(drop=True)
        df.insert(0, "Ranking", range(1, len(df) + 1))

        out = df[["Ranking", "Squadra", "Tot", "EFF", "Buone%", "Cop%", "Neg%", "Over%", "Err%"]].rename(columns={
            "Squadra": "squadra",
            "EFF": "Eff",
            "Buone%": "Buone",
            "Cop%": "Cop",
            "Neg%": "Neg",
            "Over%": "Over",
            "Err%": "Err",
        }).copy()

        def highlight_perugia(row):
            is_perugia = "perugia" in str(row["squadra"]).lower()
            style = "background-color: #fff3cd; font-weight: 800;" if is_perugia else ""
            return [style] * len(row)

        styled = (
            out.style
              .apply(highlight_perugia, axis=1)
              .set_properties(subset=["Eff"], **{"background-color": "#e7f5ff", "font-weight": "900"})
              .set_properties(subset=["EFF"], **{"background-color": "#e7f5ff", "font-weight": "900"})
              .format({
                  "Ranking": "{:.0f}",
                  "Tot": "{:.0f}",
                  "Eff": "{:.1f}",
                  "Buone": "{:.1f}",
                  "Cop": "{:.1f}",
                  "Neg": "{:.1f}",
                  "Over": "{:.1f}",
                  "Err": "{:.1f}",
              })
              .set_table_styles([
                  {"selector": "th", "props": [("font-size", "22px"), ("text-align", "left"), ("padding", "8px 10px")]},
                  {"selector": "td", "props": [("font-size", "21px"), ("padding", "8px 10px")]},
              ])
        )
        render_any_table(styled)
        return


    st.info("Sezione in costruzione. Ho già impostato il menu dei fondamentali: ora decidiamo insieme le metriche e le tabelle per ciascuna voce.")

    # Placeholder strutturale per sviluppo step-by-step (senza patch sparse)
    if voce == "Battuta":
        st.subheader("Battuta")

        # ===== FILTRI TIPO BATTUTA (SQ / SM) =====
        cbt1, cbt2 = st.columns(2)
        with cbt1:
            use_spin = st.checkbox("Battuta SPIN", value=True, key="fund_srv_spin")
        with cbt2:
            use_float = st.checkbox("Battuta FLOAT", value=True, key="fund_srv_float")

        # Regola: se spunti entrambe -> tutte; se spunti una -> solo quella; se nessuna -> tutte (fallback)
        if not use_spin and not use_float:
            use_spin = True
            use_float = True

        allowed_types = set()
        if use_spin:
            allowed_types.add("SQ")
        if use_float:
            allowed_types.add("SM")

        st.caption(
            "L’efficienza della battuta è calcolata in questo modo: "
            "(Punti + ½ P.*0,8 + Pos.*0,45 + Esc*0,3 + Neg*0,15 – Err) / Tot. "
            "Dove i coefficienti sono le % medie di B.Point con quel tipo di ricezione del campionato."
        )

        # ===== MATCHES NEL RANGE =====
        with engine.begin() as conn:
            rows = conn.execute(
                text("""
                    SELECT team_a, team_b, scout_text
                    FROM matches
                    WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        if not rows:
            st.info("Nessun match nel range selezionato.")
            return

        def parse_rallies(scout_text: str):
            if not scout_text:
                return []
            lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
            scout_lines = [ln for ln in lines if ln[0] in ("*", "a")]

            rallies = []
            current = []
            for raw in scout_lines:
                c = code6(raw)
                if not c:
                    continue
                if is_serve(c):
                    if current:
                        rallies.append(current)
                    current = [c]
                    continue
                if not current:
                    continue
                current.append(c)
                if is_home_point(c) or is_away_point(c):
                    rallies.append(current)
                    current = []
            return rallies

        def fix_team(name: str) -> str:
            # riusa le stesse regole già usate altrove
            return canonical_team(name)

        agg = {}  # team -> counts

        def ensure(team: str):
            if team not in agg:
                agg[team] = {
                    "Team": team,
                    "Tot": 0,
                    "Punti": 0,
                    "Half": 0,
                    "Err": 0,
                    "Pos": 0,
                    "Esc": 0,
                    "Neg": 0,
                }
            return agg[team]

        def serve_type(c6: str) -> str:
            return c6[3:5] if c6 and len(c6) >= 5 else ""

        def serve_sign_local(c6: str) -> str:
            return c6[5] if c6 and len(c6) >= 6 else ""

        for r in rows:
            ta = fix_team(r.get("team_a") or "")
            tb = fix_team(r.get("team_b") or "")
            rallies = parse_rallies(r.get("scout_text") or "")
            for rally in rallies:
                if not rally:
                    continue
                first = rally[0]
                if not is_serve(first):
                    continue

                stype = serve_type(first)
                if stype not in allowed_types:
                    continue

                # serving team
                if first.startswith("*"):
                    team = ta
                elif first.startswith("a"):
                    team = tb
                else:
                    continue

                rec = ensure(team)
                rec["Tot"] += 1

                sgn = serve_sign_local(first)
                if sgn == "#":
                    rec["Punti"] += 1
                elif sgn == "/":
                    rec["Half"] += 1
                elif sgn == "=":
                    rec["Err"] += 1
                elif sgn == "+":
                    rec["Pos"] += 1
                elif sgn == "!":
                    rec["Esc"] += 1
                elif sgn == "-":
                    rec["Neg"] += 1

        df = pd.DataFrame(list(agg.values()))
        if df.empty or df["Tot"].sum() == 0:
            st.info("Nessuna battuta trovata per i filtri selezionati.")
            return

        # percentuali su Tot
        def pct(num, den):
            return (100.0 * num / den) if den else 0.0

        df["Punti%"] = df.apply(lambda r: pct(r["Punti"], r["Tot"]), axis=1)
        df["Half%"]  = df.apply(lambda r: pct(r["Half"],  r["Tot"]), axis=1)
        df["Err%"]   = df.apply(lambda r: pct(r["Err"],   r["Tot"]), axis=1)
        df["Pos%"]   = df.apply(lambda r: pct(r["Pos"],   r["Tot"]), axis=1)
        df["Esc%"]   = df.apply(lambda r: pct(r["Esc"],   r["Tot"]), axis=1)
        df["Neg%"]   = df.apply(lambda r: pct(r["Neg"],   r["Tot"]), axis=1)

        # Eff = (Punti + Half*0.8 + Pos*0.45 + Esc*0.3 + Neg*0.15 - Err)/Tot
        # Nota: qui Punti/Half/... sono conteggi (non %)
        df["EFF"] = df.apply(
            lambda r: (
                (
                    r["Punti"]
                    + r["Half"] * 0.8
                    + r["Pos"] * 0.45
                    + r["Esc"] * 0.3
                    + r["Neg"] * 0.15
                    - r["Err"]
                )
                / r["Tot"]
                * 100.0
            ) if r["Tot"] else 0.0,
            axis=1
        )

        # Ordina per eff decrescente
        df = df.sort_values(by=["EFF", "Tot"], ascending=[False, False]).reset_index(drop=True)
        df.insert(0, "Rank", range(1, len(df) + 1))

        out = df[[
            "Rank", "Team", "Tot", "EFF",
            "Punti%", "Half%", "Err%", "Pos%", "Esc%", "Neg%"
        ]].rename(columns={
            "Rank": "Rank",
            "Team": "Team",
            "Tot": "Tot",
            "EFF": "EFF",
            "Punti%": "Punti",
            "Half%": "½ P",
            "Err%": "Err.",
            "Pos%": "Pos",
            "Esc%": "Esc",
            "Neg%": "Neg",
        }).copy()

        # stile: Perugia + evidenzia EFF
        def highlight_perugia(row):
            is_perugia = "perugia" in str(row["Team"]).lower()
            style = "background-color: #fff3cd; font-weight: 800;" if is_perugia else ""
            return [style] * len(row)

        styled = (
            out.style
              .apply(highlight_perugia, axis=1)
              .set_properties(subset=["EFF"], **{"background-color": "#e7f5ff", "font-weight": "900"})
              .format({
                  "Rank": "{:.0f}",
                  "Tot": "{:.0f}",
                  "EFF": "{:.1f}",
                  "Punti": "{:.1f}",
                  "½ P": "{:.1f}",
                  "Err.": "{:.1f}",
                  "Pos": "{:.1f}",
                  "Esc": "{:.1f}",
                  "Neg": "{:.1f}",
              })
              .set_table_styles([
                  {"selector": "th", "props": [("font-size", "22px"), ("text-align", "left"), ("padding", "8px 10px")]},
                  {"selector": "td", "props": [("font-size", "21px"), ("padding", "8px 10px")]},
              ])
        )

        render_any_table(styled)


    elif voce == "Ricezione":
        st.subheader("Ricezione")

        # ===== FILTRI TIPO BATTUTA AVVERSARIA (SQ / SM) =====
        cbt1, cbt2 = st.columns(2)
        with cbt1:
            use_spin = st.checkbox("Ricezione SPIN", value=True, key="fund_rec_spin")
        with cbt2:
            use_float = st.checkbox("Ricezione FLOAT", value=True, key="fund_rec_float")

        # Regola: entrambe -> tutte; una sola -> solo quella; nessuna -> tutte (fallback)
        if not use_spin and not use_float:
            use_spin = True
            use_float = True

        allowed_types = set()
        if use_spin:
            allowed_types.add("SQ")
        if use_float:
            allowed_types.add("SM")

        st.caption(
            "L’efficienza della ricezione è calcolata in questo modo: "
            "(Ok*0,77 + Escl*0,55 + Neg*0,38 – Mez*0,8 - Err) / Tot * 100. "
            "Dove i coefficienti sono le % medie di Side Out con quel tipo di ricezione del campionato."
        )

        # ===== MATCHES NEL RANGE =====
        with engine.begin() as conn:
            rows = conn.execute(
                text("""
                    SELECT team_a, team_b, scout_text
                    FROM matches
                    WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        if not rows:
            st.info("Nessun match nel range selezionato.")
            return

        def parse_rallies(scout_text: str):
            if not scout_text:
                return []
            lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
            scout_lines = [ln for ln in lines if ln[0] in ("*", "a")]

            rallies = []
            current = []
            for raw in scout_lines:
                c = code6(raw)
                if not c:
                    continue
                if is_serve(c):
                    if current:
                        rallies.append(current)
                    current = [c]
                    continue
                if not current:
                    continue
                current.append(c)
                if is_home_point(c) or is_away_point(c):
                    rallies.append(current)
                    current = []
            return rallies

        def serve_type(first: str) -> str:
            return first[3:5] if first and len(first) >= 5 else ""

        def first_reception(rally: list[str], prefix: str) -> str | None:
            # prima ricezione (RQ/RM) della squadra che riceve
            for c in rally:
                if len(c) >= 6 and c[0] == prefix and c[3:5] in ("RQ", "RM"):
                    return c
            return None

        agg = {}  # team -> counts

        def ensure(team: str):
            if team not in agg:
                agg[team] = {
                    "Team": team,
                    "Tot": 0,
                    "Perf": 0,   # '#'
                    "Pos": 0,    # '+'
                    "Escl": 0,   # '!'
                    "Neg": 0,    # '-'
                    "Mez": 0,    # '/'
                    "Err": 0,    # '='
                }
            return agg[team]

        for r in rows:
            ta = canonical_team(fix_team_name(r.get("team_a") or ""))
            tb = canonical_team(fix_team_name(r.get("team_b") or ""))

            rallies = parse_rallies(r.get("scout_text") or "")
            for rally in rallies:
                if not rally:
                    continue
                first = rally[0]
                if not is_serve(first):
                    continue

                stype = serve_type(first)
                if stype not in allowed_types:
                    continue

                home_served = first.startswith("*")
                away_served = first.startswith("a")

                if home_served:
                    recv_team = tb
                    recv_prefix = "a"
                elif away_served:
                    recv_team = ta
                    recv_prefix = "*"
                else:
                    continue

                rece = first_reception(rally, recv_prefix)
                if not rece:
                    continue

                sign = rece[5]
                rec = ensure(recv_team)
                rec["Tot"] += 1

                if sign == "#":
                    rec["Perf"] += 1
                elif sign == "+":
                    rec["Pos"] += 1
                elif sign == "!":
                    rec["Escl"] += 1
                elif sign == "-":
                    rec["Neg"] += 1
                elif sign == "/":
                    rec["Mez"] += 1
                elif sign == "=":
                    rec["Err"] += 1
                else:
                    # altri segni non conteggiati (eventuali)
                    pass

        df = pd.DataFrame(list(agg.values()))
        if df.empty or df["Tot"].sum() == 0:
            st.info("Nessuna ricezione trovata per i filtri selezionati.")
            return

        def pct(num, den):
            return (100.0 * num / den) if den else 0.0

        df["Perf%"] = df.apply(lambda r: pct(r["Perf"], r["Tot"]), axis=1)
        df["Pos%"]  = df.apply(lambda r: pct(r["Pos"],  r["Tot"]), axis=1)
        df["Ok%"]   = df["Perf%"] + df["Pos%"]
        df["Escl%"] = df.apply(lambda r: pct(r["Escl"], r["Tot"]), axis=1)
        df["Neg%"]  = df.apply(lambda r: pct(r["Neg"],  r["Tot"]), axis=1)
        df["Mez%"]  = df.apply(lambda r: pct(r["Mez"],  r["Tot"]), axis=1)
        df["Err%"]  = df.apply(lambda r: pct(r["Err"],  r["Tot"]), axis=1)

        # Eff = (Ok*0.77 + Escl*0.55 + Neg*0.38 – Mez*0.8 - Err)/Tot*100
        # Qui usiamo conteggi (Ok_cnt etc.) e poi *100 per renderlo percentuale
        df["Ok_cnt"] = df["Perf"] + df["Pos"]

        df["EFF"] = df.apply(
            lambda r: (
                (
                    r["Ok_cnt"] * 0.77
                    + r["Escl"] * 0.55
                    + r["Neg"] * 0.38
                    - r["Mez"] * 0.8
                    - r["Err"]
                )
                / r["Tot"]
                * 100.0
            ) if r["Tot"] else 0.0,
            axis=1
        )

        df = df.sort_values(by=["EFF", "Tot"], ascending=[False, False]).reset_index(drop=True)
        df.insert(0, "Rank", range(1, len(df) + 1))

        out = df[[
            "Rank", "Team", "Tot", "EFF",
            "Perf%", "Pos%", "Ok%", "Escl%", "Neg%", "Mez%", "Err%"
        ]].rename(columns={
            "Rank": "Rank",
            "Team": "Team",
            "Tot": "Tot",
            "EFF": "EFF",
            "Perf%": "Perf",
            "Pos%": "Pos",
            "Ok%": "OK",
            "Escl%": "Escl",
            "Neg%": "Neg",
            "Mez%": "Mez",
            "Err%": "Err",
        }).copy()

        def highlight_perugia(row):
            is_perugia = "perugia" in str(row["Team"]).lower()
            style = "background-color: #fff3cd; font-weight: 800;" if is_perugia else ""
            return [style] * len(row)

        styled = (
            out.style
              .apply(highlight_perugia, axis=1)
              .set_properties(subset=["EFF"], **{"background-color": "#e7f5ff", "font-weight": "900"})
              .format({
                  "Rank": "{:.0f}",
                  "Tot": "{:.0f}",
                  "EFF": "{:.1f}",
                  "Perf": "{:.1f}",
                  "Pos": "{:.1f}",
                  "OK": "{:.1f}",
                  "Escl": "{:.1f}",
                  "Neg": "{:.1f}",
                  "Mez": "{:.1f}",
                  "Err": "{:.1f}",
              })
              .set_table_styles([
                  {"selector": "th", "props": [("font-size", "22px"), ("text-align", "left"), ("padding", "8px 10px")]},
                  {"selector": "td", "props": [("font-size", "21px"), ("padding", "8px 10px")]},
              ])
        )

        render_any_table(styled)


    elif voce == "Attacco":
        st.subheader("Attacco")

        copt1, copt2 = st.columns(2)
        with copt1:
            use_after_recv = st.checkbox("Attacco dopo Ricezione", value=True, key="fund_att_so")
        with copt2:
            use_transition = st.checkbox("Attacco di Transizione", value=True, key="fund_att_tr")

        # Regola: entrambe -> tutti; una sola -> solo quel tipo; nessuna -> tutti (fallback)
        if not use_after_recv and not use_transition:
            use_after_recv = True
            use_transition = True

        st.caption("L’efficienza dell’attacco è calcolata in questo modo: (Punti – Murate - Errori) / Tot * 100.")

        # ===== MATCHES NEL RANGE =====
        with engine.begin() as conn:
            rows = conn.execute(
                text("""
                    SELECT team_a, team_b, scout_text
                    FROM matches
                    WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        if not rows:
            st.info("Nessun match nel range selezionato.")
            return

        def parse_rallies(scout_text: str):
            if not scout_text:
                return []
            lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
            scout_lines = [ln for ln in lines if ln[0] in ("*", "a")]

            rallies = []
            current = []
            for raw in scout_lines:
                c = code6(raw)
                if not c:
                    continue
                if is_serve(c):
                    if current:
                        rallies.append(current)
                    current = [c]
                    continue
                if not current:
                    continue
                current.append(c)
                if is_home_point(c) or is_away_point(c):
                    rallies.append(current)
                    current = []
            return rallies

        def first_attack_idx_after_serve(rally: list[str]) -> int | None:
            # primo attacco della rally (home o away) dopo la battuta
            for i, c in enumerate(rally[1:], start=1):
                if len(c) >= 6 and c[3] == "A" and c[0] in ("*", "a"):
                    return i
            return None

        def is_attack_code(c: str) -> bool:
            return len(c) >= 6 and c[3] == "A" and c[0] in ("*", "a")

        def attack_sign(c: str) -> str:
            return c[5] if c and len(c) >= 6 else ""

        agg = {}
        def ensure(team: str):
            if team not in agg:
                agg[team] = {
                    "Team": team,
                    "Tot": 0,
                    "Punti": 0,
                    "Pos": 0,
                    "Escl": 0,
                    "Neg": 0,
                    "Mur": 0,  # '/'
                    "Err": 0,  # '='
                }
            return agg[team]

        for r in rows:
            ta = canonical_team(fix_team_name(r.get("team_a") or ""))
            tb = canonical_team(fix_team_name(r.get("team_b") or ""))

            rallies = parse_rallies(r.get("scout_text") or "")
            for rally in rallies:
                if not rally:
                    continue
                if not is_serve(rally[0]):
                    continue

                fa = first_attack_idx_after_serve(rally)
                if fa is None:
                    continue

                for i in range(fa, len(rally)):
                    c = rally[i]
                    if not is_attack_code(c):
                        continue

                    is_first_attack = (i == fa)

                    if is_first_attack and not use_after_recv:
                        continue
                    if (not is_first_attack) and not use_transition:
                        continue

                    # team side by prefix
                    team = ta if c[0] == "*" else tb
                    rec = ensure(team)
                    rec["Tot"] += 1

                    s = attack_sign(c)
                    if s == "#":
                        rec["Punti"] += 1
                    elif s == "+":
                        rec["Pos"] += 1
                    elif s == "!":
                        rec["Escl"] += 1
                    elif s == "-":
                        rec["Neg"] += 1
                    elif s == "/":
                        rec["Mur"] += 1
                    elif s == "=":
                        rec["Err"] += 1

        df = pd.DataFrame(list(agg.values()))
        if df.empty or df["Tot"].sum() == 0:
            st.info("Nessun attacco trovato per i filtri selezionati.")
            return

        def pct(num, den):
            return (100.0 * num / den) if den else 0.0

        df["Punti%"] = df.apply(lambda r: pct(r["Punti"], r["Tot"]), axis=1)
        df["Pos%"]   = df.apply(lambda r: pct(r["Pos"],   r["Tot"]), axis=1)
        df["Escl%"]  = df.apply(lambda r: pct(r["Escl"],  r["Tot"]), axis=1)
        df["Neg%"]   = df.apply(lambda r: pct(r["Neg"],   r["Tot"]), axis=1)
        df["Mur%"]   = df.apply(lambda r: pct(r["Mur"],   r["Tot"]), axis=1)
        df["Err%"]   = df.apply(lambda r: pct(r["Err"],   r["Tot"]), axis=1)
        df["KO%"]    = df["Mur%"] + df["Err%"]

        # Eff = (Punti - KO)/Tot*100 -> using counts (Punti - (Mur+Err))/Tot*100
        df["EFF"] = df.apply(
            lambda r: ((r["Punti"] - (r["Mur"] + r["Err"])) / r["Tot"] * 100.0) if r["Tot"] else 0.0,
            axis=1
        )

        df = df.sort_values(by=["EFF", "Tot"], ascending=[False, False]).reset_index(drop=True)
        df.insert(0, "Rank", range(1, len(df) + 1))

        out = df[[
            "Rank", "Team", "Tot", "EFF",
            "Punti%", "Pos%", "Escl%", "Neg%", "Mur%", "Err%", "KO%"
        ]].rename(columns={
            "Rank": "Rank",
            "Team": "Team",
            "Tot": "Tot",
            "EFF": "Eff",
            "Punti%": "Punti",
            "Pos%": "Pos",
            "Escl%": "Escl",
            "Neg%": "Neg",
            "Mur%": "Mur",
            "Err%": "Err",
            "KO%": "KO",
        }).copy()

        def highlight_perugia(row):
            is_perugia = "perugia" in str(row["Team"]).lower()
            style = "background-color: #fff3cd; font-weight: 800;" if is_perugia else ""
            return [style] * len(row)

        styled = (
            out.style
              .apply(highlight_perugia, axis=1)
              .set_properties(subset=["Eff"], **{"background-color": "#e7f5ff", "font-weight": "900"})
              .format({
                  "Rank": "{:.0f}",
                  "Tot": "{:.0f}",
                  "Eff": "{:.1f}",
                  "Punti": "{:.1f}",
                  "Pos": "{:.1f}",
                  "Escl": "{:.1f}",
                  "Neg": "{:.1f}",
                  "Mur": "{:.1f}",
                  "Err": "{:.1f}",
                  "KO": "{:.1f}",
              })
              .set_table_styles([
                  {"selector": "th", "props": [("font-size", "22px"), ("text-align", "left"), ("padding", "8px 10px")]},
                  {"selector": "td", "props": [("font-size", "21px"), ("padding", "8px 10px")]},
              ])
        )

        render_any_table(styled)


    elif voce == "Muro":
        st.subheader("Muro")

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            opt_neg = st.checkbox("Muro dopo Battuta negativa", value=True, key="fund_blk_neg")
        with c2:
            opt_exc = st.checkbox("Muro dopo Battuta Esclamativa", value=True, key="fund_blk_exc")
        with c3:
            opt_pos = st.checkbox("Muro dopo Battuta Positiva", value=True, key="fund_blk_pos")
        with c4:
            opt_tr  = st.checkbox("Muro di transizione", value=True, key="fund_blk_tr")

        # Regola: se nessuna spuntata -> tutte (fallback)
        if not (opt_neg or opt_exc or opt_pos or opt_tr):
            opt_neg = opt_exc = opt_pos = opt_tr = True

        st.caption(
            "L’efficienza del muro è calcolata in questo modo: "
            "(Vincenti*2 + Positivi*0,7 + Negativi*0,07 + Coperte*0,15 - Invasioni - ManiOut) / Tot * 100."
        )

        # ===== MATCHES NEL RANGE =====
        with engine.begin() as conn:
            rows = conn.execute(
                text("""
                    SELECT team_a, team_b, scout_text
                    FROM matches
                    WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        if not rows:
            st.info("Nessun match nel range selezionato.")
            return

        def parse_rallies(scout_text: str):
            if not scout_text:
                return []
            lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
            scout_lines = [ln for ln in lines if ln[0] in ("*", "a")]

            rallies = []
            current = []
            for raw in scout_lines:
                c = code6(raw)
                if not c:
                    continue
                if is_serve(c):
                    if current:
                        rallies.append(current)
                    current = [c]
                    continue
                if not current:
                    continue
                current.append(c)
                if is_home_point(c) or is_away_point(c):
                    rallies.append(current)
                    current = []
            return rallies

        def serve_sign(c6: str) -> str:
            return c6[5] if c6 and len(c6) >= 6 else ""

        def is_attack(c6: str) -> bool:
            return len(c6) >= 6 and c6[3] == "A" and c6[0] in ("*", "a")

        def first_attack_idx(rally: list[str], attacker_prefix: str) -> int | None:
            for i in range(1, len(rally)):
                c = rally[i]
                if is_attack(c) and c[0] == attacker_prefix:
                    return i
            return None

        def first_block_after_idx(rally: list[str], start_i: int):
            for j in range(start_i + 1, len(rally)):
                c = rally[j]
                if len(c) >= 6 and c[3] == "B" and c[0] in ("*", "a"):
                    return j, c
            return None

        agg = {}
        def ensure(team: str):
            if team not in agg:
                agg[team] = {
                    "Team": team,
                    "Tot": 0,
                    "Perf": 0,
                    "Pos": 0,
                    "Neg": 0,
                    "Cop": 0,
                    "Inv": 0,
                    "Err": 0,
                }
            return agg[team]

        for r in rows:
            ta = canonical_team(fix_team_name(r.get("team_a") or ""))
            tb = canonical_team(fix_team_name(r.get("team_b") or ""))

            rallies = parse_rallies(r.get("scout_text") or "")
            for rally in rallies:
                if not rally or not is_serve(rally[0]):
                    continue

                first = rally[0]
                sgn = serve_sign(first)

                if first.startswith("*"):
                    recv_team = tb
                    recv_prefix = "a"
                    blk_team = ta
                    blk_prefix = "*"
                elif first.startswith("a"):
                    recv_team = ta
                    recv_prefix = "*"
                    blk_team = tb
                    blk_prefix = "a"
                else:
                    continue

                fa = first_attack_idx(rally, recv_prefix)
                if fa is None:
                    continue

                fb = first_block_after_idx(rally, fa)
                is_after_first_attack = False
                block_code = None
                if fb:
                    _, bc = fb
                    if bc[0] == blk_prefix:
                        is_after_first_attack = True
                        block_code = bc

                # Transizione: tutti i tocchi muro della squadra al muro che NON sono quel "primo muro dopo primo attacco"
                if opt_tr:
                    for c in rally[1:]:
                        if len(c) >= 6 and c[3] == "B" and c[0] == blk_prefix:
                            if is_after_first_attack and block_code is not None and c == block_code:
                                continue
                            rec = ensure(blk_team)
                            rec["Tot"] += 1
                            sign = c[5]
                            if sign == "#":
                                rec["Perf"] += 1
                            elif sign == "+":
                                rec["Pos"] += 1
                            elif sign == "-":
                                rec["Neg"] += 1
                            elif sign == "!":
                                rec["Cop"] += 1
                            elif sign == "/":
                                rec["Inv"] += 1
                            elif sign == "=":
                                rec["Err"] += 1

                # Dopo primo attacco, filtrato per segno battuta
                if is_after_first_attack and block_code is not None:
                    if (sgn == "-" and opt_neg) or (sgn == "!" and opt_exc) or (sgn == "+" and opt_pos):
                        rec = ensure(blk_team)
                        rec["Tot"] += 1
                        sign = block_code[5]
                        if sign == "#":
                            rec["Perf"] += 1
                        elif sign == "+":
                            rec["Pos"] += 1
                        elif sign == "-":
                            rec["Neg"] += 1
                        elif sign == "!":
                            rec["Cop"] += 1
                        elif sign == "/":
                            rec["Inv"] += 1
                        elif sign == "=":
                            rec["Err"] += 1

        df = pd.DataFrame(list(agg.values()))
        if df.empty or df["Tot"].sum() == 0:
            st.info("Nessun muro trovato per i filtri selezionati.")
            return

        def pct(num, den):
            return (100.0 * num / den) if den else 0.0

        df["Perf%"] = df.apply(lambda r: pct(r["Perf"], r["Tot"]), axis=1)
        df["Pos%"]  = df.apply(lambda r: pct(r["Pos"],  r["Tot"]), axis=1)
        df["Neg%"]  = df.apply(lambda r: pct(r["Neg"],  r["Tot"]), axis=1)
        df["Cop%"]  = df.apply(lambda r: pct(r["Cop"],  r["Tot"]), axis=1)
        df["Inv%"]  = df.apply(lambda r: pct(r["Inv"],  r["Tot"]), axis=1)
        df["Err%"]  = df.apply(lambda r: pct(r["Err"],  r["Tot"]), axis=1)

        df["EFF"] = df.apply(
            lambda r: (
                (r["Perf"] * 2.0 + r["Pos"] * 0.7 + r["Neg"] * 0.07 + r["Cop"] * 0.15 - r["Inv"] - r["Err"])
                / r["Tot"]
                * 100.0
            ) if r["Tot"] else 0.0,
            axis=1
        )

        df = df.sort_values(by=["EFF", "Tot"], ascending=[False, False]).reset_index(drop=True)
        df.insert(0, "Rank", range(1, len(df) + 1))

        out = df[[
            "Rank", "Team", "Tot", "EFF",
            "Perf%", "Pos%", "Neg%", "Cop%", "Inv%", "Err%"
        ]].rename(columns={
            "Rank": "Rank",
            "Team": "Team",
            "Tot": "Tot",
            "EFF": "Eff",
            "Perf%": "Perf",
            "Pos%": "Pos",
            "Neg%": "Neg",
            "Cop%": "Cop",
            "Inv%": "Inv",
            "Err%": "Err",
        }).copy()

        def highlight_perugia(row):
            is_perugia = "perugia" in str(row["Team"]).lower()
            style = "background-color: #fff3cd; font-weight: 800;" if is_perugia else ""
            return [style] * len(row)

        styled = (
            out.style
              .apply(highlight_perugia, axis=1)
              .set_properties(subset=["Eff"], **{"background-color": "#e7f5ff", "font-weight": "900"})
              .format({
                  "Rank": "{:.0f}",
                  "Tot": "{:.0f}",
                  "Eff": "{:.1f}",
                  "Perf": "{:.1f}",
                  "Pos": "{:.1f}",
                  "Neg": "{:.1f}",
                  "Cop": "{:.1f}",
                  "Inv": "{:.1f}",
                  "Err": "{:.1f}",
              })
              .set_table_styles([
                  {"selector": "th", "props": [("font-size", "22px"), ("text-align", "left"), ("padding", "8px 10px")]},
                  {"selector": "td", "props": [("font-size", "21px"), ("padding", "8px 10px")]},
              ])
        )

        render_any_table(styled)

    elif voce == "Difesa":
        st.subheader("Difesa")

        # Toggle (come giocatori): totale oppure per tipologia DT/DQ/DM/DH
        if "fund_tm_def_total" not in st.session_state:
            st.session_state.fund_tm_def_total = True
        for k in ("dt", "dq", "dm", "dh"):
            kk = f"fund_tm_def_{k}"
            if kk not in st.session_state:
                st.session_state[kk] = False

        def _toggle_total_tm():
            if st.session_state.fund_tm_def_total:
                st.session_state.fund_tm_def_dt = False
                st.session_state.fund_tm_def_dq = False
                st.session_state.fund_tm_def_dm = False
                st.session_state.fund_tm_def_dh = False

        def _toggle_specific_tm():
            if (st.session_state.fund_tm_def_dt or st.session_state.fund_tm_def_dq or
                st.session_state.fund_tm_def_dm or st.session_state.fund_tm_def_dh):
                st.session_state.fund_tm_def_total = False
            if (not st.session_state.fund_tm_def_total and
                not st.session_state.fund_tm_def_dt and not st.session_state.fund_tm_def_dq and
                not st.session_state.fund_tm_def_dm and not st.session_state.fund_tm_def_dh):
                st.session_state.fund_tm_def_total = True

        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            st.checkbox("Difesa su palla Spinta", key="fund_tm_def_dt", on_change=_toggle_specific_tm)
        with c2:
            st.checkbox("Difesa su 1°tempo", key="fund_tm_def_dq", on_change=_toggle_specific_tm)
        with c3:
            st.checkbox("Difesa su PIPE", key="fund_tm_def_dm", on_change=_toggle_specific_tm)
        with c4:
            st.checkbox("Difesa su H-ball", key="fund_tm_def_dh", on_change=_toggle_specific_tm)
        with c5:
            st.checkbox("Difesa TOTALE", key="fund_tm_def_total", on_change=_toggle_total_tm)

        st.caption(
            "L’efficienza della difesa è calcolata così: "
            "(Perf*2 + Pos*0,7 + Neg*0,07 + Cop*0,15 – Inv – Err) / Tot * 100."
        )

        # ===== MATCHES NEL RANGE =====
        with engine.begin() as conn:
            rows = conn.execute(
                text("""
                    SELECT team_a, team_b, scout_text
                    FROM matches
                    WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) )
                          BETWEEN :from_round AND :to_round
                """),
                {"from_round": int(from_round), "to_round": int(to_round)}
            ).mappings().all()

        if not rows:
            st.info("Nessun match nel range selezionato.")
            return

        def parse_rallies_local(scout_text: str):
            if not scout_text:
                return []
            lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
            scout_lines = [ln for ln in lines if ln[0] in ("*", "a")]

            rallies = []
            current = []
            for raw in scout_lines:
                c = code6(raw)
                if not c:
                    continue
                if is_serve(c):
                    if current:
                        rallies.append(current)
                    current = [c]
                    continue
                if not current:
                    continue
                current.append(c)
                if is_home_point(c) or is_away_point(c):
                    rallies.append(current)
                    current = []
            return rallies

        def is_defense(c6: str) -> bool:
            return len(c6) >= 6 and c6[3] == "D" and c6[0] in ("*", "a")

        def def_type(c6: str) -> str:
            return c6[3:5] if c6 and len(c6) >= 5 else ""

        def def_sign_local(c6: str) -> str:
            return c6[5] if c6 and len(c6) >= 6 else ""

        def defense_selected(c6: str) -> bool:
            if st.session_state.fund_tm_def_total:
                return True
            t = def_type(c6)
            return (
                (st.session_state.fund_tm_def_dt and t == "DT") or
                (st.session_state.fund_tm_def_dq and t == "DQ") or
                (st.session_state.fund_tm_def_dm and t == "DM") or
                (st.session_state.fund_tm_def_dh and t == "DH")
            )

        agg = {}  # team(city) -> counts

        def ensure(team_city: str):
            if team_city not in agg:
                agg[team_city] = {
                    "Team": team_city,
                    "Tot": 0,
                    "Perf": 0,  # '#'
                    "Pos": 0,   # '+'
                    "Neg": 0,   # '-'
                    "Cop": 0,   # '!'
                    "Inv": 0,   # '/'
                    "Err": 0,   # '='
                }
            return agg[team_city]

        def pct_local(num, den):
            return (100.0 * num / den) if den else 0.0

        for r in rows:
            ta_raw = fix_team_name(r.get("team_a") or "")
            tb_raw = fix_team_name(r.get("team_b") or "")
            ta = canonical_team(ta_raw)
            tb = canonical_team(tb_raw)

            rallies = parse_rallies_local(r.get("scout_text") or "")
            for rally in rallies:
                if not rally:
                    continue
                for c in rally[1:]:
                    if not is_defense(c):
                        continue
                    if not defense_selected(c):
                        continue

                    team_city = ta if c[0] == "*" else tb
                    rec = ensure(team_city)
                    rec["Tot"] += 1

                    s = def_sign_local(c)
                    if s == "#":
                        rec["Perf"] += 1
                    elif s == "+":
                        rec["Pos"] += 1
                    elif s == "!":
                        rec["Cop"] += 1
                    elif s == "-":
                        rec["Neg"] += 1
                    elif s == "/":
                        rec["Inv"] += 1
                    elif s == "=":
                        rec["Err"] += 1

        df = pd.DataFrame(list(agg.values()))
        if df.empty:
            st.info("Nessuna difesa trovata per i filtri selezionati.")
            return

        df["Perf%"] = df.apply(lambda r: pct_local(r["Perf"], r["Tot"]), axis=1)
        df["Pos%"]  = df.apply(lambda r: pct_local(r["Pos"],  r["Tot"]), axis=1)
        df["Neg%"]  = df.apply(lambda r: pct_local(r["Neg"],  r["Tot"]), axis=1)
        df["Cop%"]  = df.apply(lambda r: pct_local(r["Cop"],  r["Tot"]), axis=1)
        df["Inv%"]  = df.apply(lambda r: pct_local(r["Inv"],  r["Tot"]), axis=1)
        df["Err%"]  = df.apply(lambda r: pct_local(r["Err"],  r["Tot"]), axis=1)

        df["EFF"] = df.apply(
            lambda r: (
                (r["Perf"] * 2.0 + r["Pos"] * 0.7 + r["Neg"] * 0.07 + r["Cop"] * 0.15 - r["Inv"] - r["Err"])
                / r["Tot"]
                * 100.0
            ) if r["Tot"] else 0.0,
            axis=1
        )

        df = df.sort_values(by=["EFF", "Tot"], ascending=[False, False]).reset_index(drop=True).head(12)
        df.insert(0, "Rank", range(1, len(df) + 1))

        out = df[[
            "Rank", "Team", "Tot", "EFF",
            "Perf%", "Pos%", "Neg%", "Cop%", "Inv%", "Err%"
        ]].rename(columns={
            "EFF": "Eff",
            "Perf%": "Perf",
            "Pos%": "Pos",
            "Neg%": "Neg",
            "Cop%": "Cop",
            "Inv%": "Inv",
            "Err%": "Err",
        }).copy()

        def highlight_perugia(row):
            is_perugia = "perugia" in str(row["Team"]).lower()
            style = "background-color: #fff3cd; font-weight: 800;" if is_perugia else ""
            return [style] * len(row)

        styled = (
            out.style
              .apply(highlight_perugia, axis=1)
              .set_properties(subset=["Eff"], **{"background-color": "#e7f5ff", "font-weight": "900"})
              .format({
                  "Rank": "{:.0f}",
                  "Tot": "{:.0f}",
                  "Eff": "{:.1f}",
                  "Perf": "{:.1f}",
                  "Pos": "{:.1f}",
                  "Neg": "{:.1f}",
                  "Cop": "{:.1f}",
                  "Inv": "{:.1f}",
                  "Err": "{:.1f}",
              })
              .set_table_styles([
                  {"selector": "th", "props": [("font-size", "22px"), ("text-align", "left"), ("padding", "8px 10px")]},
                  {"selector": "td", "props": [("font-size", "21px"), ("padding", "8px 10px")]},
              ])
        )
        st.dataframe(styled, width="stretch", hide_index=True)
        return



# =========================
# UI: CLASSIFICHE FONDAMENTALI - GIOCATORI (per ruolo)
# =========================
def render_fondamentali_players():
    st.header("Classifiche Fondamentali - Giocatori (per ruolo)")

    fondamentale = st.radio(
        "Seleziona fondamentale",
        ["Battuta", "Ricezione", "Attacco", "Muro", "Difesa"],
        index=0,
        key="fund_pl_fond",
    )

    # ===== RANGE GIORNATE =====
    with engine.begin() as conn:
        bounds = conn.execute(text("""
            SELECT MIN(round_number) AS min_r, MAX(round_number) AS max_r
            FROM matches
            WHERE round_number IS NOT NULL
        """)).mappings().first()

    min_r = int((bounds["min_r"] or 1))
    max_r = int((bounds["max_r"] or 1))

    from_round, to_round, _range_label = get_selected_match_range()
    st.caption(f"📌 Range: {_range_label}")

    if from_round > to_round:
        st.error("Range non valido: 'Da giornata' deve essere <= 'A giornata'.")
        st.stop()

    # ===== ROSTER / RUOLI =====
    season = st.text_input("Stagione roster", value="2025-26", key="fund_pl_season")

    with engine.begin() as conn:
        roster_rows = conn.execute(text("""
            SELECT team_raw, team_norm, jersey_number, player_name, role, created_at
            FROM roster
            WHERE season = :season
        """), {"season": season}).mappings().all()

    if not roster_rows:
        st.warning("Roster vuoto per questa stagione: importa prima i ruoli (pagina Import Ruoli).")
        return

    df_roster = pd.DataFrame(roster_rows)
    df_roster["created_at"] = df_roster["created_at"].fillna("")
    df_roster = (
        df_roster.sort_values(by=["team_norm", "jersey_number", "created_at"])
                 .drop_duplicates(subset=["team_norm", "jersey_number"], keep="last")
    )

    roles_all = sorted(df_roster["role"].dropna().unique().tolist())
    roles_sel = st.multiselect(
        "Filtra per ruolo (selezione multipla)",
        options=roles_all,
        default=roles_all,
        key="fund_pl_roles",
    )

    min_hits = st.number_input(
        "Numero minimo di colpi (sotto questo valore il giocatore non viene mostrato)",
        min_value=0,
        max_value=1000,
        value=10,
        step=1,
        key="fund_pl_min_hits",
    )

    # ===== MATCHES NEL RANGE =====
    with engine.begin() as conn:
        matches = conn.execute(text("""
            SELECT team_a, team_b, scout_text
            FROM matches
            WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
        """), {"from_round": int(from_round), "to_round": int(to_round)}).mappings().all()

    if not matches:
        st.info("Nessun match nel range selezionato.")
        return

    def pct(num, den):
        return (100.0 * num / den) if den else 0.0

    def highlight_perugia(row):
        is_perugia = "perugia" in str(row.get("Squadra", "")).lower()
        style = "background-color: #fff3cd; font-weight: 800;" if is_perugia else ""
        return [style] * len(row)

    # =======================
    # BATTUTA
    # =======================
    if fondamentale == "Battuta":
        cb1, cb2 = st.columns(2)
        with cb1:
            use_spin = st.checkbox("Battuta SPIN", value=True, key="fund_pl_srv_spin")
        with cb2:
            use_float = st.checkbox("Battuta FLOAT", value=True, key="fund_pl_srv_float")

        if not use_spin and not use_float:
            use_spin = True
            use_float = True

        allowed_types = set()
        if use_spin:
            allowed_types.add("SQ")
        if use_float:
            allowed_types.add("SM")

        st.caption(
            "L’efficienza della battuta è calcolata in questo modo: "
            "(Punti + ½ P.*0,8 + Pos.*0,45 + Esc*0,3 + Neg*0,15 – Err) / Tot * 100. "
            "Dove i coefficienti sono le % medie di B.Point con quel tipo di ricezione del campionato."
        )

        agg = {}

        def ensure(team_raw: str, num: int):
            tnorm = team_norm(team_raw)
            key = (tnorm, num)
            if key not in agg:
                agg[key] = {
                    "team_norm": tnorm,
                    "Squadra": canonical_team(team_raw),
                    "N°": num,
                    "Tot": 0,
                    "Punti": 0,
                    "Half": 0,
                    "Err": 0,
                    "Pos": 0,
                    "Esc": 0,
                    "Neg": 0,
                }
            return agg[key]

        def serve_type(c6: str) -> str:
            return c6[3:5] if c6 and len(c6) >= 5 else ""

        def serve_sign_local(c6: str) -> str:
            return c6[5] if c6 and len(c6) >= 6 else ""

        for m in matches:
            ta = fix_team_name(m.get("team_a") or "")
            tb = fix_team_name(m.get("team_b") or "")
            scout_text = m.get("scout_text") or ""
            if not scout_text:
                continue

            for raw in str(scout_text).splitlines():
                raw = raw.strip()
                if not raw or raw[0] not in ("*", "a"):
                    continue
                c6 = code6(raw)
                if not c6 or not is_serve(c6):
                    continue

                stype = serve_type(c6)
                if stype not in allowed_types:
                    continue

                team = ta if c6[0] == "*" else tb

                num = serve_player_number(c6)
                if num is None:
                    continue

                rec = ensure(team, num)
                rec["Tot"] += 1

                sgn = serve_sign_local(c6)
                if sgn == "#":
                    rec["Punti"] += 1
                elif sgn == "/":
                    rec["Half"] += 1
                elif sgn == "=":
                    rec["Err"] += 1
                elif sgn == "+":
                    rec["Pos"] += 1
                elif sgn == "!":
                    rec["Esc"] += 1
                elif sgn == "-":
                    rec["Neg"] += 1

        df = pd.DataFrame(list(agg.values()))
        if df.empty:
            st.info("Nessuna battuta trovata per i filtri selezionati.")
            return

        df = df.merge(
            df_roster[["team_norm", "jersey_number", "player_name", "role", "team_raw"]],
            left_on=["team_norm", "N°"],
            right_on=["team_norm", "jersey_number"],
            how="left",
        ).drop(columns=["jersey_number"])

        df.rename(columns={"player_name": "Nome giocatore", "role": "Ruolo"}, inplace=True)
        df["Nome giocatore"] = df["Nome giocatore"].fillna(df["N°"].apply(lambda x: f"N°{int(x):02d}"))
        df["Ruolo"] = df["Ruolo"].fillna("(non in roster)")
        df["Squadra"] = df["team_raw"].apply(canonical_team).fillna(df["Squadra"])
        df = df.drop(columns=["team_raw"])

        # Canonicalizza sempre la colonna Squadra (solo città)
        df["Squadra"] = df["Squadra"].apply(canonical_team)


        if roles_sel:
            df = df[df["Ruolo"].isin(roles_sel)].copy()

        df = df[df["Tot"] >= int(min_hits)].copy()
        if df.empty:
            st.info("Nessun giocatore supera il filtro del numero minimo di colpi.")
            return

        df["Punti%"] = df.apply(lambda r: pct(r["Punti"], r["Tot"]), axis=1)
        df["Half%"]  = df.apply(lambda r: pct(r["Half"],  r["Tot"]), axis=1)
        df["Err%"]   = df.apply(lambda r: pct(r["Err"],   r["Tot"]), axis=1)
        df["Pos%"]   = df.apply(lambda r: pct(r["Pos"],   r["Tot"]), axis=1)
        df["Esc%"]   = df.apply(lambda r: pct(r["Esc"],   r["Tot"]), axis=1)
        df["Neg%"]   = df.apply(lambda r: pct(r["Neg"],   r["Tot"]), axis=1)

        df["EFF"] = df.apply(
            lambda r: (
                (r["Punti"] + r["Half"] * 0.8 + r["Pos"] * 0.45 + r["Esc"] * 0.3 + r["Neg"] * 0.15 - r["Err"])
                / r["Tot"]
                * 100.0
            ) if r["Tot"] else 0.0,
            axis=1
        )

        df = df.sort_values(by=["EFF", "Tot"], ascending=[False, False]).reset_index(drop=True)
        df.insert(0, "Rank", range(1, len(df) + 1))

        out = df[[
            "Rank", "Nome giocatore", "Squadra", "Tot", "EFF",
            "Punti%", "Half%", "Err%", "Pos%", "Esc%", "Neg%"
        ]].rename(columns={
            "Tot": "Tot",
            "EFF": "EFF",
            "Punti%": "Punti",
            "Half%": "½ P",
            "Err%": "Err.",
            "Pos%": "Pos",
            "Esc%": "Esc",
            "Neg%": "Neg",
        }).copy()

        styled = (
            out.style
              .apply(highlight_perugia, axis=1)
              .set_properties(subset=["EFF"], **{"background-color": "#e7f5ff", "font-weight": "900"})
              .format({
                  "Rank": "{:.0f}",
                  "Tot": "{:.0f}",
                  "EFF": "{:.1f}",
                  "Punti": "{:.1f}",
                  "½ P": "{:.1f}",
                  "Err.": "{:.1f}",
                  "Pos": "{:.1f}",
                  "Esc": "{:.1f}",
                  "Neg": "{:.1f}",
              })
              .set_table_styles([
                  {"selector": "th", "props": [("font-size", "22px"), ("text-align", "left"), ("padding", "8px 10px")]},
                  {"selector": "td", "props": [("font-size", "21px"), ("padding", "8px 10px")]},
              ])
        )
        render_focus_4_players(out, key_base=f"fund_{fondamentale}_{from_round}_{to_round}")

        render_any_table(styled)
        return

    # =======================
    # RICEZIONE
    # =======================
    if fondamentale == "Ricezione":
        cb1, cb2 = st.columns(2)
        with cb1:
            use_spin = st.checkbox("Ricezione SPIN", value=True, key="fund_pl_rec_spin")
        with cb2:
            use_float = st.checkbox("Ricezione FLOAT", value=True, key="fund_pl_rec_float")

        if not use_spin and not use_float:
            use_spin = True
            use_float = True

        allowed_types = set()
        if use_spin:
            allowed_types.add("SQ")
        if use_float:
            allowed_types.add("SM")

        st.caption(
            "L’efficienza della ricezione è calcolata in questo modo: "
            "(Ok*0,77 + Escl*0,55 + Neg*0,38 – Mez*0,8 - Err) / Tot * 100. "
            "Dove i coefficienti sono le % medie di Side Out con quel tipo di ricezione del campionato."
        )

        def parse_rallies(scout_text: str):
            if not scout_text:
                return []
            lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
            scout_lines = [ln for ln in lines if ln[0] in ("*", "a")]

            rallies = []
            current = []
            for raw in scout_lines:
                c = code6(raw)
                if not c:
                    continue
                if is_serve(c):
                    if current:
                        rallies.append(current)
                    current = [c]
                    continue
                if not current:
                    continue
                current.append(c)
                if is_home_point(c) or is_away_point(c):
                    rallies.append(current)
                    current = []
            return rallies

        def serve_type(first: str) -> str:
            return first[3:5] if first and len(first) >= 5 else ""

        def first_reception(rally: list[str], prefix: str):
            for c in rally:
                if len(c) >= 6 and c[0] == prefix and c[3:5] in ("RQ", "RM"):
                    return c
            return None

        agg = {}

        def ensure(team_raw: str, num: int):
            tnorm = team_norm(team_raw)
            key = (tnorm, num)
            if key not in agg:
                agg[key] = {
                    "team_norm": tnorm,
                    "Squadra": canonical_team(team_raw),
                    "N°": num,
                    "Tot": 0,
                    "Perf": 0,
                    "Pos": 0,
                    "Escl": 0,
                    "Neg": 0,
                    "Mez": 0,
                    "Err": 0,
                }
            return agg[key]

        for m in matches:
            ta = fix_team_name(m.get("team_a") or "")
            tb = fix_team_name(m.get("team_b") or "")
            rallies = parse_rallies(m.get("scout_text") or "")
            for rally in rallies:
                if not rally or not is_serve(rally[0]):
                    continue
                first = rally[0]
                stype = serve_type(first)
                if stype not in allowed_types:
                    continue

                if first.startswith("*"):
                    recv_team = tb
                    recv_prefix = "a"
                else:
                    recv_team = ta
                    recv_prefix = "*"

                rece = first_reception(rally, recv_prefix)
                if not rece:
                    continue

                num = serve_player_number(rece)
                if num is None:
                    continue

                sign = rece[5]
                rec = ensure(recv_team, num)
                rec["Tot"] += 1
                if sign == "#":
                    rec["Perf"] += 1
                elif sign == "+":
                    rec["Pos"] += 1
                elif sign == "!":
                    rec["Escl"] += 1
                elif sign == "-":
                    rec["Neg"] += 1
                elif sign == "/":
                    rec["Mez"] += 1
                elif sign == "=":
                    rec["Err"] += 1

        df = pd.DataFrame(list(agg.values()))
        if df.empty:
            st.info("Nessuna ricezione trovata per i filtri selezionati.")
            return

        df = df.merge(
            df_roster[["team_norm", "jersey_number", "player_name", "role", "team_raw"]],
            left_on=["team_norm", "N°"],
            right_on=["team_norm", "jersey_number"],
            how="left",
        ).drop(columns=["jersey_number"])

        df.rename(columns={"player_name": "Nome giocatore", "role": "Ruolo"}, inplace=True)
        df["Nome giocatore"] = df["Nome giocatore"].fillna(df["N°"].apply(lambda x: f"N°{int(x):02d}"))
        df["Ruolo"] = df["Ruolo"].fillna("(non in roster)")
        df["Squadra"] = df["team_raw"].apply(canonical_team).fillna(df["Squadra"])
        df = df.drop(columns=["team_raw"])

        # Canonicalizza sempre la colonna Squadra (solo città)
        df["Squadra"] = df["Squadra"].apply(canonical_team)


        if roles_sel:
            df = df[df["Ruolo"].isin(roles_sel)].copy()

        df = df[df["Tot"] >= int(min_hits)].copy()
        if df.empty:
            st.info("Nessun giocatore supera il filtro del numero minimo di colpi.")
            return

        df["Perf%"] = df.apply(lambda r: pct(r["Perf"], r["Tot"]), axis=1)
        df["Pos%"]  = df.apply(lambda r: pct(r["Pos"],  r["Tot"]), axis=1)
        df["OK%"]   = df["Perf%"] + df["Pos%"]
        df["Escl%"] = df.apply(lambda r: pct(r["Escl"], r["Tot"]), axis=1)
        df["Neg%"]  = df.apply(lambda r: pct(r["Neg"],  r["Tot"]), axis=1)
        df["Mez%"]  = df.apply(lambda r: pct(r["Mez"],  r["Tot"]), axis=1)
        df["Err%"]  = df.apply(lambda r: pct(r["Err"],  r["Tot"]), axis=1)

        ok_cnt = df["Perf"] + df["Pos"]
        df["EFF"] = ((ok_cnt * 0.77 + df["Escl"] * 0.55 + df["Neg"] * 0.38 - df["Mez"] * 0.8 - df["Err"]) / df["Tot"]) * 100.0

        df = df.sort_values(by=["EFF", "Tot"], ascending=[False, False]).reset_index(drop=True)
        df.insert(0, "Rank", range(1, len(df) + 1))

        out = df[[
            "Rank", "Nome giocatore", "Squadra", "Tot", "EFF",
            "Perf%", "Pos%", "OK%", "Escl%", "Neg%", "Mez%", "Err%"
        ]].rename(columns={
            "Tot": "Tot",
            "EFF": "Eff",
            "Perf%": "Perf",
            "Pos%": "Pos",
            "OK%": "OK",
            "Escl%": "Escl",
            "Neg%": "Neg",
            "Mez%": "Mez",
            "Err%": "Err",
        }).copy()

        styled = (
            out.style
              .apply(highlight_perugia, axis=1)
              .set_properties(subset=["Eff"], **{"background-color": "#e7f5ff", "font-weight": "900"})
              .format({
                  "Rank": "{:.0f}",
                  "Tot": "{:.0f}",
                  "Eff": "{:.1f}",
                  "Perf": "{:.1f}",
                  "Pos": "{:.1f}",
                  "OK": "{:.1f}",
                  "Escl": "{:.1f}",
                  "Neg": "{:.1f}",
                  "Mez": "{:.1f}",
                  "Err": "{:.1f}",
              })
              .set_table_styles([
                  {"selector": "th", "props": [("font-size", "22px"), ("text-align", "left"), ("padding", "8px 10px")]},
                  {"selector": "td", "props": [("font-size", "21px"), ("padding", "8px 10px")]},
              ])
        )
        render_any_table(styled)
        return

    
    # =======================
    # ATTACCO
    # =======================
    if fondamentale == "Attacco":
        copt1, copt2 = st.columns(2)
        with copt1:
            use_after_recv = st.checkbox("Attacco dopo Ricezione", value=True, key="fund_pl_att_so")
        with copt2:
            use_transition = st.checkbox("Attacco di Transizione", value=True, key="fund_pl_att_tr")

        if not use_after_recv and not use_transition:
            use_after_recv = True
            use_transition = True

        st.caption("L’efficienza dell’attacco è calcolata in questo modo: (Punti – Murate - Errori) / Tot * 100.")

        def parse_rallies(scout_text: str):
            if not scout_text:
                return []
            lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
            scout_lines = [ln for ln in lines if ln[0] in ("*", "a")]

            rallies = []
            current = []
            for raw in scout_lines:
                c = code6(raw)
                if not c:
                    continue
                if is_serve(c):
                    if current:
                        rallies.append(current)
                    current = [c]
                    continue
                if not current:
                    continue
                current.append(c)
                if is_home_point(c) or is_away_point(c):
                    rallies.append(current)
                    current = []
            return rallies

        def is_attack_code(c: str) -> bool:
            return len(c) >= 6 and c[3] == "A" and c[0] in ("*", "a")

        def attack_sign(c: str) -> str:
            return c[5] if c and len(c) >= 6 else ""

        def first_attack_idx_after_serve(rally: list[str]) -> int | None:
            for i, c in enumerate(rally[1:], start=1):
                if is_attack_code(c):
                    return i
            return None

        agg = {}

        def ensure(team_raw: str, num: int):
            tnorm = team_norm(team_raw)
            key = (tnorm, num)
            if key not in agg:
                agg[key] = {
                    "team_norm": tnorm,
                    "Squadra": canonical_team(team_raw),
                    "N°": num,
                    "Tot": 0,
                    "Punti": 0,
                    "Pos": 0,
                    "Escl": 0,
                    "Neg": 0,
                    "Mur": 0,
                    "Err": 0,
                }
            return agg[key]

        for m in matches:
            ta = fix_team_name(m.get("team_a") or "")
            tb = fix_team_name(m.get("team_b") or "")
            rallies = parse_rallies(m.get("scout_text") or "")

            for rally in rallies:
                if not rally or not is_serve(rally[0]):
                    continue

                fa = first_attack_idx_after_serve(rally)
                if fa is None:
                    continue

                for i in range(fa, len(rally)):
                    c = rally[i]
                    if not is_attack_code(c):
                        continue

                    is_first = (i == fa)
                    if is_first and not use_after_recv:
                        continue
                    if (not is_first) and not use_transition:
                        continue

                    team_raw = ta if c[0] == "*" else tb
                    num = serve_player_number(c)
                    if num is None:
                        continue

                    rec = ensure(team_raw, num)
                    rec["Tot"] += 1
                    s = attack_sign(c)
                    if s == "#":
                        rec["Punti"] += 1
                    elif s == "+":
                        rec["Pos"] += 1
                    elif s == "!":
                        rec["Escl"] += 1
                    elif s == "-":
                        rec["Neg"] += 1
                    elif s == "/":
                        rec["Mur"] += 1
                    elif s == "=":
                        rec["Err"] += 1

        df = pd.DataFrame(list(agg.values()))
        if df.empty:
            st.info("Nessun attacco trovato per i filtri selezionati.")
            return

        df = df.merge(
            df_roster[["team_norm", "jersey_number", "player_name", "role", "team_raw"]],
            left_on=["team_norm", "N°"],
            right_on=["team_norm", "jersey_number"],
            how="left",
        ).drop(columns=["jersey_number"])

        df.rename(columns={"player_name": "Nome giocatore", "role": "Ruolo"}, inplace=True)
        df["Nome giocatore"] = df["Nome giocatore"].fillna(df["N°"].apply(lambda x: f"N°{int(x):02d}"))
        df["Ruolo"] = df["Ruolo"].fillna("(non in roster)")
        df["Squadra"] = df["team_raw"].apply(canonical_team).fillna(df["Squadra"])
        df = df.drop(columns=["team_raw"])

        # Canonicalizza sempre la colonna Squadra (solo città)
        df["Squadra"] = df["Squadra"].apply(canonical_team)


        if roles_sel:
            df = df[df["Ruolo"].isin(roles_sel)].copy()

        df = df[df["Tot"] >= int(min_hits)].copy()
        if df.empty:
            st.info("Nessun giocatore supera il filtro del numero minimo di colpi.")
            return

        df["Punti%"] = df.apply(lambda r: pct(r["Punti"], r["Tot"]), axis=1)
        df["Pos%"]   = df.apply(lambda r: pct(r["Pos"],   r["Tot"]), axis=1)
        df["Escl%"]  = df.apply(lambda r: pct(r["Escl"],  r["Tot"]), axis=1)
        df["Neg%"]   = df.apply(lambda r: pct(r["Neg"],   r["Tot"]), axis=1)
        df["Mur%"]   = df.apply(lambda r: pct(r["Mur"],   r["Tot"]), axis=1)
        df["Err%"]   = df.apply(lambda r: pct(r["Err"],   r["Tot"]), axis=1)
        df["KO%"]    = df["Mur%"] + df["Err%"]

        df["EFF"] = df.apply(
            lambda r: ((r["Punti"] - (r["Mur"] + r["Err"])) / r["Tot"] * 100.0) if r["Tot"] else 0.0,
            axis=1
        )

        df = df.sort_values(by=["EFF", "Tot"], ascending=[False, False]).reset_index(drop=True)
        df.insert(0, "Rank", range(1, len(df) + 1))

        out = df[[
            "Rank", "Nome giocatore", "Squadra", "Tot", "EFF",
            "Punti%", "Pos%", "Escl%", "Neg%", "Mur%", "Err%", "KO%"
        ]].rename(columns={
            "Tot": "Tot",
            "EFF": "Eff",
            "Punti%": "Punti",
            "Pos%": "Pos",
            "Escl%": "Escl",
            "Neg%": "Neg",
            "Mur%": "Mur",
            "Err%": "Err",
            "KO%": "KO",
        }).copy()

        styled = (
            out.style
              .apply(highlight_perugia, axis=1)
              .set_properties(subset=["Eff"], **{"background-color": "#e7f5ff", "font-weight": "900"})
              .format({
                  "Rank": "{:.0f}",
                  "Tot": "{:.0f}",
                  "Eff": "{:.1f}",
                  "Punti": "{:.1f}",
                  "Pos": "{:.1f}",
                  "Escl": "{:.1f}",
                  "Neg": "{:.1f}",
                  "Mur": "{:.1f}",
                  "Err": "{:.1f}",
                  "KO": "{:.1f}",
              })
              .set_table_styles([
                  {"selector": "th", "props": [("font-size", "22px"), ("text-align", "left"), ("padding", "8px 10px")]},
                  {"selector": "td", "props": [("font-size", "21px"), ("padding", "8px 10px")]},
              ])
        )
        render_any_table(styled)
        return

    
    # =======================
    # MURO
    # =======================
    if fondamentale == "Muro":
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            opt_neg = st.checkbox("Muro dopo Battuta negativa", value=True, key="fund_pl_blk_neg")
        with c2:
            opt_exc = st.checkbox("Muro dopo Battuta Esclamativa", value=True, key="fund_pl_blk_exc")
        with c3:
            opt_pos = st.checkbox("Muro dopo Battuta Positiva", value=True, key="fund_pl_blk_pos")
        with c4:
            opt_tr  = st.checkbox("Muro di transizione", value=True, key="fund_pl_blk_tr")

        if not (opt_neg or opt_exc or opt_pos or opt_tr):
            opt_neg = opt_exc = opt_pos = opt_tr = True

        st.caption(
            "L’efficienza del muro è calcolata in questo modo: "
            "(Vincenti*2 + Positivi*0,7 + Negativi*0,07 + Coperte*0,15 - Invasioni - ManiOut) / Tot * 100."
        )

        def parse_rallies(scout_text: str):
            if not scout_text:
                return []
            lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
            scout_lines = [ln for ln in lines if ln[0] in ("*", "a")]

            rallies = []
            current = []
            for raw in scout_lines:
                c = code6(raw)
                if not c:
                    continue
                if is_serve(c):
                    if current:
                        rallies.append(current)
                    current = [c]
                    continue
                if not current:
                    continue
                current.append(c)
                if is_home_point(c) or is_away_point(c):
                    rallies.append(current)
                    current = []
            return rallies

        def serve_sign(c6: str) -> str:
            return c6[5] if c6 and len(c6) >= 6 else ""

        def is_attack(c6: str) -> bool:
            return len(c6) >= 6 and c6[3] == "A" and c6[0] in ("*", "a")

        def first_attack_idx(rally: list[str], attacker_prefix: str):
            for i in range(1, len(rally)):
                c = rally[i]
                if is_attack(c) and c[0] == attacker_prefix:
                    return i
            return None

        def first_block_after_idx(rally: list[str], start_i: int):
            for j in range(start_i + 1, len(rally)):
                c = rally[j]
                if len(c) >= 6 and c[3] == "B" and c[0] in ("*", "a"):
                    return j, c
            return None

        agg = {}  # (team_norm, num) -> counts

        def ensure(team_raw: str, num: int):
            tnorm = team_norm(team_raw)
            key = (tnorm, num)
            if key not in agg:
                agg[key] = {
                    "team_norm": tnorm,
                    "Squadra": canonical_team(team_raw),
                    "N°": num,
                    "Tot": 0,
                    "Perf": 0,
                    "Pos": 0,
                    "Neg": 0,
                    "Cop": 0,
                    "Inv": 0,
                    "Err": 0,
                }
            return agg[key]

        def add_block(team_raw: str, num: int, block_code: str):
            rec = ensure(team_raw, num)
            rec["Tot"] += 1
            sign = block_code[5] if len(block_code) >= 6 else ""
            if sign == "#":
                rec["Perf"] += 1
            elif sign == "+":
                rec["Pos"] += 1
            elif sign == "-":
                rec["Neg"] += 1
            elif sign == "!":
                rec["Cop"] += 1
            elif sign == "/":
                rec["Inv"] += 1
            elif sign == "=":
                rec["Err"] += 1

        for m in matches:
            ta = fix_team_name(m.get("team_a") or "")
            tb = fix_team_name(m.get("team_b") or "")
            rallies = parse_rallies(m.get("scout_text") or "")

            for rally in rallies:
                if not rally or not is_serve(rally[0]):
                    continue

                first = rally[0]
                sgn = serve_sign(first)

                if first.startswith("*"):
                    recv_team = tb
                    recv_prefix = "a"
                    blk_team = ta
                    blk_prefix = "*"
                else:
                    recv_team = ta
                    recv_prefix = "*"
                    blk_team = tb
                    blk_prefix = "a"

                fa = first_attack_idx(rally, recv_prefix)
                if fa is None:
                    continue

                fb = first_block_after_idx(rally, fa)
                is_after_first_attack = False
                block_code = None
                if fb:
                    _, bc = fb
                    if bc[0] == blk_prefix:
                        is_after_first_attack = True
                        block_code = bc

                # Transizione: tutti i tocchi muro della squadra al muro che NON sono quel "primo muro dopo primo attacco"
                if opt_tr:
                    for c in rally[1:]:
                        if len(c) >= 6 and c[3] == "B" and c[0] == blk_prefix:
                            if is_after_first_attack and block_code is not None and c == block_code:
                                continue
                            num = serve_player_number(c)
                            if num is None:
                                continue
                            add_block(blk_team, num, c)

                # Dopo primo attacco, filtrato per segno battuta
                if is_after_first_attack and block_code is not None:
                    if (sgn == "-" and opt_neg) or (sgn == "!" and opt_exc) or (sgn == "+" and opt_pos):
                        num = serve_player_number(block_code)
                        if num is None:
                            continue
                        add_block(blk_team, num, block_code)

        df = pd.DataFrame(list(agg.values()))
        if df.empty:
            st.info("Nessun muro trovato per i filtri selezionati.")
            return

        df = df.merge(
            df_roster[["team_norm", "jersey_number", "player_name", "role", "team_raw"]],
            left_on=["team_norm", "N°"],
            right_on=["team_norm", "jersey_number"],
            how="left",
        ).drop(columns=["jersey_number"])

        df.rename(columns={"player_name": "Nome giocatore", "role": "Ruolo"}, inplace=True)
        df["Nome giocatore"] = df["Nome giocatore"].fillna(df["N°"].apply(lambda x: f"N°{int(x):02d}"))
        df["Ruolo"] = df["Ruolo"].fillna("(non in roster)")
        df["Squadra"] = df["team_raw"].apply(canonical_team).fillna(df["Squadra"])
        df = df.drop(columns=["team_raw"])

        # Canonicalizza sempre la colonna Squadra (solo città)
        df["Squadra"] = df["Squadra"].apply(canonical_team)


        if roles_sel:
            df = df[df["Ruolo"].isin(roles_sel)].copy()

        df = df[df["Tot"] >= int(min_hits)].copy()
        if df.empty:
            st.info("Nessun giocatore supera il filtro del numero minimo di colpi.")
            return

        df["Perf%"] = df.apply(lambda r: pct(r["Perf"], r["Tot"]), axis=1)
        df["Pos%"]  = df.apply(lambda r: pct(r["Pos"],  r["Tot"]), axis=1)
        df["Neg%"]  = df.apply(lambda r: pct(r["Neg"],  r["Tot"]), axis=1)
        df["Cop%"]  = df.apply(lambda r: pct(r["Cop"],  r["Tot"]), axis=1)
        df["Inv%"]  = df.apply(lambda r: pct(r["Inv"],  r["Tot"]), axis=1)
        df["Err%"]  = df.apply(lambda r: pct(r["Err"],  r["Tot"]), axis=1)

        df["EFF"] = df.apply(
            lambda r: (
                (r["Perf"] * 2.0 + r["Pos"] * 0.7 + r["Neg"] * 0.07 + r["Cop"] * 0.15 - r["Inv"] - r["Err"])
                / r["Tot"]
                * 100.0
            ) if r["Tot"] else 0.0,
            axis=1
        )

        df = df.sort_values(by=["EFF", "Tot"], ascending=[False, False]).reset_index(drop=True)
        df.insert(0, "Rank", range(1, len(df) + 1))

        out = df[[
            "Rank", "Nome giocatore", "Squadra", "Tot", "EFF",
            "Perf%", "Pos%", "Neg%", "Cop%", "Inv%", "Err%"
        ]].rename(columns={
            "Tot": "Tot",
            "EFF": "Eff",
            "Perf%": "Perf",
            "Pos%": "Pos",
            "Neg%": "Neg",
            "Cop%": "Cop",
            "Inv%": "Inv",
            "Err%": "Err",
        }).copy()

        styled = (
            out.style
              .apply(highlight_perugia, axis=1)
              .set_properties(subset=["Eff"], **{"background-color": "#e7f5ff", "font-weight": "900"})
              .format({
                  "Rank": "{:.0f}",
                  "Tot": "{:.0f}",
                  "Eff": "{:.1f}",
                  "Perf": "{:.1f}",
                  "Pos": "{:.1f}",
                  "Neg": "{:.1f}",
                  "Cop": "{:.1f}",
                  "Inv": "{:.1f}",
                  "Err": "{:.1f}",
              })
              .set_table_styles([
                  {"selector": "th", "props": [("font-size", "22px"), ("text-align", "left"), ("padding", "8px 10px")]},
                  {"selector": "td", "props": [("font-size", "21px"), ("padding", "8px 10px")]},
              ])
        )
        render_any_table(styled)
        return

    
    # =======================
    # DIFESA
    # =======================
    if fondamentale == "Difesa":
        # stato iniziale
        if "fund_pl_def_total" not in st.session_state:
            st.session_state.fund_pl_def_total = True
        for k in ("dt", "dq", "dm", "dh"):
            kk = f"fund_pl_def_{k}"
            if kk not in st.session_state:
                st.session_state[kk] = False

        def _toggle_total():
            if st.session_state.fund_pl_def_total:
                st.session_state.fund_pl_def_dt = False
                st.session_state.fund_pl_def_dq = False
                st.session_state.fund_pl_def_dm = False
                st.session_state.fund_pl_def_dh = False

        def _toggle_specific():
            if (st.session_state.fund_pl_def_dt or st.session_state.fund_pl_def_dq or
                st.session_state.fund_pl_def_dm or st.session_state.fund_pl_def_dh):
                st.session_state.fund_pl_def_total = False
            if (not st.session_state.fund_pl_def_total and
                not st.session_state.fund_pl_def_dt and not st.session_state.fund_pl_def_dq and
                not st.session_state.fund_pl_def_dm and not st.session_state.fund_pl_def_dh):
                st.session_state.fund_pl_def_total = True

        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            st.checkbox("Difesa su palla Spinta", key="fund_pl_def_dt", on_change=_toggle_specific)
        with c2:
            st.checkbox("Difesa su 1°tempo", key="fund_pl_def_dq", on_change=_toggle_specific)
        with c3:
            st.checkbox("Difesa su PIPE", key="fund_pl_def_dm", on_change=_toggle_specific)
        with c4:
            st.checkbox("Difesa su H-ball", key="fund_pl_def_dh", on_change=_toggle_specific)
        with c5:
            st.checkbox("Difesa TOTALE", key="fund_pl_def_total", on_change=_toggle_total)

        st.caption(
            "L’efficienza della difesa è calcolata in questo modo: "
            "(Buone*2 + Coperture*0,5 + Negative*0,4 + OverTheNet*0,3 – Errori) / Tot * 100. "
            "È un indice motivatore."
        )

        def parse_rallies(scout_text: str):
            if not scout_text:
                return []
            lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
            scout_lines = [ln for ln in lines if ln[0] in ("*", "a")]

            rallies = []
            current = []
            for raw in scout_lines:
                c = code6(raw)
                if not c:
                    continue
                if is_serve(c):
                    if current:
                        rallies.append(current)
                    current = [c]
                    continue
                if not current:
                    continue
                current.append(c)
                if is_home_point(c) or is_away_point(c):
                    rallies.append(current)
                    current = []
            return rallies

        def is_defense(c6: str) -> bool:
            return len(c6) >= 6 and c6[3] == "D" and c6[0] in ("*", "a")

        def def_type(c6: str) -> str:
            return c6[3:5] if c6 and len(c6) >= 5 else ""

        def def_sign(c6: str) -> str:
            return c6[5] if c6 and len(c6) >= 6 else ""

        def defense_selected(c6: str) -> bool:
            if st.session_state.fund_pl_def_total:
                return True

            t = def_type(c6)
            s = def_sign(c6)

            if st.session_state.fund_pl_def_dt and t == "DT":
                return s != "!"
            if st.session_state.fund_pl_def_dq and t == "DQ":
                return s != "!"
            if st.session_state.fund_pl_def_dm and t == "DM":
                return s != "!"
            if st.session_state.fund_pl_def_dh and t == "DH":
                return s != "!"
            return False

        agg = {}  # (team_norm, num) -> counts

        def ensure(team_raw: str, num: int):
            tnorm = team_norm(team_raw)
            key = (tnorm, num)
            if key not in agg:
                agg[key] = {
                    "team_norm": tnorm,
                    "Squadra": canonical_team(team_raw),
                    "N°": num,
                    "Tot": 0,
                    "Perf": 0,  # '+'
                    "Cop": 0,   # '!'
                    "Neg": 0,   # '-'
                    "Over": 0,  # '/'
                    "Err": 0,   # '='
                }
            return agg[key]

        for m in matches:
            ta = fix_team_name(m.get("team_a") or "")
            tb = fix_team_name(m.get("team_b") or "")
            rallies = parse_rallies(m.get("scout_text") or "")
            for rally in rallies:
                if not rally:
                    continue
                for c in rally[1:]:
                    if not is_defense(c):
                        continue
                    if not defense_selected(c):
                        continue

                    team_raw = ta if c[0] == "*" else tb
                    num = serve_player_number(c)
                    if num is None:
                        continue

                    rec = ensure(team_raw, num)
                    rec["Tot"] += 1
                    s = def_sign(c)
                    if s == "+":
                        rec["Perf"] += 1
                    elif s == "!":
                        rec["Cop"] += 1
                    elif s == "-":
                        rec["Neg"] += 1
                    elif s == "/":
                        rec["Over"] += 1
                    elif s == "=":
                        rec["Err"] += 1

        df = pd.DataFrame(list(agg.values()))
        if df.empty:
            st.info("Nessuna difesa trovata per i filtri selezionati.")
            return

        df = df.merge(
            df_roster[["team_norm", "jersey_number", "player_name", "role", "team_raw"]],
            left_on=["team_norm", "N°"],
            right_on=["team_norm", "jersey_number"],
            how="left",
        ).drop(columns=["jersey_number"])

        df.rename(columns={"player_name": "Nome giocatore", "role": "Ruolo"}, inplace=True)
        df["Nome giocatore"] = df["Nome giocatore"].fillna(df["N°"].apply(lambda x: f"N°{int(x):02d}"))
        df["Ruolo"] = df["Ruolo"].fillna("(non in roster)")
        df["Squadra"] = df["team_raw"].apply(canonical_team).fillna(df["Squadra"])
        df = df.drop(columns=["team_raw"])

        # Canonicalizza sempre la colonna Squadra (solo città)
        df["Squadra"] = df["Squadra"].apply(canonical_team)


        if roles_sel:
            df = df[df["Ruolo"].isin(roles_sel)].copy()

        df = df[df["Tot"] >= int(min_hits)].copy()
        if df.empty:
            st.info("Nessun giocatore supera il filtro del numero minimo di colpi.")
            return

        df["Perf%"] = df.apply(lambda r: pct(r["Perf"], r["Tot"]), axis=1)
        df["Cop%"]  = df.apply(lambda r: pct(r["Cop"],  r["Tot"]), axis=1)
        df["Neg%"]  = df.apply(lambda r: pct(r["Neg"],  r["Tot"]), axis=1)
        df["Over%"] = df.apply(lambda r: pct(r["Over"], r["Tot"]), axis=1)
        df["Err%"]  = df.apply(lambda r: pct(r["Err"],  r["Tot"]), axis=1)

        df["EFF"] = df.apply(
            lambda r: (
                (r["Perf"] * 2.0 + r["Cop"] * 0.5 + r["Neg"] * 0.4 + r["Over"] * 0.3 - r["Err"])
                / r["Tot"]
                * 100.0
            ) if r["Tot"] else 0.0,
            axis=1
        )

        df = df.sort_values(by=["EFF", "Tot"], ascending=[False, False]).reset_index(drop=True)
        df.insert(0, "Rank", range(1, len(df) + 1))

        out = df[[
            "Rank", "Nome giocatore", "Squadra", "Tot", "EFF",
            "Perf%", "Cop%", "Neg%", "Over%", "Err%"
        ]].rename(columns={
            "Tot": "Tot",
            "EFF": "Eff",
            "Perf%": "Perf",
            "Cop%": "Cop",
            "Neg%": "Neg",
            "Over%": "Over",
            "Err%": "Err",
        }).copy()

        styled = (
            out.style
              .apply(highlight_perugia, axis=1)
              .set_properties(subset=["Eff"], **{"background-color": "#e7f5ff", "font-weight": "900"})
              .format({
                  "Rank": "{:.0f}",
                  "Tot": "{:.0f}",
                  "Eff": "{:.1f}",
                  "Perf": "{:.1f}",
                  "Cop": "{:.1f}",
                  "Neg": "{:.1f}",
                  "Over": "{:.1f}",
                  "Err": "{:.1f}",
              })
              .set_table_styles([
                  {"selector": "th", "props": [("font-size", "22px"), ("text-align", "left"), ("padding", "8px 10px")]},
                  {"selector": "td", "props": [("font-size", "21px"), ("padding", "8px 10px")]},
              ])
        )
        render_any_table(styled)
        return

    st.info("In costruzione: per ora sono complete le tabelle Battuta e Ricezione (giocatori).")

# =========================
# UI: PUNTI PER SET (per ruolo) + fasi
# =========================
def render_points_per_set():
    st.header("Punti per Set")
    st.sidebar.caption("BUILD: PUNTI_PER_SET_V11 (POINTS/SET ALL)")

    # ===== FILTRO RANGE GIORNATE =====
    with engine.begin() as conn:
        bounds = conn.execute(text("""
            SELECT MIN(round_number) AS min_r, MAX(round_number) AS max_r
            FROM matches
            WHERE round_number IS NOT NULL
        """)).mappings().first()

    min_r = int((bounds["min_r"] or 1))
    max_r = int((bounds["max_r"] or 1))

    from_round, to_round, _range_label = get_selected_match_range()
    st.caption(f"📌 Range: {_range_label}")

    if from_round > to_round:
        st.error("Range non valido: 'Da giornata' deve essere <= 'A giornata'.")
        st.stop()

    # ===== FILTRO RUOLI (da roster) =====
    season = st.text_input("Stagione roster", value="2025-26", key="pps_season")
    with engine.begin() as conn:
        roster_rows = conn.execute(text("""
            SELECT team_raw, team_norm, jersey_number, player_name, role, created_at
            FROM roster
            WHERE season = :season
        """), {"season": season}).mappings().all()

    if not roster_rows:
        st.warning("Roster vuoto per questa stagione: importa prima i ruoli (pagina Import Ruoli).")
        return

    df_roster = pd.DataFrame(roster_rows)
    df_roster["created_at"] = df_roster["created_at"].fillna("")
    df_roster = (
        df_roster.sort_values(by=["team_norm", "jersey_number", "created_at"])
                 .drop_duplicates(subset=["team_norm", "jersey_number"], keep="last")
    )

    roles_all = sorted(df_roster["role"].dropna().unique().tolist())
    roles_sel = st.multiselect(
        "Filtri ruoli (puoi selezionarne quanti vuoi)",
        options=roles_all,
        default=roles_all,
        key="pps_roles",
    )

    # ===== FASI (anche entrambe) =====
    cf1, cf2 = st.columns(2)
    with cf1:
        use_sideout = st.checkbox("Fase Side Out", value=True, key="pps_so")
    with cf2:
        use_break = st.checkbox("Fase Break", value=True, key="pps_bp")

    if not use_sideout and not use_break:
        use_sideout = True
        use_break = True

    st.caption("I punti considerati sono solo: **Battuta (#)**, **Attacco (#)**, **Muro (#)**.")

    # ===== MATCHES =====
    with engine.begin() as conn:
        matches = conn.execute(text("""
            SELECT id AS match_id, team_a, team_b, scout_text
            FROM matches
            WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
        """), {"from_round": int(from_round), "to_round": int(to_round)}).mappings().all()

    if not matches:
        st.info("Nessun match nel range selezionato.")
        return

    SET_RE = re.compile(r"\*\*(\d)set\b", re.IGNORECASE)

    def parse_rallies_and_sets(scout_text: str):
        if not scout_text:
            return [], set()

        lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
        current_set = None
        sets_seen = set()
        rallies = []
        current = []
        current_set_for_rally = None

        for ln in lines:
            mm = SET_RE.search(ln)
            if mm:
                if current:
                    rallies.append((current_set_for_rally, current))
                    current = []
                    current_set_for_rally = None
                current_set = int(mm.group(1))
                sets_seen.add(current_set)
                continue

            if current_set is None:
                continue

            c6 = code6(ln)
            if not c6 or c6[0] not in ("*", "a"):
                continue

            if is_serve(c6):
                if current:
                    rallies.append((current_set_for_rally, current))
                current = [c6]
                current_set_for_rally = current_set
                continue

            if not current:
                continue

            current.append(c6)

            if is_home_point(c6) or is_away_point(c6):
                rallies.append((current_set_for_rally, current))
                current = []
                current_set_for_rally = None

        if current:
            rallies.append((current_set_for_rally, current))

        return rallies, sets_seen

    def player_num_from_code(c6: str):
        mm = re.match(r"^[\*a](\d{2})", c6)
        return int(mm.group(1)) if mm else None

    def is_player_action_any(c6: str) -> bool:
        if len(c6) < 5 or c6[0] not in ("*", "a"):
            return False
        if c6[3:5] in ("SQ", "SM", "RQ", "RM"):
            return True
        if len(c6) >= 4 and c6[3] in ("A", "B", "D"):
            return True
        return False

    def point_kind(c6: str):
        if len(c6) < 6 or c6[5] != "#":
            return None
        if c6[3:5] in ("SQ", "SM"):
            return "serve"
        if c6[3] == "A":
            return "attack"
        if c6[3] == "B":
            return "block"
        return None

    def is_error_forced_point(c6: str) -> bool:
        if len(c6) < 6:
            return False
        if c6[3:5] in ("SQ", "SM") and c6[5] == "=":
            return True
        if c6[3] == "A" and c6[5] == "=":
            return True
        if c6[3] == "B" and c6[5] == "/":
            return True
        return False

    players = {}
    team_tot = {}

    def ensure_player(team_raw: str, num: int):
        tnorm = team_norm(team_raw)
        key = (tnorm, num)
        if key not in players:
            players[key] = {
                "team_norm": tnorm,
                "Nome Team": canonical_team(team_raw),
                "N°": num,
                "sets": set(),  # (match_id, set_no)
                "pts_total": 0,
                "pts_serve": 0,
                "pts_attack": 0,
                "pts_block": 0,
            }
        return players[key]

    def ensure_team(team_raw: str):
        tnorm = team_norm(team_raw)
        if tnorm not in team_tot:
            team_tot[tnorm] = {
                "team_norm": tnorm,
                "Nome Team": canonical_team(team_raw),
                "sets_total": 0,
                "pts_total": 0,
                "err_avv_total": 0,
            }
        return team_tot[tnorm]

    for m in matches:
        match_id = int(m.get("match_id") or 0)
        team_a = (m.get("team_a") or "")
        team_b = (m.get("team_b") or "")

        rallies, sets_seen = parse_rallies_and_sets(m.get("scout_text") or "")
        n_sets = len(sets_seen) if sets_seen else 0

        if n_sets:
            ensure_team(team_a)["sets_total"] += n_sets
            ensure_team(team_b)["sets_total"] += n_sets

        for set_no, rally in rallies:
            if not rally or set_no is None:
                continue
            first = rally[0]
            if not is_serve(first):
                continue

            serve_prefix = first[0]

            # set giocati giocatore: >=1 azione nel set (match_id, set_no)
            for c in rally:
                if not is_player_action_any(c):
                    continue
                num = player_num_from_code(c)
                if num is None:
                    continue
                team_raw = team_a if c[0] == "*" else team_b
                ensure_player(team_raw, num)["sets"].add((match_id, int(set_no)))

            # winner
            home_won = any(is_home_point(x) for x in rally)
            away_won = any(is_away_point(x) for x in rally)
            if home_won or away_won:
                if home_won:
                    win_prefix = "*"
                    win_team = team_a
                    lose_prefix = "a"
                else:
                    win_prefix = "a"
                    win_team = team_b
                    lose_prefix = "*"

                is_break_team = (serve_prefix == win_prefix)
                is_sideout_team = (serve_prefix != win_prefix)
                if (is_sideout_team and use_sideout) or (is_break_team and use_break):
                    if any((x[0] == lose_prefix and is_error_forced_point(x)) for x in rally):
                        ensure_team(win_team)["err_avv_total"] += 1

            # punti giocatore
            for c in rally:
                kind = point_kind(c)
                if kind is None:
                    continue
                num = player_num_from_code(c)
                if num is None:
                    continue

                player_prefix = c[0]
                is_break = (serve_prefix == player_prefix)
                is_sideout = (serve_prefix != player_prefix)

                if (is_sideout and not use_sideout) or (is_break and not use_break):
                    continue

                team_raw = team_a if player_prefix == "*" else team_b
                rec = ensure_player(team_raw, num)
                rec["pts_total"] += 1
                if kind == "serve":
                    rec["pts_serve"] += 1
                elif kind == "attack":
                    rec["pts_attack"] += 1
                elif kind == "block":
                    rec["pts_block"] += 1

                ensure_team(team_raw)["pts_total"] += 1

    df_players = pd.DataFrame(list(players.values()))
    if df_players.empty:
        st.info("Nessun dato punti trovato nel range selezionato.")
        return

    df_players = df_players.merge(
        df_roster[["team_norm", "jersey_number", "player_name", "role", "team_raw"]],
        left_on=["team_norm", "N°"],
        right_on=["team_norm", "jersey_number"],
        how="left",
    ).drop(columns=["jersey_number"])

    df_players.rename(columns={"player_name": "Nome giocatore", "role": "Ruolo"}, inplace=True)
    df_players["Nome giocatore"] = df_players["Nome giocatore"].fillna(df_players["N°"].apply(lambda x: f"N°{int(x):02d}"))
    df_players["Ruolo"] = df_players["Ruolo"].fillna("(non in roster)")
    df_players["Nome Team"] = df_players["team_raw"].apply(canonical_team).fillna(df_players["Nome Team"])
    df_players = df_players.drop(columns=["team_raw"])

    # Canonicalizza sempre il nome squadra (solo città)
    df_players["Nome Team"] = df_players["Nome Team"].apply(canonical_team)


    if roles_sel:
        df_players = df_players[df_players["Ruolo"].isin(roles_sel)].copy()

    df_players["Set giocati dal Giocatore"] = df_players["sets"].apply(lambda s: len(s) if isinstance(s, set) else 0)

    df_team = pd.DataFrame(list(team_tot.values()))
    df_team["Punti per Set (Team)"] = df_team.apply(lambda r: (r["pts_total"] / r["sets_total"]) if r["sets_total"] else 0.0, axis=1)
    df_team["Errori Avv/Set"] = df_team.apply(lambda r: (r["err_avv_total"] / r["sets_total"]) if r["sets_total"] else 0.0, axis=1)

    df_players = df_players.merge(
        df_team[["team_norm", "sets_total", "Punti per Set (Team)", "Errori Avv/Set"]],
        on="team_norm",
        how="left",
    )

    df_players["Punti per Set (Giocatore)"] = df_players.apply(
        lambda r: (r["pts_total"] / r["Set giocati dal Giocatore"]) if r["Set giocati dal Giocatore"] else 0.0,
        axis=1
    )

    df_players["% Ply/Team"] = df_players.apply(
        lambda r: (100.0 * r["Punti per Set (Giocatore)"] / r["Punti per Set (Team)"]) if (r["Punti per Set (Team)"] and r["Punti per Set (Team)"] > 0) else 0.0,
        axis=1
    )

    df_rank = df_players.sort_values(
        by=["Punti per Set (Giocatore)", "Set giocati dal Giocatore"],
        ascending=[False, False]
    ).reset_index(drop=True)
    df_rank.insert(0, "Ranking", range(1, len(df_rank) + 1))

    out = df_rank[[
        "Ranking",
        "Nome Team",
        "sets_total",
        "Punti per Set (Team)",
        "Errori Avv/Set",
        "Nome giocatore",
        "Set giocati dal Giocatore",
        "Punti per Set (Giocatore)",
        "% Ply/Team",
        "pts_serve",
        "pts_attack",
        "pts_block",
    ]].rename(columns={
        "sets_total": "Set giocati dal Team",
        "pts_serve": "Punti in Battuta",
        "pts_attack": "Punti in Attacco",
        "pts_block": "Punti a Muro",
    }).copy()

    # --- Trasforma Punti in Battuta/Attacco/Muro in valori PER SET del giocatore ---
    denom = pd.to_numeric(out["Set giocati dal Giocatore"], errors="coerce").replace(0, pd.NA)
    for col in ["Punti in Battuta", "Punti in Attacco", "Punti a Muro"]:
        out[col] = (pd.to_numeric(out[col], errors="coerce") / denom).astype(float).fillna(0.0)

    # formatter 1 decimale senza zeri finali
    def _fmt1(x):
        try:
            if x is None:
                return ""
            v = float(x)
            s = f"{v:.1f}"
            return s.rstrip("0").rstrip(".")
        except Exception:
            return x

    def highlight_perugia(row):
        is_perugia = "perugia" in str(row.get("Nome Team", "")).lower()
        style = "background-color: #fff3cd; font-weight: 800;" if is_perugia else ""
        return [style] * len(row)

    styled = (
        out.style
          .apply(highlight_perugia, axis=1)
          .format({
              "Punti per Set (Team)": _fmt1,
              "Errori Avv/Set": _fmt1,
              "Punti per Set (Giocatore)": _fmt1,
              "% Ply/Team": _fmt1,
              "Punti in Battuta": _fmt1,
              "Punti in Attacco": _fmt1,
              "Punti a Muro": _fmt1,
          })
    )
    render_focus_4_players(out, key_base=f"pps_{from_round}_{to_round}")


    render_any_table(styled)


# =========================
# UI: HOME DASHBOARD (3 TAB)
# =========================
def render_home_dashboard():
    st.header("Home – Dashboard")

    # ===== RANGE GIORNATE =====
    with engine.begin() as conn:
        bounds = conn.execute(text("""
            SELECT MIN(round_number) AS min_r, MAX(round_number) AS max_r
            FROM matches
            WHERE round_number IS NOT NULL
        """)).mappings().first()

    min_r = int((bounds["min_r"] or 1))
    max_r = int((bounds["max_r"] or 1))

    from_round, to_round, _range_label = get_selected_match_range()
    st.caption(f"📌 Range: {_range_label}")

    if from_round > to_round:
        st.error("Range non valido: 'Da giornata' deve essere <= 'A giornata'.")
        st.stop()

    # ===== SQUADRA (default Perugia) =====
    with engine.begin() as conn:
        teams = conn.execute(text("""
            SELECT DISTINCT team_a AS t FROM matches
            UNION
            SELECT DISTINCT team_b AS t FROM matches
        """)).mappings().all()
    team_list = sorted({canonical_team(r["t"] or "") for r in teams if (r.get("t") or "").strip()})
    # default: PERUGIA se presente
    default_team = "PERUGIA" if "PERUGIA" in team_list else (team_list[0] if team_list else "")

    team_focus = st.selectbox("Squadra", team_list, index=(team_list.index(default_team) if default_team in team_list else 0), key="home_team")

    # ===== DATA (matches in range) =====
    with engine.begin() as conn:
        matches = conn.execute(text("""
            SELECT id AS match_id, team_a, team_b, scout_text
            FROM matches
            WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
        """), {"from_round": int(from_round), "to_round": int(to_round)}).mappings().all()

    if not matches:
        st.info("Nessun match nel range selezionato.")
        return

    SET_RE = re.compile(r"\*\*(\d)set\b", re.IGNORECASE)

    def parse_rallies(scout_text: str):
        if not scout_text:
            return []
        lines = [ln.strip() for ln in str(scout_text).splitlines() if ln and ln.strip()]
        current_set = None
        rallies = []
        current = []

        for ln in lines:
            m = SET_RE.search(ln)
            if m:
                current_set = int(m.group(1))
                continue

            if current_set is None:
                continue

            c6 = code6(ln)
            if not c6 or c6[0] not in ("*", "a"):
                continue

            if is_serve(c6):
                if current:
                    rallies.append(current)
                current = [c6]
                continue

            if not current:
                continue
            current.append(c6)

            if is_home_point(c6) or is_away_point(c6):
                rallies.append(current)
                current = []

        if current:
            rallies.append(current)
        return rallies

    def serve_sign(c6: str) -> str:
        return c6[5] if c6 and len(c6) >= 6 else ""

    def serve_type(c6: str) -> str:
        return c6[3:5] if c6 and len(c6) >= 5 else ""

    def is_rece(c6: str) -> bool:
        return len(c6) >= 6 and c6[3:5] in ("RQ", "RM") and c6[0] in ("*", "a")

    def rece_sign(c6: str) -> str:
        return c6[5] if c6 and len(c6) >= 6 else ""

    def is_attack_code(c6: str) -> bool:
        return len(c6) >= 6 and c6[0] in ("*", "a") and c6[3] == "A"

    def is_block_code(c6: str) -> bool:
        return len(c6) >= 6 and c6[0] in ("*", "a") and c6[3] == "B"

    def is_def_code(c6: str) -> bool:
        return len(c6) >= 6 and c6[0] in ("*", "a") and c6[3] == "D"

    def team_of(prefix: str, team_a: str, team_b: str) -> str:
        return team_a if prefix == "*" else team_b

    # ===== Aggregazioni per team =====
    T = {}

    def ensure_team(team_raw: str):
        key_team = canonical_team(team_raw)
        if key_team not in T:
            T[key_team] = {
                "team_norm": key_team,
                "Team": key_team,
                "sets_total": 0,

                "so_att": 0, "so_win": 0,
                "so_spin_att": 0, "so_spin_win": 0,
                "so_float_att": 0, "so_float_win": 0,
                "so_dir_win": 0,
                "so_play_att": 0, "so_play_win": 0,
                "so_good_att": 0, "so_good_win": 0,
                "so_exc_att": 0, "so_exc_win": 0,
                "so_neg_att": 0, "so_neg_win": 0,

                "bp_att": 0, "bp_win": 0,
                "bp_play_att": 0, "bp_play_win": 0,
                "bp_neg_att": 0, "bp_neg_win": 0,
                "bp_exc_att": 0, "bp_exc_win": 0,
                "bp_pos_att": 0, "bp_pos_win": 0,
                "bp_half_att": 0, "bp_half_win": 0,
                "bt_ace": 0, "bt_err": 0,

                "srv_tot": 0, "srv_hash": 0, "srv_half": 0, "srv_pos": 0, "srv_exc": 0, "srv_neg": 0, "srv_err": 0,

                "rec_tot": 0, "rec_hash": 0, "rec_pos": 0, "rec_exc": 0, "rec_neg": 0, "rec_half": 0, "rec_err": 0,

                "att_tot": 0, "att_hash": 0, "att_pos": 0, "att_exc": 0, "att_neg": 0, "att_blk": 0, "att_err": 0,
                "att_first_tot": 0, "att_first_hash": 0, "att_first_blk": 0, "att_first_err": 0,
                "att_tr_tot": 0, "att_tr_hash": 0, "att_tr_blk": 0, "att_tr_err": 0,

                "blk_tot": 0, "blk_hash": 0, "blk_pos": 0, "blk_neg": 0, "blk_cov": 0, "blk_inv": 0, "blk_err": 0,

                "def_tot": 0, "def_pos": 0, "def_cov": 0, "def_neg": 0, "def_over": 0, "def_err": 0,
            }
        return T[key_team]

    # Sets per match
    for m in matches:
        ta_raw = (m.get("team_a") or "")
        tb_raw = (m.get("team_b") or "")
        ta = ta_raw
        tb = tb_raw
        scout_text = m.get("scout_text") or ""
        sets_seen = set(int(x) for x in SET_RE.findall(scout_text))
        n_sets = len(sets_seen)
        if n_sets:
            ensure_team(ta)["sets_total"] += n_sets
            ensure_team(tb)["sets_total"] += n_sets

    for m in matches:
        ta_raw = (m.get("team_a") or "")
        tb_raw = (m.get("team_b") or "")
        ta = ta_raw
        tb = tb_raw
        rallies = parse_rallies(m.get("scout_text") or "")
        for r in rallies:
            if not r or not is_serve(r[0]):
                continue
            first = r[0]
            s_prefix = first[0]
            s_team = team_of(s_prefix, ta, tb)
            rcv_prefix = "a" if s_prefix == "*" else "*"
            rcv_team = team_of(rcv_prefix, ta, tb)
            sgn = serve_sign(first)
            stype = serve_type(first)

            home_won = any(is_home_point(x) for x in r)
            away_won = any(is_away_point(x) for x in r)
            s_team_won = (home_won and s_prefix == "*") or (away_won and s_prefix == "a")
            rcv_team_won = (home_won and rcv_prefix == "*") or (away_won and rcv_prefix == "a")

            # SideOut (ricevente)
            if any(is_rece(x) and x[0] == rcv_prefix for x in r):
                t = ensure_team(rcv_team)
                t["so_att"] += 1
                if rcv_team_won:
                    t["so_win"] += 1

            if stype == "SQ":
                t = ensure_team(rcv_team)
                t["so_spin_att"] += 1
                if rcv_team_won:
                    t["so_spin_win"] += 1
            if stype == "SM":
                t = ensure_team(rcv_team)
                t["so_float_att"] += 1
                if rcv_team_won:
                    t["so_float_win"] += 1

            rece = None
            for x in r:
                if is_rece(x) and x[0] == rcv_prefix:
                    rece = x
                    break

            if rece:
                rs = rece_sign(rece)
                t = ensure_team(rcv_team)
                if rs in ("#", "+", "!", "-"):
                    t["so_play_att"] += 1
                    if rcv_team_won:
                        t["so_play_win"] += 1
                if rs in ("#", "+"):
                    t["so_good_att"] += 1
                    if rcv_team_won:
                        t["so_good_win"] += 1
                if rs == "!":
                    t["so_exc_att"] += 1
                    if rcv_team_won:
                        t["so_exc_win"] += 1
                if rs == "-":
                    t["so_neg_att"] += 1
                    if rcv_team_won:
                        t["so_neg_win"] += 1

            first_att = None
            for x in r:
                if is_attack_code(x) and x[0] == rcv_prefix:
                    first_att = x
                    break
            if first_att and len(first_att) >= 6 and first_att[5] == "#" and rcv_team_won:
                ensure_team(rcv_team)["so_dir_win"] += 1

            # Break (battitore)
            bt = ensure_team(s_team)
            bt["bp_att"] += 1
            if s_team_won:
                bt["bp_win"] += 1

            if sgn not in ("#", "="):
                bt["bp_play_att"] += 1
                if s_team_won:
                    bt["bp_play_win"] += 1

            if sgn == "-":
                bt["bp_neg_att"] += 1
                if s_team_won:
                    bt["bp_neg_win"] += 1
            if sgn == "!":
                bt["bp_exc_att"] += 1
                if s_team_won:
                    bt["bp_exc_win"] += 1
            if sgn == "+":
                bt["bp_pos_att"] += 1
                if s_team_won:
                    bt["bp_pos_win"] += 1
            if sgn == "/":
                bt["bp_half_att"] += 1
                if s_team_won:
                    bt["bp_half_win"] += 1

            if sgn == "#":
                bt["bt_ace"] += 1
            if sgn == "=":
                bt["bt_err"] += 1

            # Serve distribution
            bt["srv_tot"] += 1
            if sgn == "#":
                bt["srv_hash"] += 1
            elif sgn == "/":
                bt["srv_half"] += 1
            elif sgn == "+":
                bt["srv_pos"] += 1
            elif sgn == "!":
                bt["srv_exc"] += 1
            elif sgn == "-":
                bt["srv_neg"] += 1
            elif sgn == "=":
                bt["srv_err"] += 1

            # Reception distribution
            if rece:
                rt = ensure_team(rcv_team)
                rt["rec_tot"] += 1
                rs = rece_sign(rece)
                if rs == "#":
                    rt["rec_hash"] += 1
                elif rs == "+":
                    rt["rec_pos"] += 1
                elif rs == "!":
                    rt["rec_exc"] += 1
                elif rs == "-":
                    rt["rec_neg"] += 1
                elif rs == "/":
                    rt["rec_half"] += 1
                elif rs == "=":
                    rt["rec_err"] += 1

            # Attack + first vs transition
            first_attack_index = None
            for i, x in enumerate(r[1:], start=1):
                if is_attack_code(x):
                    first_attack_index = i
                    break

            for i, x in enumerate(r[1:], start=1):
                if not is_attack_code(x):
                    continue
                at = ensure_team(team_of(x[0], ta, tb))
                at["att_tot"] += 1
                sign = x[5]
                if sign == "#":
                    at["att_hash"] += 1
                elif sign == "+":
                    at["att_pos"] += 1
                elif sign == "!":
                    at["att_exc"] += 1
                elif sign == "-":
                    at["att_neg"] += 1
                elif sign == "/":
                    at["att_blk"] += 1
                elif sign == "=":
                    at["att_err"] += 1

                if first_attack_index is not None:
                    if i == first_attack_index:
                        at["att_first_tot"] += 1
                        if sign == "#":
                            at["att_first_hash"] += 1
                        if sign == "/":
                            at["att_first_blk"] += 1
                        if sign == "=":
                            at["att_first_err"] += 1
                    elif i > first_attack_index:
                        at["att_tr_tot"] += 1
                        if sign == "#":
                            at["att_tr_hash"] += 1
                        if sign == "/":
                            at["att_tr_blk"] += 1
                        if sign == "=":
                            at["att_tr_err"] += 1

            # Block
            for x in r[1:]:
                if not is_block_code(x):
                    continue
                bl = ensure_team(team_of(x[0], ta, tb))
                bl["blk_tot"] += 1
                sign = x[5]
                if sign == "#":
                    bl["blk_hash"] += 1
                elif sign == "+":
                    bl["blk_pos"] += 1
                elif sign == "-":
                    bl["blk_neg"] += 1
                elif sign == "!":
                    bl["blk_cov"] += 1
                elif sign == "/":
                    bl["blk_inv"] += 1
                elif sign == "=":
                    bl["blk_err"] += 1

            # Defense
            for x in r[1:]:
                if not is_def_code(x):
                    continue
                df = ensure_team(team_of(x[0], ta, tb))
                df["def_tot"] += 1
                sign = x[5]
                if sign == "+":
                    df["def_pos"] += 1
                elif sign == "!":
                    df["def_cov"] += 1
                elif sign == "-":
                    df["def_neg"] += 1
                elif sign == "/":
                    df["def_over"] += 1
                elif sign == "=":
                    df["def_err"] += 1

    dfT = pd.DataFrame(list(T.values()))
    if dfT.empty:
        st.info("Nessun dato nel range selezionato.")
        return

    # ===== DF team da DB (come tabelle nelle sezioni) per SIDE OUT + BREAK =====
    # Questo evita discrepanze: la Home usa gli stessi numeratori/denominatori importati nel DB.
    with engine.begin() as conn:
        rows_db = conn.execute(text("""
            SELECT
                team_a, team_b,
                COALESCE(so_home_attempts,0) AS so_home_attempts,
                COALESCE(so_home_wins,0)     AS so_home_wins,
                COALESCE(so_away_attempts,0) AS so_away_attempts,
                COALESCE(so_away_wins,0)     AS so_away_wins,

                COALESCE(so_spin_home_attempts,0) AS so_spin_home_attempts,
                COALESCE(so_spin_home_wins,0)     AS so_spin_home_wins,
                COALESCE(so_spin_away_attempts,0) AS so_spin_away_attempts,
                COALESCE(so_spin_away_wins,0)     AS so_spin_away_wins,

                COALESCE(so_float_home_attempts,0) AS so_float_home_attempts,
                COALESCE(so_float_home_wins,0)     AS so_float_home_wins,
                COALESCE(so_float_away_attempts,0) AS so_float_away_attempts,
                COALESCE(so_float_away_wins,0)     AS so_float_away_wins,

                COALESCE(so_dir_home_wins,0) AS so_dir_home_wins,
                COALESCE(so_dir_away_wins,0) AS so_dir_away_wins,

                COALESCE(so_play_home_attempts,0) AS so_play_home_attempts,
                COALESCE(so_play_home_wins,0)     AS so_play_home_wins,
                COALESCE(so_play_away_attempts,0) AS so_play_away_attempts,
                COALESCE(so_play_away_wins,0)     AS so_play_away_wins,

                COALESCE(so_good_home_attempts,0) AS so_good_home_attempts,
                COALESCE(so_good_home_wins,0)     AS so_good_home_wins,
                COALESCE(so_good_away_attempts,0) AS so_good_away_attempts,
                COALESCE(so_good_away_wins,0)     AS so_good_away_wins,

                COALESCE(so_exc_home_attempts,0) AS so_exc_home_attempts,
                COALESCE(so_exc_home_wins,0)     AS so_exc_home_wins,
                COALESCE(so_exc_away_attempts,0) AS so_exc_away_attempts,
                COALESCE(so_exc_away_wins,0)     AS so_exc_away_wins,

                COALESCE(so_neg_home_attempts,0) AS so_neg_home_attempts,
                COALESCE(so_neg_home_wins,0)     AS so_neg_home_wins,
                COALESCE(so_neg_away_attempts,0) AS so_neg_away_attempts,
                COALESCE(so_neg_away_wins,0)     AS so_neg_away_wins,

                COALESCE(bp_home_attempts,0) AS bp_home_attempts,
                COALESCE(bp_home_wins,0)     AS bp_home_wins,
                COALESCE(bp_away_attempts,0) AS bp_away_attempts,
                COALESCE(bp_away_wins,0)     AS bp_away_wins,

                COALESCE(bp_play_home_attempts,0) AS bp_play_home_attempts,
                COALESCE(bp_play_home_wins,0)     AS bp_play_home_wins,
                COALESCE(bp_play_away_attempts,0) AS bp_play_away_attempts,
                COALESCE(bp_play_away_wins,0)     AS bp_play_away_wins
            FROM matches
            WHERE ( (CASE matches.phase WHEN 'A' THEN 1 WHEN 'R' THEN 2 WHEN 'POQ' THEN 3 WHEN 'POS' THEN 4 WHEN 'POF' THEN 5 ELSE 99 END) * 100 + COALESCE(matches.round_number,0) ) BETWEEN :from_round AND :to_round
        """), {"from_round": int(from_round), "to_round": int(to_round)}).mappings().all()

    agg_db = {}
    def _ensure_db(team_raw: str):
        t = canonical_team(team_raw)
        if t not in agg_db:
            agg_db[t] = {
                "Team": t,
                "so_att": 0, "so_win": 0,
                "so_spin_att": 0, "so_spin_win": 0,
                "so_float_att": 0, "so_float_win": 0,
                "so_dir_win": 0,
                "so_play_att": 0, "so_play_win": 0,
                "so_good_att": 0, "so_good_win": 0,
                "so_exc_att": 0, "so_exc_win": 0,
                "so_neg_att": 0, "so_neg_win": 0,
                "bp_att": 0, "bp_win": 0,
                "bp_play_att": 0, "bp_play_win": 0,
            }
        return agg_db[t]

    for r in rows_db:
        ha = _ensure_db(r.get('team_a') or '')
        hb = _ensure_db(r.get('team_b') or '')
        ha["so_att"] += int(r.get('so_home_attempts') or 0)
        ha["so_win"] += int(r.get('so_home_wins') or 0)
        hb["so_att"] += int(r.get('so_away_attempts') or 0)
        hb["so_win"] += int(r.get('so_away_wins') or 0)

        ha["so_spin_att"] += int(r.get('so_spin_home_attempts') or 0)
        ha["so_spin_win"] += int(r.get('so_spin_home_wins') or 0)
        hb["so_spin_att"] += int(r.get('so_spin_away_attempts') or 0)
        hb["so_spin_win"] += int(r.get('so_spin_away_wins') or 0)

        ha["so_float_att"] += int(r.get('so_float_home_attempts') or 0)
        ha["so_float_win"] += int(r.get('so_float_home_wins') or 0)
        hb["so_float_att"] += int(r.get('so_float_away_attempts') or 0)
        hb["so_float_win"] += int(r.get('so_float_away_wins') or 0)

        ha["so_dir_win"] += int(r.get('so_dir_home_wins') or 0)
        hb["so_dir_win"] += int(r.get('so_dir_away_wins') or 0)

        ha["so_play_att"] += int(r.get('so_play_home_attempts') or 0)
        ha["so_play_win"] += int(r.get('so_play_home_wins') or 0)
        hb["so_play_att"] += int(r.get('so_play_away_attempts') or 0)
        hb["so_play_win"] += int(r.get('so_play_away_wins') or 0)

        ha["so_good_att"] += int(r.get('so_good_home_attempts') or 0)
        ha["so_good_win"] += int(r.get('so_good_home_wins') or 0)
        hb["so_good_att"] += int(r.get('so_good_away_attempts') or 0)
        hb["so_good_win"] += int(r.get('so_good_away_wins') or 0)

        ha["so_exc_att"] += int(r.get('so_exc_home_attempts') or 0)
        ha["so_exc_win"] += int(r.get('so_exc_home_wins') or 0)
        hb["so_exc_att"] += int(r.get('so_exc_away_attempts') or 0)
        hb["so_exc_win"] += int(r.get('so_exc_away_wins') or 0)

        ha["so_neg_att"] += int(r.get('so_neg_home_attempts') or 0)
        ha["so_neg_win"] += int(r.get('so_neg_home_wins') or 0)
        hb["so_neg_att"] += int(r.get('so_neg_away_attempts') or 0)
        hb["so_neg_win"] += int(r.get('so_neg_away_wins') or 0)

        ha["bp_att"] += int(r.get('bp_home_attempts') or 0)
        ha["bp_win"] += int(r.get('bp_home_wins') or 0)
        hb["bp_att"] += int(r.get('bp_away_attempts') or 0)
        hb["bp_win"] += int(r.get('bp_away_wins') or 0)

        ha["bp_play_att"] += int(r.get('bp_play_home_attempts') or 0)
        ha["bp_play_win"] += int(r.get('bp_play_home_wins') or 0)
        hb["bp_play_att"] += int(r.get('bp_play_away_attempts') or 0)
        hb["bp_play_win"] += int(r.get('bp_play_away_wins') or 0)

    dfT_SO_BP = pd.DataFrame(list(agg_db.values()))

    def _pct(num, den):
        return (100.0 * num / den) if den else 0.0

    def build_df(value_series: pd.Series, higher_is_better: bool = True):
        df = pd.DataFrame({"Team": dfT["Team"], "Value": value_series})
        df["Value"] = pd.to_numeric(df["Value"], errors="coerce").fillna(0.0)
        df = df.sort_values(by="Value", ascending=not higher_is_better).reset_index(drop=True)
        df["Rank"] = range(1, len(df) + 1)
        df.attrs["higher_is_better"] = higher_is_better
        return df

    def build_df_db(value_series: pd.Series, higher_is_better: bool = True):
        # usa dfT_SO_BP (coerente con le tabelle nelle sezioni)
        df = pd.DataFrame({"Team": dfT_SO_BP["Team"], "Value": value_series})
        df["Value"] = pd.to_numeric(df["Value"], errors="coerce").fillna(0.0)
        df = df.sort_values(by="Value", ascending=not higher_is_better).reset_index(drop=True)
        df["Rank"] = range(1, len(df) + 1)
        df.attrs["higher_is_better"] = higher_is_better
        return df

    def render_card(title: str, df: pd.DataFrame, fmt: str = "{:.1f}", higher_is_better: bool | None = None):
        # higher_is_better: se None, prova a leggere da df.attrs
        hib = higher_is_better
        if hib is None:
            hib = bool(df.attrs.get("higher_is_better", True))

        # trova riga team
        team_row = None
        for _, r in df.iterrows():
            if norm(r["Team"]) == norm(team_focus):
                team_row = r
                break
            if "perugia" in norm(team_focus) and "perugia" in norm(r["Team"]):
                team_row = r
                break
        if team_row is None:
            team_row = df.iloc[0] if not df.empty else None

        val = float(team_row["Value"]) if team_row is not None else 0.0
        rk = int(team_row["Rank"]) if team_row is not None else 0

        # delta dal 3° posto
        third_val = float(df.iloc[2]["Value"]) if len(df) >= 3 else float(df.iloc[-1]["Value"])
        # per metrica "più alto è meglio": delta = val - third
        # per "più basso è meglio": delta = third - val (positivo = sei davanti, negativo = dietro)
        delta = (val - third_val) if hib else (third_val - val)

        # colore in base al rank
        if rk and rk <= 3:
            color = "#2f9e44"   # green
        elif rk and rk <= 6:
            color = "#f08c00"   # orange
        else:
            color = "#495057"   # gray

        st.markdown(f"**{title}**")

        # valore + rank + delta (stile 'da panchina')
        delta_txt = f"{delta:+.1f} vs 3°"
        # se 'lower is better' e delta==inf (es. Err/Pti), gestisci
        if not (delta == delta and abs(delta) != float("inf")):
            delta_txt = "—"

        st.markdown(
            f"""
            <div style="display:flex; align-items:baseline; justify-content:space-between; gap:12px; padding:8px 10px; border:1px solid #e9ecef; border-radius:12px;">
              <div>
                <div style="font-size:34px; font-weight:900; color:{color}; line-height:1;">{fmt.format(val)}</div>
                <div style="font-size:14px; color:#868e96; margin-top:2px;">{team_focus} • Rank {rk}/12</div>
              </div>
              <div style="text-align:right;">
                <div style="font-size:14px; color:#868e96;">gap podio</div>
                <div style="font-size:18px; font-weight:800; color:#343a40;">{delta_txt}</div>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        top3 = df.head(3).copy()
        top3["Team"] = top3["Team"].apply(canonical_team)
        top3["Value"] = top3["Value"].astype(float).round(1)

        def _hl_perugia_top3(row):
            is_per = "perugia" in str(row["Team"]).lower()
            style = "background-color: #fff3cd; font-weight: 800;" if is_per else ""
            return [style] * len(row)

        st.dataframe(
            top3[["Rank", "Team", "Value"]]
                .style
                .apply(_hl_perugia_top3, axis=1)
                .format({"Rank": "{:.0f}", "Value": "{:.1f}"}),
            width="stretch",
            hide_index=True,
        )

    tab_so, tab_bp, tab_eff = st.tabs(["SIDE OUT", "BREAK", "EFFICIENZE"])

    with tab_so:
        cols = st.columns(3)
        metrics = [
            ("Side Out TOTALE", build_df_db(dfT_SO_BP.apply(lambda r: _pct(r["so_win"], r["so_att"]), axis=1))),
            ("Side Out SPIN", build_df_db(dfT_SO_BP.apply(lambda r: _pct(r["so_spin_win"], r["so_spin_att"]), axis=1))),
            ("Side Out FLOAT", build_df_db(dfT_SO_BP.apply(lambda r: _pct(r["so_float_win"], r["so_float_att"]), axis=1))),
            ("Side Out DIRETTO", build_df_db(dfT_SO_BP.apply(lambda r: _pct(r["so_dir_win"], r["so_att"]), axis=1))),
            ("Side Out GIOCATO", build_df_db(dfT_SO_BP.apply(lambda r: _pct(r["so_play_win"], r["so_play_att"]), axis=1))),
            ("Side Out con RICE BUONA", build_df_db(dfT_SO_BP.apply(lambda r: _pct(r["so_good_win"], r["so_good_att"]), axis=1))),
            ("Side Out con RICE ESCLAMATIVA", build_df_db(dfT_SO_BP.apply(lambda r: _pct(r["so_exc_win"], r["so_exc_att"]), axis=1))),
            ("Side Out con RICE NEGATIVA", build_df_db(dfT_SO_BP.apply(lambda r: _pct(r["so_neg_win"], r["so_neg_att"]), axis=1))),
        ]
        for i, (title, dfm) in enumerate(metrics):
            with cols[i % 3]:
                render_card(title, dfm)

    with tab_bp:
        cols = st.columns(3)

        def ratio_err_ace(r):
            ace = float(r.get("bt_ace", 0) or 0)
            err = float(r.get("bt_err", 0) or 0)
            return (err / ace) if ace else float("inf")

        metrics = [
            ("BREAK TOTALE", build_df_db(dfT_SO_BP.apply(lambda r: _pct(r["bp_win"], r["bp_att"]), axis=1))),
            ("BREAK GIOCATO", build_df_db(dfT_SO_BP.apply(lambda r: _pct(r["bp_play_win"], r["bp_play_att"]), axis=1))),
            ("BREAK con BT. NEGATIVA", build_df(dfT.apply(lambda r: _pct(r["bp_neg_win"], r["bp_neg_att"]), axis=1))),
            ("BREAK con BT. ESCLAMATIVA", build_df(dfT.apply(lambda r: _pct(r["bp_exc_win"], r["bp_exc_att"]), axis=1))),
            ("BREAK con BT. POSITIVA", build_df(dfT.apply(lambda r: _pct(r["bp_pos_win"], r["bp_pos_att"]), axis=1))),
            ("BREAK con BT. 1/2 PUNTO", build_df(dfT.apply(lambda r: _pct(r["bp_half_win"], r["bp_half_att"]), axis=1))),
            ("BT punto/errore/ratio (Err/Pti)", build_df(dfT.apply(lambda r: ratio_err_ace(r), axis=1), higher_is_better=False)),
        ]

        for i, (title, dfm) in enumerate(metrics):
            with cols[i % 3]:
                if "Err/Pti" in title:
                    render_card(title, dfm, fmt="{:.2f}", higher_is_better=False)
                else:
                    render_card(title, dfm)

    with tab_eff:
        st.subheader("Filtri Efficienze")
        cflag = st.columns(3)
        with cflag[0]:
            st.checkbox("Battuta SPIN", value=True, key="home_srv_spin")
            st.checkbox("Battuta FLOAT", value=True, key="home_srv_float")
        with cflag[1]:
            st.checkbox("Ricezione SPIN", value=True, key="home_rec_spin")
            st.checkbox("Ricezione FLOAT", value=True, key="home_rec_float")
        with cflag[2]:
            att_first = st.checkbox("Attacco dopo Ricezione", value=True, key="home_att_first")
            att_tr = st.checkbox("Attacco di Transizione", value=True, key="home_att_tr")

        if not att_first and not att_tr:
            att_first = att_tr = True

        def eff_serve(r):
            tot = r["srv_tot"]
            if not tot:
                return 0.0
            return ((r["srv_hash"] + r["srv_half"]*0.8 + r["srv_pos"]*0.45 + r["srv_exc"]*0.3 + r["srv_neg"]*0.15 - r["srv_err"]) / tot) * 100.0

        def eff_rece(r):
            tot = r["rec_tot"]
            if not tot:
                return 0.0
            ok = r["rec_hash"] + r["rec_pos"]
            return ((ok*0.77 + r["rec_exc"]*0.55 + r["rec_neg"]*0.38 - r["rec_half"]*0.8 - r["rec_err"]) / tot) * 100.0

        def eff_att_total(r):
            tot = r["att_tot"]
            if not tot:
                return 0.0
            ko = r["att_blk"] + r["att_err"]
            return ((r["att_hash"] - ko) / tot) * 100.0

        def eff_att_first(r):
            tot = r["att_first_tot"]
            if not tot:
                return 0.0
            ko = r["att_first_blk"] + r["att_first_err"]
            return ((r["att_first_hash"] - ko) / tot) * 100.0

        def eff_att_tr(r):
            tot = r["att_tr_tot"]
            if not tot:
                return 0.0
            ko = r["att_tr_blk"] + r["att_tr_err"]
            return ((r["att_tr_hash"] - ko) / tot) * 100.0

        def eff_block(r):
            tot = r["blk_tot"]
            if not tot:
                return 0.0
            return ((r["blk_hash"]*2 + r["blk_pos"]*0.7 + r["blk_neg"]*0.07 + r["blk_cov"]*0.15 - r["blk_inv"] - r["blk_err"]) / tot) * 100.0

        def eff_def(r):
            tot = r["def_tot"]
            if not tot:
                return 0.0
            return ((r["def_pos"]*2 + r["def_cov"]*0.5 + r["def_neg"]*0.4 + r["def_over"]*0.3 - r["def_err"]) / tot) * 100.0

        cols = st.columns(3)
        eff_cards = [
            ("Efficienza BATTUTA", build_df(dfT.apply(lambda r: eff_serve(r), axis=1))),
            ("Efficienza RICEZIONE", build_df(dfT.apply(lambda r: eff_rece(r), axis=1))),
        ]

        if att_first and att_tr:
            eff_cards.append(("Efficienza ATTACCO (Tot)", build_df(dfT.apply(lambda r: eff_att_total(r), axis=1))))
        elif att_first:
            eff_cards.append(("Efficienza ATTACCO (Dopo Ricezione)", build_df(dfT.apply(lambda r: eff_att_first(r), axis=1))))
        else:
            eff_cards.append(("Efficienza ATTACCO (Transizione)", build_df(dfT.apply(lambda r: eff_att_tr(r), axis=1))))

        eff_cards += [
            ("Efficienza MURO TOTALE", build_df(dfT.apply(lambda r: eff_block(r), axis=1))),
            ("Efficienza DIFESA TOTALE", build_df(dfT.apply(lambda r: eff_def(r), axis=1))),
        ]

        for i, (title, dfm) in enumerate(eff_cards):
            with cols[i % 3]:
                render_card(title, dfm, fmt="{:.1f}")


# =========================
# MAIN
# =========================
init_db()

st.sidebar.title("Volley App")

# Filtri Partite (Andata/Ritorno/Playoff)
_mf_from, _mf_to, _mf_label = sidebar_match_filters()
page = st.sidebar.radio(
    "Vai a:",
    [
        "Home",
        "Import DVW (solo staff)",
        "Import Ruoli (solo staff)",
        "Indici Side Out - Squadre",
        "Indici Fase Break - Squadre",
        "GRAFICI 4 Quadranti",
        "Indici Side Out - Giocatori (per ruolo)",
        "Indici Break Point - Giocatori (per ruolo)",
        "Classifiche Fondamentali - Squadre",
        "Classifiche Fondamentali - Giocatori (per ruolo)",
        "Punti per Set",
    ],
    key="nav_page",
)

ADMIN_MODE = st.sidebar.checkbox("Modalità staff (admin)", value=True, key="admin_mode")

if page == "Home":
    render_home_dashboard()

elif page == "Import DVW (solo staff)":
    if not staff_unlocked():
        st.error("Area riservata.")
        st.stop()
    render_import(ADMIN_MODE)
elif page == "Import Ruoli (solo staff)":
    if not staff_unlocked():
        st.error("Area riservata.")
        st.stop()
    render_import_ruoli(ADMIN_MODE)
elif page == "Indici Side Out - Squadre":
    render_sideout_team()

elif page == "Indici Fase Break - Squadre":
    render_break_team()

elif page == "GRAFICI 4 Quadranti":
    render_grafici_4_quadranti()


elif page == "Indici Side Out - Giocatori (per ruolo)":
    render_sideout_players_by_role()

elif page == "Indici Break Point - Giocatori (per ruolo)":
    render_break_players_by_role()


elif page == "Classifiche Fondamentali - Squadre":
    render_fondamentali_team()

elif page == "Classifiche Fondamentali - Giocatori (per ruolo)":
    render_fondamentali_players()

elif page == "Punti per Set":
    render_points_per_set()

else:
    st.header(page)
    st.info("In costruzione.")
