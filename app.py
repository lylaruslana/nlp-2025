import re
import time
import random
from pathlib import Path
import pandas as pd
import streamlit as st
import plotly.express as px

# -----------------------------------------------------------------------------
# KONFIGURASI HALAMAN
# -----------------------------------------------------------------------------
st.set_page_config(page_title="Laporan Capaian NLP 2025", layout="wide")
DEFAULT_FILE = Path("Form Capaian PRSDI 2025.xlsx")

# -----------------------------------------------------------------------------
# CSS INJECTION (FRAME BIRU MUDA #87CEFA)
# -----------------------------------------------------------------------------
st.markdown("""
    <style>
    /* Membuat garis pinggir (Frame) */
    [data-testid="stAppViewContainer"]::before {
        content: '';
        position: fixed;
        top: 10px;
        left: 10px;
        right: 10px;
        bottom: 10px;
        border: 5px solid #87CEFA;
        border-radius: 10px;
        z-index: 999999;
        pointer-events: none;
    }
    
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        padding-left: 3rem;
        padding-right: 3rem;
    }
    </style>
    """, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# DAFTAR NAMA TIM INTI
# -----------------------------------------------------------------------------
TIM_NLP_INTI = [
    "Elvira Nurfadhilah", "Agung Santosa", "Prabu Kresna Putra", "Siska Pebiana",
    "Radhiyatul Fajri", "M. Teduh Uliniansyah", "Lyla Ruslana Aini",
    "Dr. Eng. Ir Yuyun", "Ir. Andi Djalal Latief", "Dr. Rini Wijayanti",
    "Iftitahu Ni'mah", "Nuraisa Novia Hidayati", "Dr. Dipl.(FH) Ing Asril",
    "Ir. Gunarso", "Ir. Tri Sampurno", "Dr. Kokoy Siti Komariah", "Nuryani",
    "Dian Isnaeni Nurul Afra", "Siti Saleha"
]

# -----------------------------------------------------------------------------
# 1. FUNGSI PEMBERSIHAN & EKSTRAKSI
# -----------------------------------------------------------------------------

def _clean_str(x):
    if pd.isna(x): return pd.NA
    s = str(x).strip()
    if s.lower() in {"nan", "none", "", "-", "0"}: return pd.NA
    return s

def _clean_money(x):
    if pd.isna(x): return 0
    s = str(x).replace("Rp", "").replace(" ", "").replace(".", "").replace(",", ".")
    return pd.to_numeric(s, errors='coerce')

def split_contributors(raw: str) -> list[str]:
    if raw is None or pd.isna(raw): return []
    s = str(raw).strip()
    if not s: return []
    s = re.sub(r"\s+(dan|and|&)\s+", ";", s, flags=re.IGNORECASE)
    s = s.replace("|", ";").replace("/", ";").replace("\n", ";")
    if ";" in s: parts = s.split(";")
    else: parts = s.split(";") 
    cleaned_parts = []
    for p in parts:
        p = p.strip()
        if len(p) > 2 and not p.isdigit() and p.lower() not in {"nan", "none"}:
            cleaned_parts.append(p)
    return list(set(cleaned_parts))

def _find_header_row(df_raw: pd.DataFrame, keywords=("judul", "kelompok")):
    kw = [k.lower() for k in keywords]
    for i in range(min(50, len(df_raw))):
        row_str = df_raw.iloc[i].astype(str).str.lower().tolist()
        hit = sum(any(k in cell for k in kw) for cell in row_str)
        if hit >= 1: return i
    return None

def extract_contributors_robust(df: pd.DataFrame, sheet_kind: str) -> pd.Series:
    if sheet_kind in ["SDM_STUDI", "SDM_MOBILITAS"]:
        for col in df.columns:
            if "nama sdm" in str(col).lower():
                return df[col].apply(_clean_str)

    idx_kr = None
    for i, col in enumerate(df.columns):
        if "kelompok" in str(col).lower() and "riset" in str(col).lower():
            idx_kr = i
            break
    
    if idx_kr is not None:
        start = idx_kr + 1
        end = min(start + 15, len(df.columns))
        subset = df.iloc[:, start:end]
        return subset.apply(
            lambda row: "; ".join([str(x).strip() for x in row if pd.notna(x) and str(x).strip() not in ["", "-", "nan", "0"]]),
            axis=1
        )
    
    target_cols = [c for c in df.columns if any(x in str(c).lower() for x in ["penulis", "inventor", "pic", "sivitas"])]
    if target_cols:
        start = df.columns.get_loc(target_cols[0])
        end = min(start + 10, len(df.columns))
        subset = df.iloc[:, start:end]
        return subset.apply(
            lambda row: "; ".join([str(x).strip() for x in row if pd.notna(x) and str(x).strip() not in ["", "-", "nan"]]),
            axis=1
        )

    return pd.Series([pd.NA] * len(df), index=df.index)

def normalize_sheet(df: pd.DataFrame, jenis_luaran: str, sheet_code: str) -> pd.DataFrame:
    if df.empty: return pd.DataFrame()
    df = df.loc[:, ~df.columns.duplicated()]
    col_mapping = {}
    assigned = set()

    for c in df.columns:
        c_low = str(c).lower().strip()
        target = None
        if "judul" in c_low: target = "judul"
        elif "kelompok" in c_low and "riset" in c_low: target = "kelompok_riset"
        elif "periode input" in c_low: target = "periode_input"
        elif "status" in c_low and "upload" not in c_low: target = "status"
        elif "jenis" in c_low and "luaran" not in c_low: target = "jenis_detail"
        elif "reputasi" in c_low: target = "reputasi"
        elif any(x in c_low for x in ["nama jurnal", "prosiding", "mitra", "sumber", "pihak", "universitas"]): target = "venue"
        elif any(x in c_low for x in ["nilai", "dana", "nominal"]): target = "nilai"

        if target and target not in assigned:
            col_mapping[c] = target
            assigned.add(target)

    out = df.rename(columns=col_mapping).copy()
    out["jenis_luaran"] = jenis_luaran
    out["kontributor_raw"] = extract_contributors_robust(df, sheet_code)
    
    for req in ["judul", "kelompok_riset", "venue", "nilai", "status", "reputasi", "jenis_detail"]:
        if req not in out.columns: out[req] = pd.NA

    out["periode_input"] = pd.to_datetime(out.get("periode_input", pd.NaT), errors="coerce")
    out["bulan"] = out["periode_input"].dt.to_period("M").astype(str)
    out["nilai_num"] = out["nilai"].apply(_clean_money)
    
    for c in ["judul", "kelompok_riset", "kontributor_raw", "venue", "status", "reputasi"]:
        out[c] = out[c].apply(_clean_str)

    return out.dropna(subset=["judul", "kelompok_riset"], how="all")

def load_data(excel_path):
    xls = pd.ExcelFile(excel_path)
    all_dfs = []
    configs = [
        ("PI", "PI", "Publikasi Internasional"),
        ("PN", "PN", "Publikasi Nasional"),
        ("KI", "KI", "Kekayaan Intelektual"),
        ("PWRP", "PWRP", "Purwarupa"),
        (["Dana", "PKS"], "DANA", "Dana & Kerjasama"),
        (["SDM", "Studi"], "SDM_STUDI", "SDM (Studi Lanjut)"),
        (["PostDoc", "VR", "Visiting"], "SDM_MOBILITAS", "SDM (Mobilitas/VR)")
    ]

    for keywords, code, label in configs:
        sheet_name = None
        if isinstance(keywords, str): keywords = [keywords]
        for s in xls.sheet_names:
            if any(k.lower() in s.lower() for k in keywords):
                sheet_name = s
                break
        if sheet_name:
            try:
                raw = xls.parse(sheet_name, header=None)
                hdr = _find_header_row(raw)
                df = xls.parse(sheet_name, header=hdr) if hdr is not None else xls.parse(sheet_name)
                all_dfs.append(normalize_sheet(df, label, code))
            except Exception: pass

    if not all_dfs: return pd.DataFrame()
    return pd.concat(all_dfs, ignore_index=True)

# -----------------------------------------------------------------------------
# 3. UI STREAMLIT
# -----------------------------------------------------------------------------

st.title("📊 Capaian Riset 2025: Pengolahan Bahasa Alami (NLP)")

if not DEFAULT_FILE.exists():
    st.error(f"⚠️ File data tidak ditemukan! Pastikan file `{DEFAULT_FILE}` ada di folder yang sama dengan script ini.")
    st.stop()

# 1. LOAD DATA OTOMATIS
with st.spinner("Memuat data..."):
    df_all = load_data(DEFAULT_FILE)

if df_all.empty:
    st.error("Data kosong atau format Excel tidak sesuai template.")
    st.stop()

# 2. AUTO-FILTER KHUSUS NLP
all_kr = sorted([x for x in df_all["kelompok_riset"].dropna().unique() if str(x).strip()])
target_kr = next((x for x in all_kr if "nlp" in x.lower() or "bahasa alami" in x.lower()), None)

if target_kr:
    df_kr = df_all[df_all["kelompok_riset"] == target_kr].copy()
else:
    st.error("Data KR 'Pengolahan Bahasa Alami' tidak ditemukan di file ini.")
    st.stop()

# --- METRICS LENGKAP ---
st.divider()

dana_total = df_kr[df_kr["jenis_luaran"].str.contains("Dana", na=False)]["nilai_num"].sum()
hki_count = len(df_kr[df_kr["jenis_luaran"].str.contains("Kekayaan", na=False)])
pi_count = len(df_kr[df_kr["jenis_luaran"].str.contains("Publikasi Internasional", na=False)])
pn_count = len(df_kr[df_kr["jenis_luaran"].str.contains("Publikasi Nasional", na=False)])
pwrp_count = len(df_kr[df_kr["jenis_luaran"].str.contains("Purwarupa", na=False)])
sdm_vr_count = len(df_kr[df_kr["jenis_luaran"].str.contains("Mobilitas", na=False)])
sdm_studi_count = len(df_kr[df_kr["jenis_luaran"].str.contains("Studi", na=False)])

row1 = st.columns(4)
row1[0].metric("Dana Eksternal", f"Rp {dana_total:,.0f}", border=True)
row1[1].metric("Kekayaan Intelektual", hki_count, border=True)
row1[2].metric("Publikasi Internasional", pi_count, border=True)
row1[3].metric("Publikasi Nasional", pn_count, border=True)

row2 = st.columns(3)
row2[0].metric("Purwarupa", pwrp_count, border=True)
row2[1].metric("SDM (Mobilitas/VR)", sdm_vr_count, border=True)
row2[2].metric("SDM (Studi Lanjut)", sdm_studi_count, border=True)

# --- VISUALISASI SEBARAN ---
st.divider()
st.subheader("📈 Sebaran Distribusi (Seluruh KR)")

c1, c2 = st.columns([1, 1])
with c1:
    st.markdown("**Komposisi Jenis Luaran**")
    fig = px.pie(df_kr, names="jenis_luaran", title="")
    fig.update_traces(textinfo='label+value+percent', textposition='inside')
    fig.update_layout(showlegend=False, margin=dict(t=20, b=20, l=20, r=20))
    st.plotly_chart(fig, use_container_width=True)

with c2:
    st.markdown("**Sebaran Status / Reputasi**")
    df_kr["detail_info"] = df_kr.apply(
        lambda x: x["reputasi"] if pd.notna(x.get("reputasi")) else (x["status"] if pd.notna(x.get("status")) else "Unspecified"), 
        axis=1
    )
    status_counts = df_kr["detail_info"].value_counts().reset_index()
    status_counts.columns = ["Status/Reputasi", "Jumlah"]
    
    if not status_counts.empty:
        fig2 = px.bar(status_counts, x="Status/Reputasi", y="Jumlah", text="Jumlah", color="Jumlah", color_continuous_scale="Blues")
        fig2.update_layout(margin=dict(t=20, b=20, l=20, r=20))
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("Data status/reputasi tidak tersedia.")

# --- HELPERS ---
def is_core_team(name):
    for core in TIM_NLP_INTI:
        core_clean = core.split(",")[0].strip().lower()
        name_clean = name.strip().lower()
        if core_clean in name_clean: return True
    return False

def color_status(val):
    if pd.isna(val): return ""
    val = str(val).lower()
    if any(x in val for x in ['publish', 'suda', 'selesai', 'sertif', 'granted']):
        return 'background-color: #d1e7dd; color: #0f5132' # Hijau
    elif any(x in val for x in ['accept', 'terdaftar', 'submit']):
        return 'background-color: #cff4fc; color: #055160' # Biru
    elif any(x in val for x in ['draft', 'rev', 'belum']):
        return 'background-color: #fff3cd; color: #664d03' # Kuning
    return ""

# --- TABEL DETAIL ---
st.divider()
st.subheader("📋 Rincian Capaian per Kategori")

df_table = df_kr.copy()
patterns = [re.escape(name.split(",")[0].strip()) for name in TIM_NLP_INTI]
pattern_str = "|".join(patterns)
df_table = df_table[df_table["kontributor_raw"].astype(str).str.contains(pattern_str, case=False, na=False)]

unique_types = sorted(df_table["jenis_luaran"].unique())

for tipe in unique_types:
    with st.expander(f"📂 {tipe} (Total: {len(df_table[df_table['jenis_luaran']==tipe])})", expanded=False):
        subset = df_table[df_table["jenis_luaran"] == tipe].copy()
        
        # --- GRAFIK KONTRIBUTOR ---
        cat_contribs = []
        for raw in subset["kontributor_raw"].dropna(): cat_contribs.extend(split_contributors(raw))
        
        if cat_contribs:
            s_cat = pd.Series(cat_contribs).value_counts().reset_index()
            s_cat.columns = ["Nama", "Jumlah"]
            s_cat_core = s_cat[s_cat["Nama"].apply(is_core_team)].copy()
            
            if not s_cat_core.empty:
                st.markdown(f"**Kontribusi Tim Inti di {tipe}:**")
                dyn_height = max(200, len(s_cat_core) * 35)
                fig_sub = px.bar(
                    s_cat_core, x="Jumlah", y="Nama", orientation='h', text="Jumlah",
                    color="Jumlah", color_continuous_scale="Viridis"
                )
                fig_sub.update_layout(yaxis={'categoryorder':'total ascending'}, height=dyn_height, margin=dict(l=0, r=0, t=0, b=0))
                st.plotly_chart(fig_sub, use_container_width=True)
        
        # --- TABEL ---
        show_df = subset[["judul", "venue", "status", "nilai_num", "kontributor_raw"]].rename(columns={
            "judul": "Judul / Kegiatan", "venue": "Venue / Mitra",
            "status": "Status", "nilai_num": "Nilai (Rp)",
            "kontributor_raw": "Kontributor"
        })
        
        styled_df = show_df.style\
            .map(color_status, subset=['Status'])\
            .format({"Nilai (Rp)": "Rp {:,.0f}"})
        
        st.dataframe(
            styled_df, 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "Kontributor": st.column_config.TextColumn(width="medium"),
                "Judul / Kegiatan": st.column_config.TextColumn(width="large")
            }
        )

# -----------------------------------------------------------------------------
# FITUR BARU: WALL OF FAME (DENGAN ANIMASI)
# -----------------------------------------------------------------------------
st.divider()
st.header("🏆 Top Kontributor")
st.caption("")

# Siapkan data REAL terlebih dahulu
all_contribs = []
for raw in df_kr["kontributor_raw"].dropna(): all_contribs.extend(split_contributors(raw))

if all_contribs:
    s_contrib = pd.Series(all_contribs).value_counts().reset_index()
    s_contrib.columns = ["Nama", "Jumlah Output"]
    # Filter hanya Tim Inti
    s_contrib_core = s_contrib[s_contrib["Nama"].apply(is_core_team)].copy()
    
    # Placeholder untuk grafik
    chart_placeholder = st.empty()
    
    # Tombol Aksi
    if st.button("🎲 Tampilkan Peringkat"):
        # 1. Animasi Mengacak (Shuffling Effect)
        progress_text = "Menngolah data..."
        my_bar = st.progress(0, text=progress_text)

        for i in range(15): # Loop animasi 15 frame
            time.sleep(0.1) # Kecepatan animasi
            
            # Buat data dummy random dari nama tim inti
            random_data = pd.DataFrame({
                "Nama": random.sample(TIM_NLP_INTI, 10),
                "Jumlah Output": [random.randint(1, 20) for _ in range(10)]
            })
            
            # Tampilkan chart sementara (dummy)
            fig_temp = px.bar(random_data, x="Jumlah Output", y="Nama", orientation='h', title="🎲 Mengolah...")
            fig_temp.update_layout(xaxis=dict(range=[0, 30])) # Lock axis biar ga goyang
            chart_placeholder.plotly_chart(fig_temp, use_container_width=True)
            my_bar.progress((i + 1) * 6, text=progress_text)
            
        my_bar.empty()
        
        # 2. Tampilkan Data Asli (Final Reveal)
        if not s_contrib_core.empty:
            st.balloons() # Efek Balon
            
            final_height = max(500, len(s_contrib_core) * 40)
            fig_final = px.bar(
                s_contrib_core, 
                x="Jumlah Output", 
                y="Nama", 
                orientation='h', 
                text="Jumlah Output",
                title="✨ Peringkat Produktivitas Tim Inti (Total Output)",
                color="Jumlah Output",
                color_continuous_scale="Viridis"
            )
            fig_final.update_layout(yaxis={'categoryorder':'total ascending'}, height=final_height)
            
            # Replace placeholder dengan grafik asli
            chart_placeholder.plotly_chart(fig_final, use_container_width=True)
        else:
            chart_placeholder.warning("Belum ada data output yang terekam untuk Tim Inti.")
else:
    st.info("Belum ada data kontributor.")
