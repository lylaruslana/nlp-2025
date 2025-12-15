import re
import time
import random
import zipfile
from pathlib import Path
import pandas as pd
import streamlit as st
import plotly.express as px
import streamlit.components.v1 as components

# -----------------------------------------------------------------------------
# KONFIGURASI HALAMAN & FILE
# -----------------------------------------------------------------------------
st.set_page_config(page_title="Laporan Capaian NLP 2025", page_icon="💎", layout="wide")

# DEFINISI NAMA FILE
FILE_CAPAIAN = Path("Form Capaian PRSDI 2025 final.xlsx")
FILE_TARGET = Path("Target dan Capaian KRNLP 2025.xlsx")

# -----------------------------------------------------------------------------
# CSS INJECTION
# -----------------------------------------------------------------------------
st.markdown("""
    <style>
    [data-testid="stAppViewContainer"] {
    border: 5px solid #87CEFA;
    border-radius: 10px;
    margin: 10px;
    padding: 0;
    box-sizing: border-box;
    background: transparent;
    }
    
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        padding-left: 3rem;
        padding-right: 3rem;
    }

    .metric-container {
        display: flex;
        flex-wrap: wrap;
        gap: 15px;
        justify-content: center;
        margin-bottom: 20px;
        margin-top: 10px;
    }
    .metric-pill {
        background-color: #f0f8ff;
        border: 1px solid #87CEFA;
        border-radius: 50px;
        padding: 8px 20px;
        text-align: center;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        min-width: 140px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
    }
    .metric-label {
        font-size: 0.85rem;
        color: #555;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .metric-value {
        font-size: 1.1rem;
        font-weight: bold;
        color: #000;
        margin-top: 2px;
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
    "Dian Isnaeni Nurul Afra", "Siti Saleha", "Yaniasih"
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
    try:
        xls = pd.ExcelFile(excel_path)
    except zipfile.BadZipFile:
        st.error(f"❌ **File Rusak:** File `{excel_path.name}` tidak dapat dibaca. Kemungkinan file corrupt, kosong (0 bytes), atau formatnya bukan .xlsx (misal CSV yang direname).")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"❌ Terjadi kesalahan saat membaca file `{excel_path.name}`: {e}")
        return pd.DataFrame()

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

st.title("💎 Capaian Riset 2025: Pengolahan Bahasa Alami (NLP)")

if not FILE_CAPAIAN.exists():
    st.error(f"⚠️ File `{FILE_CAPAIAN}` tidak ditemukan! Pastikan file ada di folder ini.")
    st.stop()

# 1. LOAD DATA UTAMA
with st.spinner("Memuat data capaian..."):
    df_all = load_data(FILE_CAPAIAN)

if df_all.empty:
    st.warning("⚠️ Data tidak dapat dimuat. Silakan cek pesan error di atas atau pastikan file Excel tidak rusak.")
    st.stop()

# 2. AUTO-FILTER NLP
all_kr = sorted([x for x in df_all["kelompok_riset"].dropna().unique() if str(x).strip()])
target_kr = next((x for x in all_kr if "nlp" in x.lower() or "bahasa alami" in x.lower()), None)

if target_kr:
    df_kr = df_all[df_all["kelompok_riset"] == target_kr].copy()
else:
    st.error("Data KR 'Pengolahan Bahasa Alami' tidak ditemukan di file ini.")
    st.stop()

# --- HITUNG METRICS ---
dana_total = df_kr[df_kr["jenis_luaran"].str.contains("Dana", na=False)]["nilai_num"].sum()
hki_count = len(df_kr[df_kr["jenis_luaran"].str.contains("Kekayaan", na=False)])
pi_count = len(df_kr[df_kr["jenis_luaran"].str.contains("Publikasi Internasional", na=False)])
pn_count = len(df_kr[df_kr["jenis_luaran"].str.contains("Publikasi Nasional", na=False)])
pwrp_count = len(df_kr[df_kr["jenis_luaran"].str.contains("Purwarupa", na=False)])
sdm_vr_count = len(df_kr[df_kr["jenis_luaran"].str.contains("Mobilitas", na=False)])
sdm_studi_count = len(df_kr[df_kr["jenis_luaran"].str.contains("Studi", na=False)])

# --- TAMPILKAN METRICS (BENTUK OVAL) ---
st.markdown(f"""
<div class="metric-container">
    <div class="metric-pill">
        <div class="metric-label">Dana Eks.</div>
        <div class="metric-value">Rp {dana_total:,.0f}</div>
    </div>
    <div class="metric-pill">
        <div class="metric-label">HKI</div>
        <div class="metric-value">{hki_count}</div>
    </div>
    <div class="metric-pill">
        <div class="metric-label">Pub. Internasional</div>
        <div class="metric-value">{pi_count}</div>
    </div>
    <div class="metric-pill">
        <div class="metric-label">Pub. Nasional</div>
        <div class="metric-value">{pn_count}</div>
    </div>
    <div class="metric-pill">
        <div class="metric-label">Purwarupa</div>
        <div class="metric-value">{pwrp_count}</div>
    </div>
    <div class="metric-pill">
        <div class="metric-label">SDM (Mobilitas/VR)</div>
        <div class="metric-value">{sdm_vr_count}</div>
    </div>
    <div class="metric-pill">
        <div class="metric-label">SDM (Studi)</div>
        <div class="metric-value">{sdm_studi_count}</div>
    </div>
</div>
""", unsafe_allow_html=True)

# --- VISUALISASI SEBARAN ---
st.divider()
st.subheader("💎 Sebaran Distribusi")

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
st.subheader("💎 Rincian Capaian per Kategori")

df_table = df_kr.copy()
patterns = [re.escape(name.split(",")[0].strip()) for name in TIM_NLP_INTI]
pattern_str = "|".join(patterns)
df_table = df_table[df_table["kontributor_raw"].astype(str).str.contains(pattern_str, case=False, na=False)]

unique_types = sorted(df_table["jenis_luaran"].unique())

for tipe in unique_types:
    with st.expander(f"💎 {tipe} (Total: {len(df_table[df_table['jenis_luaran']==tipe])})", expanded=False):
        subset = df_table[df_table["jenis_luaran"] == tipe].copy()
        
        # --- GRAFIK KONTRIBUTOR (SORTED: Count DESC, Name ASC) ---
        cat_contribs = []
        for raw in subset["kontributor_raw"].dropna(): cat_contribs.extend(split_contributors(raw))
        
        if cat_contribs:
            s_cat = pd.Series(cat_contribs).value_counts().reset_index()
            s_cat.columns = ["Nama", "Jumlah"]
            s_cat_core = s_cat[s_cat["Nama"].apply(is_core_team)].copy()
            
            # SORTING UTAMA: Jumlah (Ascending/Kecil->Besar), Nama (Descending/Z->A)
            s_cat_core = s_cat_core.sort_values(by=["Jumlah", "Nama"], ascending=[True, False])
            
            if not s_cat_core.empty:
                st.markdown(f"**Kontribusi Tim di {tipe}:**")
                dyn_height = max(200, len(s_cat_core) * 35)
                fig_sub = px.bar(
                    s_cat_core, x="Jumlah", y="Nama", orientation='h', text="Jumlah",
                    color="Jumlah", color_continuous_scale="Viridis"
                )
                fig_sub.update_layout(height=dyn_height, margin=dict(l=0, r=0, t=0, b=0))
                st.plotly_chart(fig_sub, use_container_width=True)
        
        # --- TABEL KUSTOMISASI PER TIPE ---
        cols_to_show = []
        renames = {"kontributor_raw": "Kontributor"}

        if "Dana" in tipe:
            cols_to_show = ["judul", "venue", "nilai_num", "kontributor_raw"]
            renames.update({"judul": "Judul / Kegiatan", "venue": "Venue / Mitra", "nilai_num": "Nilai (Rp)"})

        elif "Kekayaan" in tipe:
            cols_to_show = ["judul", "status", "kontributor_raw"]
            renames.update({"judul": "Judul / KI", "status": "Status"})

        elif "Publikasi" in tipe:
            cols_to_show = ["judul", "venue", "status", "kontributor_raw"]
            renames.update({"judul": "Judul Artikel", "venue": "Nama Jurnal/Conf", "status": "Status"})

        elif "Purwarupa" in tipe:
            cols_to_show = ["judul", "venue", "status", "kontributor_raw"]
            renames.update({"judul": "Nama Produk", "venue": "Kegiatan/Mitra", "status": "Status"})

        # [MODIFIKASI] MENGHILANGKAN KOLOM NAMA SDM DI TABEL SDM MOBILITAS
        elif "Mobilitas" in tipe:
            cols_to_show = ["status", "kontributor_raw"]
            renames.update({"status": "Judul / Kegiatan"})

        elif "Studi" in tipe:
            cols_to_show = ["venue", "status", "kontributor_raw"]
            renames.update({"venue": "Universitas / Tujuan", "status": "Status"})

        else:
            cols_to_show = ["judul", "venue", "status", "nilai_num", "kontributor_raw"]
            renames.update({"judul": "Judul", "venue": "Venue", "status": "Status", "nilai_num": "Nilai (Rp)"})

        show_df = subset[cols_to_show].rename(columns=renames)
        
        # Styling conditionally
        styler = show_df.style
        
        status_col = None
        if "Status" in show_df.columns: status_col = "Status"
        elif "Judul / Kegiatan" in show_df.columns and "Mobilitas" in tipe: status_col = "Judul / Kegiatan"
        
        if status_col:
            styler = styler.map(color_status, subset=[status_col])
            
        if "Nilai (Rp)" in show_df.columns:
            styler = styler.format({"Nilai (Rp)": "Rp {:,.0f}"})
        
        st.dataframe(
            styler, 
            use_container_width=True, 
            hide_index=True
        )

# =============================================================================
# MONITOR TARGET & CAPAIAN
# =============================================================================
st.divider()
st.header("💎 Monitor Target & Capaian (KRNLP 2025)")

tab_inklusif, tab_eksklusif = st.tabs(["💎 Target Inklusif (Tim)", "💎 Target Eksklusif (Personal)"])

if FILE_TARGET.exists():
    
    # --- TAB 1: INKLUSIF ---
    with tab_inklusif:
        try:
            # Membaca Sheet "Target Inklusif"
            df_ink = pd.read_excel(FILE_TARGET, sheet_name="Target Inklusif", header=1)
            
            # Bersihkan data
            df_ink = df_ink.dropna(how='all', axis=0)
            df_ink = df_ink.loc[:, ~df_ink.columns.str.contains('^Unnamed')] # Hapus kolom tanpa header
            
            # --- VISUALISASI ---
            def extract_num(s):
                if pd.isna(s): return 0
                match = re.search(r'(\d+)', str(s))
                return int(match.group(1)) if match else 0

            df_chart_ink = df_ink.copy()
            df_chart_ink['Num_Target'] = df_chart_ink['Jumlah Target'].apply(extract_num)
            df_chart_ink['Num_Capaian'] = df_chart_ink['Jumlah Capaian'].apply(extract_num)
            
            # Melt data agar bisa diplot grouped bar
            df_melt_ink = df_chart_ink.melt(
                id_vars=["Item"], 
                value_vars=["Num_Target", "Num_Capaian"], 
                var_name="Tipe", 
                value_name="Jumlah"
            )
            df_melt_ink["Tipe"] = df_melt_ink["Tipe"].replace({"Num_Target": "Target", "Num_Capaian": "Capaian"})
            
            c_ink_1, c_ink_2 = st.columns([1, 1.5])
            
            with c_ink_1:
                st.dataframe(df_ink, use_container_width=True, hide_index=True)
            
            with c_ink_2:
                fig_ink = px.bar(
                    df_melt_ink, 
                    x="Item", 
                    y="Jumlah", 
                    color="Tipe", 
                    barmode="group",
                    title="Perbandingan Target vs Capaian",
                    color_discrete_map={"Target": "#FF7F50", "Capaian": "#40E0D0"} # Orange & Turquoise
                )
                st.plotly_chart(fig_ink, use_container_width=True)

        except zipfile.BadZipFile:
            st.error(f"❌ File '{FILE_TARGET.name}' rusak atau bukan format .xlsx yang valid.")
        except Exception as e:
            st.warning(f"Info: {e}. Pastikan sheet 'Target Inklusif' tersedia.")

    # --- TAB 2: EKSKLUSIF ---
    with tab_eksklusif:
        try:
            # Membaca Sheet "Target Eksklusif"
            # Data dimulai dari baris ke-4 (index 3) di Excel
            df_eks_raw = pd.read_excel(FILE_TARGET, sheet_name="Target Eksklusif", header=None, skiprows=3)
            
            # AMBIL 4 KATEGORI EKSKLUSIF: Pub Int, KI, Studi, KKM
            df_eks_clean = df_eks_raw.iloc[:, 1:10].copy()
            
            # Rename Kolom
            df_eks_clean.columns = [
                "Nama", 
                "Pub. Int (T)", "Pub. Int (C)", 
                "KI (T)", "KI (C)",
                "Studi (T)", "Studi (C)",
                "KKM (T)", "KKM (C)"
            ]
            
            # Bersihkan baris
            df_eks_clean = df_eks_clean.dropna(subset=["Nama"])
            df_eks_clean = df_eks_clean[df_eks_clean["Nama"].str.lower() != 'total']
            
            # Bersihkan angka (ganti '-' dengan 0)
            num_cols = [c for c in df_eks_clean.columns if c != "Nama"]
            for c in num_cols:
                df_eks_clean[c] = pd.to_numeric(
                    df_eks_clean[c].astype(str).str.replace('-', '0'), 
                    errors='coerce'
                ).fillna(0)

            # MENAMPILKAN DATA (TANPA TOTAL AKUMULASI)
            st.dataframe(df_eks_clean, use_container_width=True, hide_index=True)
            
            st.markdown("### 💎 Rincian Target vs Capaian per Kategori")
            
            # SIAPKAN DATA UNTUK GRAFIK RINCI (FACETED)
            rows = []
            for _, row in df_eks_clean.iterrows():
                name = row['Nama']
                # Pub Int
                rows.append({'Nama': name, 'Kategori': 'Pub. Int', 'Tipe': 'Target', 'Jumlah': row['Pub. Int (T)']})
                rows.append({'Nama': name, 'Kategori': 'Pub. Int', 'Tipe': 'Capaian', 'Jumlah': row['Pub. Int (C)']})
                # KI
                rows.append({'Nama': name, 'Kategori': 'KI', 'Tipe': 'Target', 'Jumlah': row['KI (T)']})
                rows.append({'Nama': name, 'Kategori': 'KI', 'Tipe': 'Capaian', 'Jumlah': row['KI (C)']})
                # Studi
                rows.append({'Nama': name, 'Kategori': 'Studi', 'Tipe': 'Target', 'Jumlah': row['Studi (T)']})
                rows.append({'Nama': name, 'Kategori': 'Studi', 'Tipe': 'Capaian', 'Jumlah': row['Studi (C)']})
                # KKM
                rows.append({'Nama': name, 'Kategori': 'KKM', 'Tipe': 'Target', 'Jumlah': row['KKM (T)']})
                rows.append({'Nama': name, 'Kategori': 'KKM', 'Tipe': 'Capaian', 'Jumlah': row['KKM (C)']})
            
            df_detail = pd.DataFrame(rows)
            
            # Buat Facet Plot Vertikal (1 Kolom agar mudah dibandingkan per orang)
            # facet_row="Kategori" akan membuat 4 baris chart (Pub, KI, Studi, KKM)
            fig_eks_detail = px.bar(
                df_detail, 
                x="Nama", 
                y="Jumlah", 
                color="Tipe", 
                barmode="group",
                facet_col="Kategori", 
                facet_col_wrap=1,  # Stack secara vertikal
                height=1000,       # Tinggi disesuaikan
                title="Detail Target vs Capaian per Kategori",
                color_discrete_map={"Target": "#FF7F50", "Capaian": "#40E0D0"}
            )
            # Pastikan label X (Nama) terlihat di semua subplot
            fig_eks_detail.update_xaxes(matches='x', showticklabels=True)
            st.plotly_chart(fig_eks_detail, use_container_width=True)
            
        except zipfile.BadZipFile:
            st.error(f"❌ File '{FILE_TARGET.name}' rusak atau bukan format .xlsx yang valid.")
        except Exception as e:
            st.warning(f"Info: {e}. Pastikan sheet 'Target Eksklusif' tersedia.")
else:
    st.error(f"⚠️ File '{FILE_TARGET}' tidak ditemukan di direktori yang sama.")

# -----------------------------------------------------------------------------
# FITUR: WALL OF FAME (DIPINDAHKAN KE PALING BAWAH)
# -----------------------------------------------------------------------------
st.divider()
st.markdown('<div id="top-kontributor"></div>', unsafe_allow_html=True)
st.header("💎 Top Kontributor")
st.caption("")

# [PERBAIKAN] LOGIKA PERHITUNGAN KONTRIBUTOR DITARUH DISINI
all_contribs = []
for raw in df_kr["kontributor_raw"].dropna(): all_contribs.extend(split_contributors(raw))

if all_contribs:
    s_contrib = pd.Series(all_contribs).value_counts().reset_index()
    s_contrib.columns = ["Nama", "Jumlah Output"]
    s_contrib_core = s_contrib[s_contrib["Nama"].apply(is_core_team)].copy()
    
    # SORTING FINAL (Sama seperti grafik di atas)
    s_contrib_core = s_contrib_core.sort_values(by=["Jumlah Output", "Nama"], ascending=[True, False])
    
    chart_placeholder = st.empty()
    
    if st.button(" Tampilkan Peringkat"):
        components.html("""
        <script>
        const doc = window.parent.document;
        const el = doc.getElementById('top-kontributor');
        if (el) el.scrollIntoView({behavior: 'smooth', block: 'start'});
        </script>
        """, height=0)

        progress_text = "Mengolah data..."
        my_bar = st.progress(0, text=progress_text)

        for i in range(15):
            time.sleep(0.1)
            random_data = pd.DataFrame({
                "Nama": random.sample(TIM_NLP_INTI, 10),
                "Jumlah Output": [random.randint(1, 20) for _ in range(10)]
            })
            # Acak grafik dummy
            fig_temp = px.bar(random_data, x="Jumlah Output", y="Nama", orientation='h', title="🎲 Mengolah...")
            fig_temp.update_layout(xaxis=dict(range=[0, 30]))
            chart_placeholder.plotly_chart(fig_temp, use_container_width=True)
            my_bar.progress((i + 1) * 6, text=progress_text)
            
        my_bar.empty()
        
        if not s_contrib_core.empty:
            st.balloons()
            final_height = max(500, len(s_contrib_core) * 40)
            fig_final = px.bar(
                s_contrib_core, 
                x="Jumlah Output", 
                y="Nama", 
                orientation='h', 
                text="Jumlah Output",
                title="✨ Peringkat Berdasar Keterlibatan",
                color="Jumlah Output",
                color_continuous_scale="Viridis"
            )
            fig_final.update_layout(height=final_height)
            chart_placeholder.plotly_chart(fig_final, use_container_width=True)
        else:
            chart_placeholder.warning("Belum ada data output yang terekam untuk Tim.")
else:
    st.info("Belum ada data kontributor.")
