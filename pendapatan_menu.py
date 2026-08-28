import io
import re
from datetime import datetime

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Header kolom TEMPLATE PENERIMAAN (Other Deposit) Accurate. Setiap
# transaksi diwakili OLEH DUA BARIS: satu baris "HEADER" lalu satu baris
# "ACCOUNT" tepat di bawahnya. Kolom yang sama (posisi B dst) punya arti
# BEDA tergantung baris itu HEADER atau ACCOUNT — makanya file template
# aslinya punya 2 baris label (baris 1 untuk arti kolom saat HEADER, baris
# 2 untuk arti kolom saat ACCOUNT).
# ---------------------------------------------------------------------------
HEADER_ROW_LABELS = [
    "HEADER", "No Penerimaan", "Kas/Bank", "No Cek", "Tanggal", "Tanggal Cek",
    "Kurs", "Catatan", "Cabang", "Pemberi",
    "Kustom Karakter 1", "Kustom Karakter 2", "Kustom Karakter 3", "Kustom Karakter 4",
    "Kustom Karakter 5", "Kustom Karakter 6", "Kustom Karakter 7", "Kustom Karakter 8",
    "Kustom Karakter 9", "Kustom Karakter 10",
    "Kustom Angka 1", "Kustom Angka 2", "Kustom Angka 3", "Kustom Angka 4", "Kustom Angka 5",
    "Kustom Angka 6", "Kustom Angka 7", "Kustom Angka 8", "Kustom Angka 9", "Kustom Angka 10",
    "Kustom Tanggal 1", "Kustom Tanggal 2",
]
ACCOUNT_ROW_LABELS = [
    "ACCOUNT", "No Akun", "Nama Akun", "Nilai Akun", "Nama Departemen", "No Proyek",
    "Catatan Akun",
    "Kategori Keuangan 1", "Kategori Keuangan 2", "Kategori Keuangan 3", "Kategori Keuangan 4",
    "Kategori Keuangan 5", "Kategori Keuangan 6", "Kategori Keuangan 7", "Kategori Keuangan 8",
    "Kategori Keuangan 9", "Kategori Keuangan 10",
] + [None] * 15
N_COLS = 32


def normalize(s: str) -> str:
    s = str(s).upper()
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def find_column(df: pd.DataFrame, candidates: list) -> str:
    cols_norm = {normalize(c): c for c in df.columns}
    for cand in candidates:
        cand_n = normalize(cand)
        if cand_n in cols_norm:
            return cols_norm[cand_n]
    for cand in candidates:
        cand_n = normalize(cand)
        for cn, orig in cols_norm.items():
            if cand_n in cn or cn in cand_n:
                return orig
    return None


def to_excel_bytes(rows: list, sheet_name: str) -> io.BytesIO:
    df = pd.DataFrame(rows)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, header=False, sheet_name=sheet_name)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# Departemen/Lembaga yang tersedia.
# ---------------------------------------------------------------------------
DEPT_UNITS = [
    {"nama": "AGQ1", "dept": "AGQ1", "cabang": "Kantor Pusat"},
    {"nama": "KC01", "dept": "KC01", "cabang": "Kantor Pusat"},
    {"nama": "SD01", "dept": "SD01", "cabang": "Kantor Pusat"},
    {"nama": "SPQ1", "dept": "SPQ1", "cabang": "Kantor Pusat"},
    {"nama": "TD01", "dept": "TD01", "cabang": "Kantor Pusat"},
    {"nama": "TK01", "dept": "TK01", "cabang": "Kantor Pusat"},
    {"nama": "TK02", "dept": "TK02", "cabang": "Kantor Pusat"},
]

# ---------------------------------------------------------------------------
# Tabel konversi No Akun Bank (untuk transaksi Transfer) — sama seperti
# menu Payment lain.
# ---------------------------------------------------------------------------
BANK_ACCOUNT_TABLE = [
    ("BSI TAUD - 1330000038", "11013"),
    ("BSI TKIT 1 - 7364769932", "11030"),
    ("BSI TK 2 - 1330000089", "11017"),
    ("BSI SDIT - 7364770175", "11029"),
    ("BSI TABUNGAN SDIT - 7364770396", "11036"),
    ("BSI - 483 - 7193011483", "11001"),
    ("BSI - 8964252740", "11011"),
    ("BSI KC - 7341714472", "11005"),
    ("BSI LTQ 1 - 7180990156", "11018"),
]
BANK_ACCOUNT_MAP = {normalize(nama): akun for nama, akun in BANK_ACCOUNT_TABLE}

CASH_ACCOUNT_TABLE = [
    ("TD01", "100020305"),
    ("TK01", "100020306"),
    ("TK02", "100020307"),
    ("SD01", "100020303"),
    ("SPQ1", "100020304"),
    ("KC01", "100020302"),
]
CASH_ACCOUNT_MAP = {normalize(dept): akun for dept, akun in CASH_ACCOUNT_TABLE}

# ---------------------------------------------------------------------------
# Tabel konversi No Akun (baris ACCOUNT) per Jenis Pendapatan x Departemen.
# ---------------------------------------------------------------------------
JENIS_ACCOUNT_MAP = {
    "Dana Forsat": {
        "AGQ1": "2100601", "KC01": "2100602", "SD01": "2100603",
        "SPQ1": "2100604", "TD01": "2100605", "TK01": "2100606", "TK02": "2100607",
    },
    "Seragam": {
        "AGQ1": "2100701", "KC01": "2100702", "SD01": "2100703",
        "SPQ1": "2100704", "TD01": "2100705", "TK01": "2100706", "TK02": "2100707",
    },
}

DARI_RE = re.compile(r"^(\S+)\s*-\s*(.+?)\s*\(Kelas", re.IGNORECASE)


def parse_dari(s):
    """Parse kolom 'Dari' format 'NIS - Nama (Kelas : ...)' -> (nis, nama)."""
    if not s or pd.isna(s):
        return "", ""
    s = str(s).strip()
    m = DARI_RE.match(s)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return "", s


def get_kas_bank(metode: str, akun_bank: str, dept: str):
    if metode == "transfer":
        return BANK_ACCOUNT_MAP.get(normalize(akun_bank)) if akun_bank else None
    return CASH_ACCOUNT_MAP.get(normalize(dept))


def render_pendapatan_menu():
    st.title("💰 Template Pendapatan")
    st.caption(
        "Upload laporan pendapatan mentah (Excel), lalu unduh file Template Penerimaan "
        "(Other Deposit) siap-import ke Accurate."
    )

    with st.expander("Lihat tabel konversi No Akun Bank & No Akun Pendapatan"):
        st.write("**No Akun Bank (untuk transaksi Transfer)**")
        st.dataframe(
            pd.DataFrame(BANK_ACCOUNT_TABLE, columns=["Nama Kas & Bank di ISMA", "No Akun"]),
            use_container_width=True, hide_index=True,
        )
        st.write("**No Akun Kas (untuk transaksi Cash, per Departemen)**")
        st.dataframe(
            pd.DataFrame(CASH_ACCOUNT_TABLE, columns=["Departemen", "No Akun"]),
            use_container_width=True, hide_index=True,
        )
        st.write("**No Akun Pendapatan (per Jenis x Departemen)**")
        rows = []
        for jenis, dept_map in JENIS_ACCOUNT_MAP.items():
            for dept, akun in dept_map.items():
                rows.append({"Jenis": jenis, "Departemen": dept, "No Akun": akun})
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    st.header("1. Pilih Lembaga/Departemen")
    dept_list = [u["nama"] for u in DEPT_UNITS]
    selected_dept_nama = st.selectbox("Lembaga/Departemen", dept_list, key="pendapatan_dept")
    selected_unit = next(u for u in DEPT_UNITS if u["nama"] == selected_dept_nama)
    dept = selected_unit["dept"]
    cabang = selected_unit["cabang"]

    col1, col2 = st.columns(2)
    col1.metric("Departemen", dept)
    col2.metric("Cabang", cabang)

    st.header("2. Upload File Laporan Pendapatan (.xlsx)")
    uploaded = st.file_uploader(
        "Pilih file Excel laporan pendapatan (format seperti 'laporan-pendapatan')",
        type=["xlsx", "xls"], key="pendapatan_upload",
    )

    if uploaded is None:
        st.info("Silakan upload file laporan pendapatan (.xlsx) untuk melanjutkan.")
        return

    try:
        raw_df = pd.read_excel(uploaded, header=None)
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        st.stop()

    header_row_idx = None
    for i, row in raw_df.iterrows():
        vals = [normalize(v) for v in row.tolist() if pd.notna(v)]
        if "NO REF" in vals and "JENIS" in vals:
            header_row_idx = i
            break

    if header_row_idx is None:
        st.error(
            "Tidak menemukan baris header (kolom 'No Ref' & 'Jenis') di file ini. "
            "Pastikan formatnya sama seperti laporan pendapatan."
        )
        st.stop()

    headers = raw_df.iloc[header_row_idx].tolist()
    data = raw_df.iloc[header_row_idx + 1:].copy()
    data.columns = headers
    data = data.reset_index(drop=True)

    # Baris ringkasan/Total di bagian bawah file dibuang — hanya baris yang
    # kolom "No"-nya berupa angka transaksi asli yang diproses.
    col_no = find_column(data, ["No"])
    col_noref = find_column(data, ["No Ref"])
    col_tgl = find_column(data, ["Tgl"])
    col_dari = find_column(data, ["Dari"])
    col_jenis = find_column(data, ["Jenis"])
    col_metode = find_column(data, ["Metode"])
    col_akunbank = find_column(data, ["Akun Bank"])
    col_tgltransfer = find_column(data, ["Tgl Transfer"])
    col_jumlah = find_column(data, ["Jumlah"])

    required_missing = [
        name for name, col in [
            ("No", col_no), ("Tgl", col_tgl), ("Dari", col_dari),
            ("Jenis", col_jenis), ("Metode", col_metode),
            ("Akun Bank", col_akunbank), ("Jumlah", col_jumlah),
        ] if col is None
    ]
    if required_missing:
        st.error("Kolom berikut tidak ditemukan di file input: " + ", ".join(required_missing))
        st.stop()

    data["_no_numeric"] = pd.to_numeric(data[col_no], errors="coerce")
    data = data[data["_no_numeric"].notna()].copy()

    st.write("Preview data mentah (setelah baris ringkasan/Total dibuang):")
    st.dataframe(data.head(20), use_container_width=True)

    data["_tgl_parsed"] = pd.to_datetime(data[col_tgl], errors="coerce", dayfirst=True)

    records = []
    for _, r in data.iterrows():
        metode_raw = str(r[col_metode]).strip().lower() if pd.notna(r[col_metode]) else ""
        if metode_raw not in ("cash", "transfer"):
            continue
        tgl = r["_tgl_parsed"]
        if pd.isna(tgl):
            continue
        jumlah = r[col_jumlah]
        if pd.isna(jumlah) or jumlah == 0:
            continue

        jenis = str(r[col_jenis]).strip() if pd.notna(r[col_jenis]) else ""
        nis, nama = parse_dari(r[col_dari])
        akun_bank = str(r[col_akunbank]).strip() if pd.notna(r[col_akunbank]) else ""
        tgl_transfer = str(r[col_tgltransfer]).strip() if col_tgltransfer and pd.notna(r[col_tgltransfer]) else ""

        records.append({
            "Jenis": jenis,
            "Nominal": float(jumlah),
            "Metode": metode_raw,
            "Tanggal": tgl,
            "TglRaw": tgl.strftime("%d-%m-%Y"),
            "Nama": nama,
            "NIS": nis,
            "AkunBank": akun_bank,
            "TglTransfer": tgl_transfer,
        })

    if not records:
        st.warning("Tidak ada transaksi valid yang bisa diproses dari file ini.")
        return

    trans_df = pd.DataFrame(records)

    # Cash: total per (tanggal, jenis) — sama seperti pola cash di menu lain.
    # Transfer: apa adanya per baris.
    cash_df = trans_df[trans_df["Metode"] == "cash"]
    transfer_df = trans_df[trans_df["Metode"] == "transfer"]

    cash_grouped = cash_df.groupby(["Tanggal", "Jenis"], as_index=False).agg(
        Nominal=("Nominal", "sum"), TglRaw=("TglRaw", "first")
    )
    cash_grouped["Metode"] = "cash"
    cash_grouped["Nama"] = ""
    cash_grouped["NIS"] = ""
    cash_grouped["AkunBank"] = ""
    cash_grouped["TglTransfer"] = ""

    final_lines = pd.concat(
        [
            cash_grouped,
            transfer_df[["Tanggal", "Jenis", "Nominal", "Metode", "Nama", "NIS", "AkunBank", "TglTransfer", "TglRaw"]],
        ],
        ignore_index=True,
    )
    final_lines = final_lines.sort_values(["Tanggal", "Jenis"])

    st.header("3. Hasil Konversi")

    output_rows = [HEADER_ROW_LABELS, ACCOUNT_ROW_LABELS]
    unmatched_kasbank_rows = []
    unmatched_akun_rows = []

    for (tanggal, jenis), sub in final_lines.groupby(["Tanggal", "Jenis"]):
        mmYY = tanggal.strftime("%m%y")
        dd = tanggal.strftime("%d")

        for line_no, (_, r) in enumerate(sub.iterrows(), start=1):
            metode_label = r["Metode"].upper()
            kas_bank = get_kas_bank(r["Metode"], r["AkunBank"], dept)
            no_akun = JENIS_ACCOUNT_MAP.get(jenis, {}).get(dept)

            no_penerimaan = f"ISMA-{metode_label}-{jenis}-{dept}-{mmYY}.{dd}.{line_no}"

            if r["Metode"] == "transfer":
                catatan = f"{r['Nama']}_{r['NIS']}"
                catatan_akun = f"{r['AkunBank']}, {r['TglTransfer']}" if r["TglTransfer"] else r["AkunBank"]
            else:
                catatan = ""
                catatan_akun = ""

            header_row = [None] * N_COLS
            header_row[0] = "HEADER"
            header_row[1] = no_penerimaan
            header_row[2] = kas_bank
            header_row[4] = r["TglRaw"]
            header_row[7] = catatan
            header_row[8] = "Kantor Pusat"
            header_row[9] = None  # Pemberi dikosongkan
            output_rows.append(header_row)

            account_row = [None] * N_COLS
            account_row[0] = "ACCOUNT"
            account_row[1] = no_akun
            account_row[3] = r["Nominal"]
            account_row[4] = dept
            account_row[6] = catatan_akun
            output_rows.append(account_row)

            if kas_bank is None:
                unmatched_kasbank_rows.append({
                    "No Penerimaan": no_penerimaan,
                    "Dicari dari": r["AkunBank"] if r["Metode"] == "transfer" else f"CASH - {dept}",
                    "Catatan": "Bank/Departemen tidak ditemukan di tabel konversi No Akun",
                })
            if no_akun is None:
                unmatched_akun_rows.append({
                    "No Penerimaan": no_penerimaan,
                    "Jenis": jenis,
                    "Catatan": "Kombinasi Jenis+Departemen tidak ditemukan di tabel No Akun Pendapatan",
                })

    if unmatched_kasbank_rows:
        st.warning(
            f"Ada {len(unmatched_kasbank_rows)} baris yang Kas/Bank-nya dikosongkan karena "
            "bank/departemen tidak ditemukan di tabel konversi."
        )
        st.dataframe(pd.DataFrame(unmatched_kasbank_rows), use_container_width=True, hide_index=True)

    if unmatched_akun_rows:
        st.warning(
            f"Ada {len(unmatched_akun_rows)} baris yang No Akun-nya dikosongkan karena "
            "kombinasi Jenis+Departemen tidak ada di tabel."
        )
        st.dataframe(pd.DataFrame(unmatched_akun_rows), use_container_width=True, hide_index=True)

    if not unmatched_kasbank_rows and not unmatched_akun_rows:
        st.success("Semua baris berhasil dipetakan tanpa masalah. 🎉")

    preview_df = pd.DataFrame(output_rows[2:])
    st.dataframe(preview_df.fillna(""), use_container_width=True)

    buf = to_excel_bytes(output_rows, "Template Penerimaan")
    st.download_button(
        "⬇️ Download File Template Penerimaan (.xlsx)",
        data=buf,
        file_name=f"Template_Pendapatan_{dept}_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
