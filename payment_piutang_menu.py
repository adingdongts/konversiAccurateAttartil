import io
import re
from datetime import datetime

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Header kolom TEMPLATE PAYMENT Accurate (urutan harus persis)
# ---------------------------------------------------------------------------
PAYMENT_HEADERS = [
    "CUSTOMER NO", "NUMBER", "BRANCH", "DATE", "EXPENSE ACCOUNT NO",
    "DESCRIPTION", "PAYMENT TOTAL", "PAYMENT NUMBER", "PAYMENT VALUE",
    "PAYING BANK", "DISCOUNT ACCOUNT NO", "TOTAL DISCOUNT",
]


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


def to_excel_bytes(df: pd.DataFrame, sheet_name: str) -> io.BytesIO:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# Master data customer (sama seperti menu ISMA / Payment Isma)
# ---------------------------------------------------------------------------
PAYMENT_UNITS = [
    {"id": "Siswa-KC01", "nama": "Siswa-KC01", "dept": "KC01", "cabang": "Kantor Pusat"},
    {"id": "Siswa-SD01", "nama": "Siswa-SD01", "dept": "SD01", "cabang": "Kantor Pusat"},
    {"id": "Siswa-SPQ1", "nama": "Siswa-SPQ1", "dept": "SPQ1", "cabang": "Kantor Pusat"},
    {"id": "Siswa-TD01", "nama": "Siswa-TD01", "dept": "TD01", "cabang": "Kantor Pusat"},
    {"id": "Siswa-TK01", "nama": "Siswa-TK01", "dept": "TK01", "cabang": "Kantor Pusat"},
    {"id": "Siswa-TK02", "nama": "Siswa-TK02", "dept": "TK02", "cabang": "Kantor Pusat"},
]

# ---------------------------------------------------------------------------
# Tabel konversi No Akun (Bank untuk transfer, Kas untuk cash) — sama
# seperti menu Payment Isma sebelumnya.
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
BANK_ACCOUNT_TABLE = [(k, v) for k, v in BANK_ACCOUNT_TABLE if v]
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
# Referensi No Piutang (digabung dari SALDO_PIUTANG_TA_2026-2027.xlsx dan
# list_piutang_versi_1.xlsx) — hanya dipakai untuk memberi WARNING kalau
# NUMBER hasil konversi tidak ditemukan di daftar ini. Tidak menghalangi
# proses, karena bisa saja invoice baru yang belum ada di daftar.
# ---------------------------------------------------------------------------
KNOWN_PIUTANG_NUMBERS = {
    "KC01.BK-2026", "KC01.DKS-2026", "KC01.IP-2026", "KC01.SRGM-2026",
    "SD01.BK-2020", "SD01.BK-2021", "SD01.BK-2022", "SD01.BK-2023",
    "SD01.BK-2024", "SD01.BK-2025", "SD01.BK-2026", "SD01.DKS-2020",
    "SD01.DKS-2022", "SD01.DKS-2023", "SD01.DKS-2024", "SD01.DKS-2025",
    "SD01.DKS-2026", "SD01.IP-2018", "SD01.IP-2019", "SD01.IP-2020",
    "SD01.IP-2021", "SD01.IP-2022", "SD01.IP-2023", "SD01.IP-2024",
    "SD01.IP-2025", "SD01.IP-2026", "SD01.PKBM-2021", "SD01.PKBM-2023",
    "SD01.PKBM-2024", "SD01.PKBM-2025", "SD01.PKBM-2026", "SD01.SPP-2020",
    "SD01.SPP-2021", "SD01.SPP-2207", "SD01.SPP-2208", "SD01.SPP-2209",
    "SD01.SPP-2210", "SD01.SPP-2211", "SD01.SPP-2212", "SD01.SPP-2301",
    "SD01.SPP-2302", "SD01.SPP-2303", "SD01.SPP-2304", "SD01.SPP-2305",
    "SD01.SPP-2306", "SD01.SPP-2307", "SD01.SPP-2308", "SD01.SPP-2309",
    "SD01.SPP-2310", "SD01.SPP-2311", "SD01.SPP-2312", "SD01.SPP-2401",
    "SD01.SPP-2402", "SD01.SPP-2403", "SD01.SPP-2404", "SD01.SPP-2405",
    "SD01.SPP-2406", "SD01.SPP-2407", "SD01.SPP-2408", "SD01.SPP-2409",
    "SD01.SPP-2410", "SD01.SPP-2411", "SD01.SPP-2412", "SD01.SPP-2501",
    "SD01.SPP-2502", "SD01.SPP-2503", "SD01.SPP-2504", "SD01.SPP-2505",
    "SD01.SPP-2506", "SD01.SPP-2507", "SD01.SPP-2508", "SD01.SPP-2509",
    "SD01.SPP-2510", "SD01.SPP-2511", "SD01.SPP-2512", "SD01.SPP-2601",
    "SD01.SPP-2602", "SD01.SPP-2603", "SD01.SPP-2604", "SD01.SPP-2605",
    "SD01.SPP-2606", "SPQ.DKS-2022", "SPQ.DKS-2023", "SPQ.KES-2022",
    "SPQ.KES-2023", "SPQ.LKS-2022", "SPQ.LKS-2023", "SPQ.PERFAS-2023",
    "SPQ.SPP-2207", "SPQ.SPP-2208", "SPQ.SPP-2209", "SPQ.SPP-2210",
    "SPQ.SPP-2211", "SPQ.SPP-2212", "SPQ.SPP-2301", "SPQ.SPP-2302",
    "SPQ.SPP-2303", "SPQ1.BK-2026", "SPQ1.DB-2026", "SPQ1.DKS-2024",
    "SPQ1.DKS-2025", "SPQ1.DKS-2026", "SPQ1.IP-2021", "SPQ1.IP-2025",
    "SPQ1.IP-2026", "SPQ1.IPM-2021", "SPQ1.KES-2024", "SPQ1.KSHTN-2026",
    "SPQ1.LKS-2025", "SPQ1.LKS-2026", "SPQ1.PF-2025", "SPQ1.PF-2026",
    "SPQ1.PPS-2021", "SPQ1.PPS-2026", "SPQ1.SEWA-2021", "SPQ1.SEWA-2024",
    "SPQ1.SPA-2025", "SPQ1.SPA-2026", "SPQ1.SPP-2021", "SPQ1.SPP-2412",
    "SPQ1.SPP-2501", "SPQ1.SPP-2502", "SPQ1.SPP-2503", "SPQ1.SPP-2504",
    "SPQ1.SPP-2505", "SPQ1.SPP-2506", "SPQ1.SPP-2507", "SPQ1.SPP-2508",
    "SPQ1.SPP-2509", "SPQ1.SPP-2510", "SPQ1.SPP-2511", "SPQ1.SPP-2512",
    "SPQ1.SPP-2601", "SPQ1.SPP-2602", "SPQ1.SPP-2603", "SPQ1.SPP-2604",
    "SPQ1.SPP-2605", "SPQ1.SPP-2606", "SPQ1.SRGM-2024", "SPQ1.SRGM-2025",
    "SPQ1.SRGM-2026", "TD01.BK-2025", "TD01.BK-2026", "TD01.DKS-2025",
    "TD01.DKS-2026", "TD01.DU-2026", "TD01.IP-2025", "TD01.IP-2026",
    "TD01.MRTL-2026", "TD01.PKBM-2025", "TD01.PKBM-2026", "TD01.SPP-2601",
    "TD01.SPP-2602", "TD01.SPP-2603", "TD01.SPP-2604", "TD01.SPP-2605",
    "TD01.SPP-2606", "TK01.BK-2021", "TK01.BK-2022", "TK01.BK-2026",
    "TK01.DKS-2021", "TK01.DKS-2022", "TK01.DKS-2026", "TK01.DU-2022",
    "TK01.DU-2026", "TK01.IP-2021", "TK01.IP-2024", "TK01.IP-2025",
    "TK01.IP-2026", "TK01.PKBM-2021", "TK01.PKBM-2025", "TK01.PKBM-2026",
    "TK01.SPP-2107", "TK01.SPP-2108", "TK01.SPP-2109", "TK01.SPP-2110",
    "TK01.SPP-2111", "TK01.SPP-2112", "TK01.SPP-2201", "TK01.SPP-2202",
    "TK01.SPP-2203", "TK01.SPP-2204", "TK01.SPP-2205", "TK01.SPP-2206",
    "TK01.SPP-2209", "TK01.SPP-2210", "TK01.SPP-2211", "TK01.SPP-2212",
    "TK01.SPP-2301", "TK01.SPP-2302", "TK01.SPP-2303", "TK01.SPP-2304",
    "TK01.SPP-2305", "TK01.SPP-2306", "TK01.SPP-2507", "TK01.SPP-2508",
    "TK01.SPP-2509", "TK01.SPP-2510", "TK01.SPP-2511", "TK01.SPP-2512",
    "TK01.SPP-2601", "TK01.SPP-2602", "TK01.SPP-2603", "TK01.SPP-2604",
    "TK01.SPP-2605", "TK01.SPP-2606", "TK01.SRGM-2021", "TK02.BK-2025",
    "TK02.BK-2026", "TK02.DKS-2024", "TK02.DKS-2025", "TK02.DKS-2026",
    "TK02.DU-2025", "TK02.DU-2026", "TK02.IP-2024", "TK02.IP-2025",
    "TK02.IP-2026", "TK02.PKBM-2024", "TK02.PKBM-2026", "TK02.SPP-2501",
    "TK02.SPP-2502", "TK02.SPP-2503", "TK02.SPP-2504", "TK02.SPP-2505",
    "TK02.SPP-2506", "TK02.SPP-2507", "TK02.SPP-2508", "TK02.SPP-2509",
    "TK02.SPP-2510", "TK02.SPP-2511", "TK02.SPP-2512", "TK02.SPP-2601",
    "TK02.SPP-2602", "TK02.SPP-2603", "TK02.SPP-2604", "TK02.SPP-2605",
    "TK02.SPP-2606",
}

# Kolom produk yang diproses & kode invoice-nya di Accurate.
PRODUCT_COLUMNS = ["Buku", "PKBM", "DKS", "SPP", "IP", "DB", "MRTL", "DU"]
PRODUCT_CODE_MAP = {
    "Buku": "BK", "PKBM": "PKBM", "DKS": "DKS",
    "SPP": "SPP", "IP": "IP", "DB": "DB",
    "MRTL": "MRTL", "DU": "DU",
}

INDO_MONTH_NAMES = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
    7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember",
}


def parse_periode_year(s):
    """Ambil tahun pertama dari kolom Periode, misal '2025/2026' -> '2025'."""
    if pd.isna(s):
        return None
    m = re.search(r"(\d{4})", str(s))
    return m.group(1) if m else None


def get_expense_account(metode: str, akun_bank: str, dept: str):
    if metode == "transfer":
        return BANK_ACCOUNT_MAP.get(normalize(akun_bank)) if akun_bank else None
    else:
        return CASH_ACCOUNT_MAP.get(normalize(dept))


def render_payment_piutang_menu():
    st.title("💳 Konversi Payment Piutang ISMA")
    st.caption(
        "Upload laporan pembayaran mentah ISMA (Excel), lalu unduh file Template "
        "Payment siap-import ke Accurate. NUMBER dibentuk mengikuti pola nomor "
        "piutang di Accurate (Buku/PKBM/DKS/IP/DB berdasarkan Periode, SPP "
        "berdasarkan Bulan Tagihan)."
    )

    with st.expander("Lihat tabel konversi No Akun (Bank & Cash)"):
        st.write("**Akun Bank (untuk transaksi Transfer)**")
        st.dataframe(
            pd.DataFrame(BANK_ACCOUNT_TABLE, columns=["Nama Kas & Bank di ISMA", "No Akun"]),
            use_container_width=True, hide_index=True,
        )
        st.write("**Akun Cash (untuk transaksi Tunai, per Departemen)**")
        st.dataframe(
            pd.DataFrame(CASH_ACCOUNT_TABLE, columns=["Departemen", "No Akun"]),
            use_container_width=True, hide_index=True,
        )

    st.header("1. Pilih Customer (Unit/Departemen)")
    unit_nama_list = [u["nama"] for u in PAYMENT_UNITS]
    selected_unit_nama = st.selectbox("Nama Pelanggan", unit_nama_list, key="payment_piutang_unit")
    selected_unit = next(u for u in PAYMENT_UNITS if u["nama"] == selected_unit_nama)
    customer_no = selected_unit["id"]
    dept = selected_unit["dept"]
    cabang = selected_unit["cabang"]

    col1, col2, col3 = st.columns(3)
    col1.metric("Customer No", customer_no)
    col2.metric("Departemen", dept)
    col3.metric("Cabang", cabang)

    st.header("2. Upload File Laporan Pembayaran (.xlsx)")
    uploaded = st.file_uploader(
        "Pilih file Excel laporan pembayaran (format seperti 'Lap Pembayaran')",
        type=["xlsx", "xls"], key="payment_piutang_upload",
    )

    if uploaded is None:
        st.info("Silakan upload file laporan pembayaran (.xlsx) untuk melanjutkan.")
        return

    try:
        raw_df = pd.read_excel(uploaded, header=None)
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        st.stop()

    # cari baris header (baris yang mengandung 'No Ref' & 'Siswa'), sama
    # seperti pola pembacaan file di menu ISMA
    header_row_idx = None
    for i, row in raw_df.iterrows():
        vals = [normalize(v) for v in row.tolist() if pd.notna(v)]
        if "NO REF" in vals and "SISWA" in vals:
            header_row_idx = i
            break

    if header_row_idx is None:
        st.error(
            "Tidak menemukan baris header (kolom 'No Ref' & 'Siswa') di file ini. "
            "Pastikan formatnya sama seperti laporan pembayaran ISMA."
        )
        st.stop()

    headers = raw_df.iloc[header_row_idx].tolist()
    data = raw_df.iloc[header_row_idx + 1:].copy()
    data.columns = headers
    data = data.dropna(how="all")

    st.write("Preview data mentah:")
    st.dataframe(data.head(20), use_container_width=True)

    col_no = find_column(data, ["No"])
    col_noref = find_column(data, ["No Ref"])
    col_tgl = find_column(data, ["Tgl"])
    col_siswa = find_column(data, ["Siswa"])
    col_nis = find_column(data, ["NIS"])
    col_periode = find_column(data, ["Periode"])
    col_bulan = find_column(data, ["Bulan"])
    col_metode = find_column(data, ["Metode"])
    col_akunbank = find_column(data, ["Akun Bank"])
    col_tgltransfer = find_column(data, ["Tgl Transfer"])

    product_cols = {}
    for prod in PRODUCT_COLUMNS:
        c = find_column(data, [prod])
        if c:
            product_cols[prod] = c

    required_missing = [
        name for name, col in [
            ("Tgl", col_tgl), ("Siswa", col_siswa), ("NIS", col_nis),
            ("Periode", col_periode), ("Bulan", col_bulan),
            ("Metode", col_metode), ("Akun Bank", col_akunbank),
        ] if col is None
    ]
    if required_missing:
        st.error("Kolom berikut tidak ditemukan di file input: " + ", ".join(required_missing))
        st.stop()
    if not product_cols:
        st.error(
            "Tidak ada kolom produk (Buku/PKBM/DKS/SPP/IP/DB) yang ditemukan di file ini."
        )
        st.stop()

    # forward-fill kolom yang kosong di baris kontinuasi
    ffill_cols = [
        c for c in [col_no, col_noref, col_tgl, col_periode, col_bulan,
                    col_metode, col_akunbank, col_tgltransfer, col_siswa, col_nis]
        if c is not None
    ]
    data[ffill_cols] = data[ffill_cols].ffill()

    data["_tgl_parsed"] = pd.to_datetime(data[col_tgl], errors="coerce", dayfirst=True)

    # Parse kolom Bulan sebagai TANGGAL asli (bukan tebak-tebakan teks),
    # supaya tidak salah baca format apapun bentuk aslinya di Excel
    # (angka, tanggal, atau teks tanggal).
    data["_bulan_parsed"] = pd.to_datetime(data[col_bulan], errors="coerce", dayfirst=True)

    bulan_periods = sorted({
        (d.year, d.month) for d in data["_bulan_parsed"].dropna()
    })
    if not bulan_periods:
        st.error(
            "Tidak bisa membaca kolom Bulan sebagai tanggal/periode. "
            "Cek isi kolom Bulan di file — formatnya tidak dikenali."
        )
        st.stop()

    st.header("3. Pilih Bulan Tagihan (khusus untuk SPP)")
    st.caption(
        "Khusus SPP: hanya baris dengan Bulan **sebelum** bulan yang dipilih di sini "
        "yang diproses (dianggap tunggakan/backlog). Bulan yang dipilih sendiri dan "
        "bulan sesudahnya TIDAK diproses. NUMBER tiap baris SPP tetap dibentuk sesuai "
        "bulan aslinya masing-masing (format DEPT.SPP-YYMM). Produk lain "
        "(Buku/PKBM/DKS/IP/DB) tidak dipengaruhi pilihan ini — dipakai kolom Periode masing-masing."
    )
    bulan_labels = [f"{INDO_MONTH_NAMES[m]} {y}" for y, m in bulan_periods]
    selected_label = st.selectbox("Bulan Tagihan (untuk SPP)", bulan_labels, key="payment_piutang_bulan")
    selected_year, selected_month = bulan_periods[bulan_labels.index(selected_label)]

    st.header("4. Pilih Periode Tagihan (khusus untuk selain SPP)")
    st.caption(
        "Dipakai untuk membentuk NUMBER produk Buku/PKBM/DKS/IP/DB "
        "(format DEPT.KODE-YYYY, tahun diambil dari tahun awal Periode). "
        "Hanya baris selain-SPP yang Periode-nya cocok yang diproses."
    )
    periode_options = sorted({
        str(v).strip() for v in data[col_periode].dropna().tolist() if str(v).strip()
    })
    if not periode_options:
        st.error("Tidak ditemukan nilai Periode di file ini.")
        st.stop()
    selected_periode = st.selectbox("Periode Tagihan (untuk selain SPP)", periode_options, key="payment_piutang_periode")
    selected_periode_year = parse_periode_year(selected_periode)
    if selected_periode_year is None:
        st.error(
            f"Tidak bisa membaca tahun dari Periode '{selected_periode}'. "
            "Pastikan formatnya seperti '2025/2026'."
        )
        st.stop()

    # ----- kumpulkan baris transaksi per produk -----
    records = []
    for _, r in data.iterrows():
        metode_raw = str(r[col_metode]).strip().lower() if pd.notna(r[col_metode]) else ""
        if metode_raw not in ("cash", "transfer"):
            continue
        tgl = r["_tgl_parsed"]
        if pd.isna(tgl):
            continue

        for prod, col in product_cols.items():
            val = r[col]
            if pd.isna(val) or val == 0:
                continue

            if prod == "SPP":
                bulan_val = r["_bulan_parsed"]
                if pd.isna(bulan_val):
                    continue
                if (bulan_val.year, bulan_val.month) >= (selected_year, selected_month):
                    continue
                spp_number = f"{dept}.SPP-{bulan_val.year % 100:02d}{bulan_val.month:02d}"
                number = spp_number
            else:
                periode_val = str(r[col_periode]).strip() if pd.notna(r[col_periode]) else None
                if periode_val != selected_periode:
                    continue
                kode = PRODUCT_CODE_MAP[prod]
                number = f"{dept}.{kode}-{selected_periode_year}"

            records.append({
                "Produk": prod,
                "Number": number,
                "Nominal": float(val),
                "Metode": metode_raw,
                "Tanggal": tgl,
                "TglRaw": r[col_tgl],
                "Siswa": str(r[col_siswa]).strip() if pd.notna(r[col_siswa]) else "",
                "NIS": str(r[col_nis]).strip() if pd.notna(r[col_nis]) else "",
                "AkunBank": str(r[col_akunbank]).strip() if pd.notna(r[col_akunbank]) else "",
            })

    if not records:
        st.warning(
            "Tidak ada transaksi yang cocok untuk diproses. Coba pilih Bulan Tagihan atau "
            "Periode Tagihan lain, atau cek isi kolom Periode/Bulan di file."
        )
        return

    trans_df = pd.DataFrame(records)

    # Buang baris yang PERSIS sama (siswa, tanggal, produk, number, nominal,
    # metode, akun bank) — biasanya karena baris kontinuasi ter-forward-fill
    # dobel atau data mentahnya memang ada duplikat. Kalau dibiarkan, hasil
    # akhirnya bisa punya NUMBER/PAYMENT NUMBER yang sama persis 2x, yang
    # tidak boleh terjadi saat import ke Accurate.
    before_dedup = len(trans_df)
    trans_df = trans_df.drop_duplicates(
        subset=["Produk", "Number", "Nominal", "Metode", "Tanggal", "Siswa", "NIS", "AkunBank"]
    ).reset_index(drop=True)
    n_dupe = before_dedup - len(trans_df)
    if n_dupe > 0:
        st.warning(
            f"Ditemukan {n_dupe} baris transaksi yang persis duplikat (siswa, tanggal, "
            "produk, nominal, metode sama semua) — otomatis dibuang supaya tidak dobel "
            "di hasil akhir."
        )

    # Cash: total per (tanggal, produk, number) — Siswa/NIS hilang karena digabung.
    # Transfer: apa adanya per baris — Siswa/NIS/AkunBank dipertahankan.
    cash_df = trans_df[trans_df["Metode"] == "cash"]
    transfer_df = trans_df[trans_df["Metode"] == "transfer"]

    cash_grouped = cash_df.groupby(["Tanggal", "Produk", "Number"], as_index=False).agg(
        Nominal=("Nominal", "sum"), TglRaw=("TglRaw", "first")
    )
    cash_grouped["Metode"] = "cash"
    cash_grouped["Siswa"] = ""
    cash_grouped["NIS"] = ""
    cash_grouped["AkunBank"] = ""

    final_lines = pd.concat(
        [
            cash_grouped,
            transfer_df[["Tanggal", "Produk", "Number", "Nominal", "Metode", "Siswa", "NIS", "AkunBank", "TglRaw"]],
        ],
        ignore_index=True,
    )
    final_lines = final_lines.sort_values(["Tanggal", "Produk"])

    st.header("5. Hasil Konversi")

    output_rows = []
    unmatched_account_rows = []
    unknown_number_rows = []
    reported_numbers = set()

    # PENTING: nomor urut (line_no) di kolom NUMBER (ISMA-...-dd.N) harus
    # berjalan terus per (Tanggal, Produk) — TIDAK boleh direset per invoice
    # (Number/piutang), supaya tidak ada 2 baris dengan NUMBER yang sama
    # persis walau invoice-nya beda (misal 1 transaksi mencakup tunggakan
    # beberapa bulan sekaligus, tiap bulan invoice-nya beda tapi tetap 1
    # tanggal+produk yang sama).
    for (tanggal, produk), sub in final_lines.groupby(["Tanggal", "Produk"]):
        mmYY = tanggal.strftime("%m%y")
        dd = tanggal.strftime("%d")

        for line_no, (_, r) in enumerate(sub.iterrows(), start=1):
            number = r["Number"]
            if number not in KNOWN_PIUTANG_NUMBERS and number not in reported_numbers:
                reported_numbers.add(number)
                unknown_number_rows.append({
                    "NUMBER": number, "Produk": produk,
                    "Catatan": "Tidak ditemukan di daftar piutang yang sudah ditanam di aplikasi",
                })

            metode_label = r["Metode"].upper()
            expense_account = get_expense_account(r["Metode"], r["AkunBank"], dept)

            row = {h: None for h in PAYMENT_HEADERS}
            row["CUSTOMER NO"] = customer_no
            row["NUMBER"] = f"ISMA-{metode_label}-{produk}-{dept}-{mmYY}-{dd}.{line_no}_P"
            row["BRANCH"] = "Kantor Pusat"
            row["DATE"] = r["TglRaw"]
            row["EXPENSE ACCOUNT NO"] = expense_account
            row["PAYMENT TOTAL"] = r["Nominal"]
            row["PAYMENT NUMBER"] = number
            row["PAYMENT VALUE"] = r["Nominal"]
            row["PAYING BANK"] = "Transfer Bank" if r["Metode"] == "transfer" else "Cash"

            if r["Metode"] == "transfer":
                row["DESCRIPTION"] = f"{r['Siswa']}_{r['NIS']}"
            else:
                row["DESCRIPTION"] = number

            output_rows.append(row)

            if expense_account is None:
                unmatched_account_rows.append({
                    "NUMBER": number,
                    "PAYING BANK": row["PAYING BANK"],
                    "Dicari dari": r["AkunBank"] if r["Metode"] == "transfer" else f"CASH - {dept}",
                    "Catatan": "Bank/Departemen tidak ditemukan di tabel konversi No Akun",
                })

    out_df = pd.DataFrame(output_rows, columns=PAYMENT_HEADERS)

    if unmatched_account_rows:
        st.warning(
            f"Ada {len(unmatched_account_rows)} baris yang EXPENSE ACCOUNT NO-nya dikosongkan "
            "karena bank/departemen tidak ditemukan di tabel konversi No Akun."
        )
        st.dataframe(pd.DataFrame(unmatched_account_rows), use_container_width=True, hide_index=True)

    if unknown_number_rows:
        st.warning(
            f"Ada {len(unknown_number_rows)} NUMBER hasil konversi yang tidak ditemukan di daftar "
            "piutang yang sudah ditanam di aplikasi (bisa jadi invoice baru — cek manual dulu)."
        )
        st.dataframe(pd.DataFrame(unknown_number_rows), use_container_width=True, hide_index=True)

    if not unmatched_account_rows and not unknown_number_rows:
        st.success("Semua baris berhasil dipetakan tanpa masalah. 🎉")

    st.dataframe(out_df.fillna(""), use_container_width=True)

    buf = to_excel_bytes(out_df, "Payment")
    st.download_button(
        "⬇️ Download File Template Payment (.xlsx)",
        data=buf,
        file_name=f"Template_Payment_Piutang_ISMA_{dept}_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
