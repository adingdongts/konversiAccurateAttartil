import io
import re
from datetime import datetime

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Header kolom TEMPLATE Accurate (urutan harus persis)
# ---------------------------------------------------------------------------
TEMPLATE_HEADERS = [
    "CUSTOMER NO", "NUMBER", "BRANCH", "DATE", "TAXABLE", "ADDRESS",
    "TOTAL INCLUDING VAT", "TAX INVOICE NUMBER", "ADVANCE INVOICE",
    "INVOICE DISCOUNT (%)", "INVOICE DISCOUNT (Rp)", "DESCRIPTION", "PO NO",
    "SHIPPING", "SHIPPING DATE", "FOB", "PAYMENT TERMS", "DUE DATE",
    "PAYING BANK", "PAYMENT VALUE",
    "CUSTOM CHARACTER 1", "CUSTOM CHARACTER 2", "CUSTOM CHARACTER 3",
    "CUSTOM CHARACTER 4", "CUSTOM CHARACTER 5", "CUSTOM CHARACTER 6",
    "CUSTOM CHARACTER 7", "CUSTOM CHARACTER 8", "CUSTOM CHARACTER 9",
    "CUSTOM CHARACTER 10",
    "CUSTOM NUMBER 1", "CUSTOM NUMBER 2", "CUSTOM NUMBER 3", "CUSTOM NUMBER 4",
    "CUSTOM NUMBER 5", "CUSTOM NUMBER 6", "CUSTOM NUMBER 7", "CUSTOM NUMBER 8",
    "CUSTOM NUMBER 9", "CUSTOM NUMBER 10",
    "CUSTOM DATE 1", "CUSTOM DATE 2",
    "VA NUMBER", "ACCOUNT RECEIVABLE NUMBER", "PAYMENT WITH UNIQUE CODE",
    "SUB COMPANY CODE",
    "ITEM:ITEM NO", "ITEM:QUANTITY", "ITEM:UNITPRICE", "ITEM:UNIT",
    "ITEM:WAREHOUSE NAME ", "ITEM:NAME",
    "ITEM:ITEM DISCOUNT (%)", "ITEM:ITEM DISCOUNT (RP)", "ITEM:ITEM NOTES",
    "ITEM:SALESMAN ID", "ITEM:DEPT NAME", "ITEM:PROJECT NO",
    "ITEM:CUSTOM CHARACTER 1", "ITEM:CUSTOM CHARACTER 2", "ITEM:CUSTOM CHARACTER 3",
    "ITEM:CUSTOM CHARACTER 4", "ITEM:CUSTOM CHARACTER 5", "ITEM:CUSTOM CHARACTER 6",
    "ITEM:CUSTOM CHARACTER 7", "ITEM:CUSTOM CHARACTER 8", "ITEM:CUSTOM CHARACTER 9",
    "ITEM:CUSTOM CHARACTER 10", "ITEM:CUSTOM CHARACTER 11", "ITEM:CUSTOM CHARACTER 12",
    "ITEM:CUSTOM CHARACTER 13", "ITEM:CUSTOM CHARACTER 14", "ITEM:CUSTOM CHARACTER 15",
    "ITEM:CUSTOM NUMBER 1", "ITEM:CUSTOM NUMBER 2", "ITEM:CUSTOM NUMBER 3",
    "ITEM:CUSTOM NUMBER 4", "ITEM:CUSTOM NUMBER 5", "ITEM:CUSTOM NUMBER 6",
    "ITEM:CUSTOM NUMBER 7", "ITEM:CUSTOM NUMBER 8", "ITEM:CUSTOM NUMBER 9",
    "ITEM:CUSTOM NUMBER 10",
    "ITEM:CUSTOM DATE 1", "ITEM:CUSTOM DATE 2",
    "ITEM:CUSTOM FINANCE CATEGORY 1", "ITEM:CUSTOM FINANCE CATEGORY 2",
    "ITEM:CUSTOM FINANCE CATEGORY 3", "ITEM:CUSTOM FINANCE CATEGORY 4",
    "ITEM:CUSTOM FINANCE CATEGORY 5", "ITEM:CUSTOM FINANCE CATEGORY 6",
    "ITEM:CUSTOM FINANCE CATEGORY 7", "ITEM:CUSTOM FINANCE CATEGORY 8",
    "ITEM:CUSTOM FINANCE CATEGORY 9", "ITEM:CUSTOM FINANCE CATEGORY 10",
    "ITEM:DELIVERY ORDER NO.", "ITEM:SALES ORDER NO",
    "EXPENSE:ACCOUNT NO", "EXPENSE:EXPENSE VALUE", "EXPENSE:EXPENSE NAME",
    "EXPENSE:EXPENSE NOTE", "EXPENSE:DEPT NAME", "EXPENSE:PROJECT NO",
    "EXPENSE:FINANCIAL CATEGORY 1", "EXPENSE:FINANCIAL CATEGORY 2",
    "EXPENSE:FINANCIAL CATEGORY 3", "EXPENSE:FINANCIAL CATEGORY 4",
    "EXPENSE:FINANCIAL CATEGORY 5", "EXPENSE:FINANCIAL CATEGORY 6",
    "EXPENSE:FINANCIAL CATEGORY 7", "EXPENSE:FINANCIAL CATEGORY 8",
    "EXPENSE:FINANCIAL CATEGORY 9", "EXPENSE:FINANCIAL CATEGORY 10",
    "EXPENSE:SALES ORDER NO",
]


def empty_row():
    return {h: None for h in TEMPLATE_HEADERS}


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


# =============================================================================
# Master data ISMA
# =============================================================================

ISMA_CUSTOMERS = [
    {"id": "Siswa-KC01", "nama": "Siswa-KC01", "dept": "KC01", "cabang": "Kantor Pusat"},
    {"id": "Siswa-SD01", "nama": "Siswa-SD01", "dept": "SD01", "cabang": "Kantor Pusat"},
    {"id": "Siswa-SPQ1", "nama": "Siswa-SPQ1", "dept": "SPQ1", "cabang": "Kantor Pusat"},
    {"id": "Siswa-TD01", "nama": "Siswa-TD01", "dept": "TD01", "cabang": "Kantor Pusat"},
    {"id": "Siswa-TK01", "nama": "Siswa-TK01", "dept": "TK01", "cabang": "Kantor Pusat"},
    {"id": "Siswa-TK02", "nama": "Siswa-TK02", "dept": "TK02", "cabang": "Kantor Pusat"},
]

# (Departemen, Nama Produk ISMA) -> Kode Barang di Accurate
ISMA_PRODUCT_MAP = {
    ("TD01", "IP"): "TD01.IP-2026",
    ("TD01", "DKS"): "TD01.DKS-26Sem1",
    ("TD01", "PKBM"): "TD01.PKBM-2026",
    ("TD01", "SPP"): "TD01.SPP-2607",
    ("TD01", "Buku"): "TD01.BK-2026",
    ("TD01", "MRTL"): "TD01.MRTL-2026",
    ("TD01", "DU"): "TD01.DU-2026",
    ("TD01", "DB"): "TD01.DB-2026",
    ("TK01", "IP"): "TK01.IP-2026",
    ("TK01", "DKS"): "TK01.DKS-26Sem1",
    ("TK01", "PKBM"): "TK01.PKBM-2026",
    ("TK01", "SPP"): "TK01.SPP-2607",
    ("TK01", "Buku"): "TK01.BK-2026",
    ("TK01", "DU"): "TK01.DU-2026",
    ("TK01", "DB"): "TK01.DB-2026",
    ("TK02", "IP"): "TK02.IP-2026",
    ("TK02", "DKS"): "TK02.DKS-26Sem1",
    ("TK02", "PKBM"): "TK02.PKBM-2026",
    ("TK02", "SPP"): "TK02.SPP-2607",
    ("TK02", "Buku"): "TK02.BK-2026",
    ("TK02", "DU"): "TK02.DU-2026",
    ("TK02", "DB"): "TK02.DB-2026",
    ("SD01", "IP"): "TK02.IP-2026",
    ("SD01", "DKS"): "TK02.DKS-2026",
    ("SD01", "PKBM"): "TK02.PKBM-2026",
    ("SD01", "SPP"): "TK02.SPP-2607",
    ("SD01", "Buku"): "TK02.BK-2026",
    ("SD01", "DU"): "TK02.DU-2026",
    ("SD01", "DB"): "TK02.DB-2026",
    ("SPQ1", "IP"): "SPQ1.IP-2026",
    ("SPQ1", "SPP"): "SPQ1.SPP-2607",
    ("SPQ1", "DKS"): "SPQ1.DKS-2026",
    ("SPQ1", "PPS"): "SPQ1.PELATIHAN-2026",
    ("SPQ1", "SPA"): "SPQ1.SEWA-2026",
    ("SPQ1", "Buku"): "SPQ1.BK-2026",
    ("SPQ1", "LKS"): "SPQ1.LKS-2026",
    ("SPQ1", "Kesehatan"): "SPQ1.KES-2026",
    ("SPQ1", "IPM"): "SPQ1.IP-2026",
    ("SPQ1", "PF"): "SPQ1.PERFAS-2026",
    ("SPQ1", "DB"): "SPQ1.DB-2026",
    ("SPQ1", "SRGM"): "SPQ1.SRGM-2026",
    ("KC01", "Buku"): "KC01.BK-2026",
    ("KC01", "DKS"): "KC01.DKS-2026",
    ("KC01", "IP"): "KC01.IP-2026",
    ("KC01", "SPP"): "KC01.SPP-2607",
    ("KC01", "DB"): "KC01.DB-2026",
}

# kolom produk yang mungkin muncul di file input Laporan Pembayaran
ISMA_PRODUCT_COLS = ["Buku", "PKBM", "DKS", "SPP", "IP", "DB", "MRTL", "DU"]


def parse_id_tanggal(s):
    """Parse tanggal format DD-MM-YYYY (string) jadi Timestamp."""
    if pd.isna(s) or s in (None, "", "-"):
        return None
    return pd.to_datetime(str(s), format="%d-%m-%Y", errors="coerce")


def render_isma_menu():
    st.title("📚 Konversi Laporan Pembayaran ISMA ")
    st.caption(
        "Upload laporan pembayaran (Excel) untuk 1 departemen/unit, "
        "lalu unduh file siap-import ke Accurate."
    )

    st.header("1. Pilih Customer (Unit/Departemen)")
    nama_list = [c["nama"] for c in ISMA_CUSTOMERS]
    selected_nama = st.selectbox("Nama Pelanggan", nama_list, key="isma_customer")
    selected_customer = next(c for c in ISMA_CUSTOMERS if c["nama"] == selected_nama)
    customer_no = selected_customer["id"]
    dept = selected_customer["dept"]
    cabang = selected_customer["cabang"]

    col1, col2, col3 = st.columns(3)
    col1.metric("Customer No", customer_no)
    col2.metric("Departemen", dept)
    col3.metric("Cabang", cabang)

    st.header("2. Upload File Laporan Pembayaran (.xlsx)")
    uploaded = st.file_uploader(
        "Pilih file Excel laporan pembayaran (format seperti 'Lap Pembayaran')",
        type=["xlsx", "xls"], key="isma_upload",
    )

    if uploaded is None:
        st.info("Silakan upload file laporan pembayaran (.xlsx) untuk melanjutkan.")
        return

    try:
        raw_df = pd.read_excel(uploaded, header=None)
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        st.stop()

    # cari baris header (baris yang mengandung 'No Ref' & 'Siswa')
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
    col_periode = find_column(data, ["Periode"])
    col_bulan = find_column(data, ["Bulan"])
    col_metode = find_column(data, ["Metode"])
    col_tgltransfer = find_column(data, ["Tgl Transfer"])
    col_siswa = find_column(data, ["Siswa"])
    col_nis = find_column(data, ["NIS"])

    product_cols_present = [c for c in ISMA_PRODUCT_COLS if find_column(data, [c])]
    product_col_map = {c: find_column(data, [c]) for c in product_cols_present}

    required_missing = [
        name for name, col in [
            ("No", col_no), ("Tgl", col_tgl), ("Periode", col_periode),
            ("Bulan", col_bulan), ("Metode", col_metode),
        ] if col is None
    ]
    if required_missing:
        st.error("Kolom berikut tidak ditemukan di file input: " + ", ".join(required_missing))
        st.stop()

    # forward-fill kolom yang kosong di baris kontinuasi
    ffill_cols = [
        c for c in [col_no, col_noref, col_tgl, col_periode, col_metode,
                    col_tgltransfer, col_siswa, col_nis]
        if c is not None
    ]
    data[ffill_cols] = data[ffill_cols].ffill()

    # ambil daftar Bulan & Periode unik dari file untuk dropdown filter
    bulan_options = sorted({str(v).strip() for v in data[col_bulan].dropna().tolist() if str(v).strip()})
    periode_options = sorted({str(v).strip() for v in data[col_periode].dropna().tolist() if str(v).strip()})

    if not bulan_options or not periode_options:
        st.error("Tidak ditemukan nilai Bulan atau Periode di file ini.")
        st.stop()

    st.header("3. Pilih Bulan & Periode Tagihan")
    st.caption(
        "Bulan Tagihan hanya dipakai untuk memfilter baris SPP. "
        "Periode Tagihan dipakai untuk memfilter produk selain SPP (Buku, PKBM, DKS, IP, DB, dll)."
    )
    fc1, fc2 = st.columns(2)
    selected_bulan = fc1.selectbox("Bulan Tagihan (untuk SPP)", bulan_options, key="isma_bulan")
    selected_periode = fc2.selectbox(
        "Periode Tagihan (tahun ajaran, untuk selain SPP)", periode_options, key="isma_periode"
    )

    # ----- melt: pecah tiap kolom produk jadi baris transaksi tersendiri -----
    records = []
    for _, r in data.iterrows():
        metode_raw = str(r[col_metode]).strip().lower() if pd.notna(r[col_metode]) else ""
        if metode_raw not in ("cash", "transfer"):
            continue
        for prod, col in product_col_map.items():
            val = r[col]
            if pd.isna(val) or val == 0:
                continue
            if prod == "SPP":
                bulan_val = str(r[col_bulan]).strip() if pd.notna(r[col_bulan]) else None
                if bulan_val != selected_bulan:
                    continue
            else:
                periode_val = str(r[col_periode]).strip() if pd.notna(r[col_periode]) else None
                if periode_val != selected_periode:
                    continue

            # Tanggal SELALU diambil dari kolom "Tgl" (bukan "Tgl Transfer"),
            # untuk cash maupun transfer.
            tgl = parse_id_tanggal(r[col_tgl])
            if tgl is None:
                continue

            siswa_val = str(r[col_siswa]).strip() if col_siswa is not None and pd.notna(r[col_siswa]) else ""
            nis_val = str(r[col_nis]).strip() if col_nis is not None and pd.notna(r[col_nis]) else ""

            records.append({
                "Produk": prod,
                "Nominal": float(val),
                "Metode": metode_raw,
                "Tanggal": tgl,
                "Siswa": siswa_val,
                "NIS": nis_val,
            })

    if not records:
        st.warning(
            "Tidak ada transaksi yang cocok dengan Bulan/Periode yang dipilih. "
            "Coba pilih Bulan atau Periode lain."
        )
        return

    trans_df = pd.DataFrame(records)

    # produk yang tidak ada di mapping departemen ini -> unregister
    unmapped_products = sorted({
        p for p in trans_df["Produk"].unique()
        if (dept, p) not in ISMA_PRODUCT_MAP
    })
    if unmapped_products:
        st.warning(
            f"Produk berikut tidak ada di daftar mapping kode barang untuk departemen {dept}, "
            "akan masuk ke daftar Unregister: " + ", ".join(unmapped_products)
        )

    mapped_df = trans_df[~trans_df["Produk"].isin(unmapped_products)].copy()
    unmapped_df = trans_df[trans_df["Produk"].isin(unmapped_products)].copy()

    # ----- Cash: total per (tanggal, produk), Siswa/NIS hilang krn digabung.
    # ----- Transfer: apa adanya (tidak digabung), Siswa/NIS dipertahankan.
    cash_df = mapped_df[mapped_df["Metode"] == "cash"]
    transfer_df = mapped_df[mapped_df["Metode"] == "transfer"]

    cash_grouped = cash_df.groupby(["Tanggal", "Produk"], as_index=False).agg(Nominal=("Nominal", "sum"))
    cash_grouped["Metode"] = "cash"
    cash_grouped["Siswa"] = ""
    cash_grouped["NIS"] = ""

    final_lines = pd.concat(
        [cash_grouped, transfer_df[["Tanggal", "Produk", "Nominal", "Metode", "Siswa", "NIS"]]],
        ignore_index=True,
    )
    final_lines = final_lines.sort_values(["Tanggal", "Produk"])

    st.header("4. Hasil Konversi")

    output_rows = []
    for (tanggal, produk), sub in final_lines.groupby(["Tanggal", "Produk"]):
        mmYY = tanggal.strftime("%m%y")
        dd = tanggal.strftime("%d")
        base_number = f"ISMA-{produk}-{dept}-{mmYY}-{dd}"
        kode_barang = ISMA_PRODUCT_MAP.get((dept, produk))
        # Tanggal ditulis persis format tanggal di file inputan (dd-mm-yyyy),
        # bukan objek datetime (supaya tidak ikut ke-set jam saat ini).
        date_str = tanggal.strftime("%d-%m-%Y")

        for line_no, (_, r) in enumerate(sub.iterrows(), start=1):
            row = empty_row()
            # NUMBER dikasih suffix .1, .2, dst supaya unik per baris.
            row["CUSTOMER NO"] = customer_no
            row["NUMBER"] = f"{base_number}.{line_no}"
            row["BRANCH"] = cabang
            row["DATE"] = date_str
            row["ITEM:ITEM NO"] = kode_barang
            row["ITEM:QUANTITY"] = 1
            row["ITEM:UNITPRICE"] = r["Nominal"]
            row["ITEM:DEPT NAME"] = dept
            if r["Metode"] == "transfer" and (r["Siswa"] or r["NIS"]):
                row["DESCRIPTION"] = f"{r['Siswa']}_{r['NIS']}"
            output_rows.append(row)

    out_df = pd.DataFrame(output_rows, columns=TEMPLATE_HEADERS)
    st.dataframe(out_df, use_container_width=True)

    unmapped_df = unmapped_df.rename(columns={"Nominal": "Nominal (belum ditotal)"})
    unmapped_df["Departemen"] = dept

    buf_main = to_excel_bytes(out_df, "Template")
    buf_unreg = to_excel_bytes(unmapped_df, "Unregister")

    dl_col1, dl_col2 = st.columns(2)
    with dl_col1:
        st.download_button(
            "⬇️ Download File Template Accurate (.xlsx)",
            data=buf_main,
            file_name=f"Template_Accurate_ISMA_{dept}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with dl_col2:
        if not unmapped_df.empty:
            st.download_button(
                "⬇️ Download Daftar Unregister (.xlsx)",
                data=buf_unreg,
                file_name=f"Unregister_ISMA_{dept}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.success("Semua produk sudah terdaftar di mapping, tidak ada yang perlu dicek.")
