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
# Tabel konversi No Akun, sesuai gambar "No Akun / Nama Kas & Bank di ISMA /
# Nama Di Accurate / Departemen / Cabang".
# BANK_ACCOUNT_TABLE dipakai untuk transaksi TRANSFER (dicocokkan dengan
# nama bank yang muncul di ITEM:ITEM NOTES, format "Bank: <nama> - <rek>").
# CASH_ACCOUNT_MAP dipakai untuk transaksi CASH (dicocokkan berdasarkan
# ITEM:DEPT NAME).
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

# Daftar unit/departemen (sama seperti pilihan customer di menu ISMA), dipakai
# supaya user pilih unit dulu sebelum upload — konsisten dengan menu ISMA.
PAYMENT_UNITS = [
    {"id": "Siswa-KC01", "nama": "Siswa-KC01", "dept": "KC01", "cabang": "Kantor Pusat"},
    {"id": "Siswa-SD01", "nama": "Siswa-SD01", "dept": "SD01", "cabang": "Kantor Pusat"},
    {"id": "Siswa-SPQ1", "nama": "Siswa-SPQ1", "dept": "SPQ1", "cabang": "Kantor Pusat"},
    {"id": "Siswa-TD01", "nama": "Siswa-TD01", "dept": "TD01", "cabang": "Kantor Pusat"},
    {"id": "Siswa-TK01", "nama": "Siswa-TK01", "dept": "TK01", "cabang": "Kantor Pusat"},
    {"id": "Siswa-TK02", "nama": "Siswa-TK02", "dept": "TK02", "cabang": "Kantor Pusat"},
]

UNMATCHED_LABEL = "‼️ TIDAK DITEMUKAN DI TABEL AKUN"

BANK_NOTE_RE = re.compile(r"Bank:\s*(.+?)(?:;|$)")


def parse_expense_info(item_notes: str, dept: str):
    """
    Dari ITEM:ITEM NOTES tentukan:
    - paying_bank: "Transfer" atau "Tunai"
    - expense_account: No Akun sesuai tabel konversi, atau None kalau
      tidak ditemukan (harus dicek manual oleh user)
    - match_key: string yang dipakai untuk pencarian akun (nama bank
      atau nama departemen), dipakai untuk warning kalau tidak ketemu.
    """
    notes = str(item_notes) if item_notes else ""
    m = BANK_NOTE_RE.search(notes)
    if m:
        bank_name = m.group(1).strip()
        akun = BANK_ACCOUNT_MAP.get(normalize(bank_name))
        return "Transfer", akun, bank_name
    else:
        # Dianggap Cash (termasuk kalau notes cuma "Cash" atau kosong)
        akun = CASH_ACCOUNT_MAP.get(normalize(dept)) if dept else None
        return "Tunai", akun, f"CASH - {dept}"


def render_payment_menu():
    st.title("💳 Konversi Template Payment ISMA")
    st.caption(
        "Upload file hasil generate **Template Accurate ISMA** (dari menu ISMA, "
        "sheet 'Template'), lalu unduh file Template Payment siap-import ke Accurate."
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
    selected_unit_nama = st.selectbox("Nama Pelanggan", unit_nama_list, key="payment_unit")
    selected_unit = next(u for u in PAYMENT_UNITS if u["nama"] == selected_unit_nama)
    selected_dept = selected_unit["dept"]

    col1, col2, col3 = st.columns(3)
    col1.metric("Customer No", selected_unit["id"])
    col2.metric("Departemen", selected_dept)
    col3.metric("Cabang", selected_unit["cabang"])

    st.header("2. Upload File Template ISMA (.xlsx)")
    uploaded = st.file_uploader(
        "Pilih file hasil generate Template Accurate ISMA (untuk unit yang dipilih di atas)",
        type=["xlsx", "xls"], key="payment_upload",
    )

    if uploaded is None:
        st.info("Silakan upload file Template ISMA (.xlsx) untuk melanjutkan.")
        return

    try:
        try:
            raw_df = pd.read_excel(uploaded, sheet_name="Template")
        except ValueError:
            # sheet "Template" tidak ada -> pakai sheet pertama
            raw_df = pd.read_excel(uploaded, sheet_name=0)
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        st.stop()

    st.write("Preview data mentah:")
    st.dataframe(raw_df.head(20), use_container_width=True)

    col_customer = find_column(raw_df, ["CUSTOMER NO"])
    col_number = find_column(raw_df, ["NUMBER"])
    col_branch = find_column(raw_df, ["BRANCH"])
    col_date = find_column(raw_df, ["DATE"])
    col_notes = find_column(raw_df, ["ITEM:ITEM NOTES"])
    col_desc = find_column(raw_df, ["DESCRIPTION"])
    col_unitprice = find_column(raw_df, ["ITEM:UNITPRICE"])
    col_dept = find_column(raw_df, ["ITEM:DEPT NAME"])

    missing = [
        name for name, col in [
            ("CUSTOMER NO", col_customer), ("NUMBER", col_number),
            ("BRANCH", col_branch), ("DATE", col_date),
            ("ITEM:ITEM NOTES", col_notes), ("ITEM:UNITPRICE", col_unitprice),
            ("ITEM:DEPT NAME", col_dept),
        ] if col is None
    ]
    if missing:
        st.error(
            "Kolom berikut tidak ditemukan di file input: " + ", ".join(missing)
            + ". Pastikan file yang di-upload adalah hasil generate Template ISMA."
        )
        st.stop()

    df = raw_df.copy()

    # Filter baris supaya hanya yang ITEM:DEPT NAME-nya sesuai unit terpilih
    # yang diproses — konsisten dengan pola "pilih unit dulu" di menu ISMA.
    df["_dept_norm"] = df[col_dept].apply(
        lambda v: normalize(v) if pd.notna(v) else ""
    )
    matching_mask = df["_dept_norm"] == normalize(selected_dept)
    n_skipped = (~matching_mask).sum()
    df_filtered = df[matching_mask].copy()

    if n_skipped > 0:
        st.warning(
            f"{n_skipped} baris di file dilewati karena ITEM:DEPT NAME-nya bukan "
            f"**{selected_dept}** (unit yang dipilih di atas)."
        )

    if df_filtered.empty:
        st.error(
            f"Tidak ada baris dengan ITEM:DEPT NAME = {selected_dept} di file ini. "
            "Pastikan file yang di-upload sesuai dengan unit yang dipilih."
        )
        st.stop()

    output_rows = []
    unmatched_rows = []

    for _, r in df_filtered.iterrows():
        # lewati baris kosong (kalau ada sisa baris blank di excel)
        if pd.isna(r[col_number]) and pd.isna(r[col_customer]):
            continue

        dept = str(r[col_dept]).strip() if col_dept and pd.notna(r[col_dept]) else selected_dept
        notes = r[col_notes] if pd.notna(r[col_notes]) else ""
        paying_bank, expense_account, match_key = parse_expense_info(notes, dept)

        row = {h: None for h in PAYMENT_HEADERS}
        row["CUSTOMER NO"] = r[col_customer]
        row["NUMBER"] = r[col_number]
        row["BRANCH"] = r[col_branch]
        row["DATE"] = r[col_date]
        row["EXPENSE ACCOUNT NO"] = expense_account
        row["DESCRIPTION"] = r[col_desc] if col_desc and pd.notna(r[col_desc]) else None
        row["PAYMENT TOTAL"] = r[col_unitprice]
        row["PAYMENT NUMBER"] = r[col_number]
        row["PAYMENT VALUE"] = r[col_unitprice]
        row["PAYING BANK"] = paying_bank
        output_rows.append(row)

        if expense_account is None:
            unmatched_rows.append({
                "NUMBER": r[col_number],
                "PAYING BANK": paying_bank,
                "Dicari dari": match_key,
                "Catatan": "Bank/Departemen tidak ditemukan di tabel konversi No Akun",
            })

    out_df = pd.DataFrame(output_rows, columns=PAYMENT_HEADERS)

    st.header("3. Hasil Konversi Template Payment")

    if unmatched_rows:
        st.warning(
            f"Ada {len(unmatched_rows)} baris yang EXPENSE ACCOUNT NO-nya dibiarkan kosong "
            "karena bank/departemen tidak ditemukan di tabel konversi. Cek daftar di bawah, "
            "lalu isi manual No Akun-nya sebelum import ke Accurate."
        )
        st.dataframe(pd.DataFrame(unmatched_rows), use_container_width=True, hide_index=True)
    else:
        st.success("Semua baris berhasil dipetakan ke No Akun. 🎉")

    st.dataframe(out_df.fillna(""), use_container_width=True)

    buf = to_excel_bytes(out_df, "Payment")
    st.download_button(
        "⬇️ Download File Template Payment (.xlsx)",
        data=buf,
        file_name=f"Template_Payment_ISMA_{selected_dept}_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )