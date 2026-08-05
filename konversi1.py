import io
import re
from datetime import datetime

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Konversi LTQ ke Template Accurate", layout="wide")

# ---------------------------------------------------------------------------
# Master data (hardcoded sesuai spesifikasi)
# ---------------------------------------------------------------------------

CUSTOMERS = [
    {"id": "C.00002", "nama": "Buku Pustaka di LTQ Cibatu", "area": "Cibatu"},
    {"id": "C.00004", "nama": "Buku Pustaka di LTQ Cikembar", "area": "Cikembar"},
    {"id": "C.00003", "nama": "Buku Pustaka di LTQ Citamiang", "area": "Citamiang"},
    {"id": "C.00001", "nama": "Buku Pustaka di LTQ Karamat", "area": "Karamat"},
    {"id": "C.00005", "nama": "Buku Pustaka di LTQ Karang Tengah", "area": "Karang Tengah"},
]

WAREHOUSE_MAP = {
    "Cibatu": "Gudang Pustaka LTQ Cibatu",
    "Cikembar": "Gudang Pustaka LTQ Cikembar",
    "Citamiang": "Gudang Pustaka LTQ Citamiang",
    "Karamat": "Gudang Pustaka LTQ Karamat",
    "Karang Tengah": "Gudang Pustaka LTQ Karang Tengah",
}

# ---------------------------------------------------------------------------
# Tabel konversi Nama Produk SIPP -> Nama Produk Accurate, PER CABANG LTQ
# Sumber: daftar_konversi_nama_produk_pustaka_di_LTQ.xlsx (ditanam langsung
# di kode supaya mapping jelas & konsisten, tidak perlu upload ulang tiap kali)
# Nilai None artinya kolom "Nama Produk di Accurate" memang kosong di tabel
# sumber -> item ini otomatis dianggap belum punya padanan di Accurate.
# ---------------------------------------------------------------------------
CONVERSION_TABLES = {
    "Karamat": [
        ("AYAT GHORIBAH", "Buku Ayat Ghorib"),
        ("AL-MATSUROT SUGHRO", "AL-MATSUROT SUGHRO"),
        ("Buku At Tartil", None),
        ("BUKU BAHASA ARAB - 2025", "BUKU BAHASA ARAB"),
        ("BUKU MENULIS HIJAIYAH", "AL-MATSUROT SUGHRO"),
        ("BUKU MENULIS JUZ 30", "BUKU HIJAIYAH"),
        ("FASILITAS KELAS ANDALUSIA", None),
        ("FASILITAS KELAS MEKAH", None),
        ("FASILITAS KELAS MADINAH", None),
        ("FASILITAS KELAS PALESTINA", None),
        ("FASILITAS KELAS PEMULA", None),
        ("FASILITAS KELAS TALAQQI", None),
        ("FASILITAS KELAS TAHSIN TEORI", None),
        ("IQRA ANAK 6 JILID", "IQRA ANAK 6 JILID"),
        ("IQRA ANAK SATUAN", "IQRO SATUAN"),
        ("IQRA DEWASA", "IQRO DEWASA (HVS PUTIH WARNA)"),
        ("IQRA DEWASA 2025", None),
        ("IQRO DEWASA KERTAS BURAM", "IQRO DEWASA KERTAS BURAM"),
        ("IQRA KLASIKAL", "IQRO KLASIK"),
        ("JUZ AMMA 28-30", "JUZ AMMA (28-30)"),
        ("JUZ AMMA 1 - 2", "JUZ AMMA (1-2)"),
        ("JUZ AMMA 26-30", "JUZ AMMA (26-30)"),
        ("KARTU INFAK", "Kartu Infak"),
        ("LEKKER KECIL", "MEJA LEKAR"),
        ("MUTABAAH ANAK", "MUTABAAH ANAK"),
        ("MUSHAF AT TARTIL B5", "MUSHAF ATTARTIL B5"),
        ("MATB5/Wakaf", None),
        ("MUTABAAH DEWASA", "MUTABAAH DEWASA"),
        ("MUSHAF HAFALAH KECIL", "MUSHAF HAFALAN KECIL"),
        ("METODE KUNCI", "METODE KUNCI"),
        ("MUSHAF SLETING KECIL", "MUSHAF SLETING KECIL"),
        ("RAPOT ANAK", "RAPOT ANAK"),
        ("RAPOT DEWASA", "RAPOT DEWASA"),
        ("RISALAH", "RISALAH"),
        ("TAS AT TARTIL", "TAS AT TARTIL"),
    ],
    "Cibatu": [
        ("AYAT GHORIBAH", "BUKU AYAT GHORIB"),
        ("AL-MATSUROT SUGHRO", "AL-MATSUROT SUGHRO"),
        ("BUKU BAHASA ARAB 2", "BUKU BAHASA ARAB"),
        ("Buku Tulis", None),
        ("BUKU MENULIS HIJAIYAH", "AL-MATSUROT SUGHRO"),
        ("BUKU MENULIS JUZ 30", "BUKU HIJAIYAH"),
        ("FASILITAS KELAS ANDALUSIA", None),
        ("FASILITAS KELAS MEKAH", None),
        ("FASILITAS KELAS MADINAH", None),
        ("FASILITAS KELAS PALESTINA", None),
        ("FASILITAS KELAS PEMULA", None),
        ("FASILITAS KELAS TALAQQI", None),
        ("FASILITAS KELAS TAHSIN TEORI", None),
        ("IQRA ANAK 6 JILID", "IQRA ANAK 6 JILID"),
        ("IQRA DEWASA", "IQRO DEWASA (HVS PUTIH WARNA)"),
        ("IQRA DEWASA 2", None),
        ("IQRA KLASIKAL", "IQRO KLASIK"),
        ("Iqro Klasikal", None),
        ("IQRA ANAK SATUAN", "IQRO SATUAN"),
        ("JUZ AMMA 28-30", "JUZ AMMA (28-30)"),
        ("JUZ AMMA 26-30", "JUZ AMMA (26-30)"),
        ("KARTU INFAK", "Kartu Infak"),
        ("MUTABAAH ANAK", "MUTABAAH ANAK"),
        ("MUSHAF AT TARTIL B5", "MUSHAF ATTARTIL B5"),
        ("MUTABAAH DEWASA", "MUTABAAH DEWASA"),
        ("MUSHAF HAFALAN KECIL (tosca)", "MUSHAF HAFALAN KECIL"),
        ("METODE KUNCI", "METODE KUNCI"),
        ("MUSHAF SLETING", "MUSHAF BESAR SLETING BARU"),
        ("Mushaf Waqaf & ibtida", None),
        ("Mushaf Hafalan Besar", "MUSHAF ATTARTIL B5"),
        ("Mushaf Hafalan Kecil", "MUSHAF HAFALAN KECIL"),
        ("RAPOT ANAK", "RAPOT ANAK"),
        ("RAPOT DEWASA", "RAPOT DEWASA"),
        ("RISALAH", "RISALAH"),
        ("Buku Risalah", "RISALAH"),
        ("Risalah", "RISALAH"),
        ("TAS AT-TARTIL", "TAS AT TARTIL"),
    ],
    "Citamiang": [
        ("AYAT GHORIBAH", "BUKU AYAT GHORIB"),
        ("Ayat Ghorib Baru", "BUKU AYAT GHORIB"),
        ("Al Matsurot", "AL-MATSUROT SUGHRO"),
        ("AL-MATSURAT SUGHRO", "AL-MATSUROT SUGHRO"),
        ("BUKU BAHASA ARAB 2", "BUKU BAHASA ARAB"),
        ("Buku Bahasa Arab BAru", "BUKU BAHASA ARAB"),
        ("Buku Attibyan", None),
        ("Buku Gambar", None),
        ("Buku Menulis Hijaiyah", "BUKU HIJAIYAH"),
        ("Buku Rumah Tahfidz", None),
        ("Buku Tulis", None),
        ("BUKU MENULIS JUZ 30", "BUKU MENULIS JUZ 30"),
        ("BUKU MENULIS HIJAIYAH", "BUKU HIJAIYAH"),
        ("CD Murottal", None),
        ("FASILITAS KELAS ANDALUSIA", None),
        ("FASILITAS KELAS MEKAH", None),
        ("FASILITAS KELAS MADINAH", None),
        ("FASILITAS KELAS PALESTINA", None),
        ("FASILITAS KELAS PEMULA", None),
        ("FASILITAS KELAS TALAQQI", None),
        ("FASILITAS KELAS TAHSIN TEORI", None),
        ("Hadits Arbain", None),
        ("IQRO ANAK 6 JILID", "IQRA ANAK 6 JILID"),
        ("IQRO DEWASA", "IQRO DEWASA (HVS PUTIH WARNA)"),
        ("IQRO DEWASA BURAM", "IQRO DEWASA KERTAS BURAM"),
        ("IQRO KLASIKAL", "IQRO KLASIK"),
        ("Iqro Anak", "IQRA ANAK 6 JILID"),
        ("Iqro Dewasa", "IQRO DEWASA (HVS PUTIH WARNA)"),
        ("Iqro Satuan", "IQRO SATUAN"),
        ("Iqro Klasikal", "IQRO KLASIK"),
        ("IQRO SATUAN", "IQRO SATUAN"),
        ("JUZ AMMA 30", "Juz Amma 30"),
        ("JUZ-AMMA (28-30)", "JUZ AMMA (28-30)"),
        ("Juz 28 - 30", "JUZ AMMA (28-30)"),
        ("JUZ-AMMA(26-30)", "JUZ AMMA (28-30)"),
        ("KARTU INFAK", "Kartu Infak"),
        ("Kaos Kaki", None),
        ("Kaos Kaki Baru", None),
        ("Kartu SPP Baru", "Kartu Infak"),
        ("LEKKER KECIL", "MEJA LEKAR"),
        ("MUTABAAH ANAK", "MUTABAAH ANAK"),
        ("MUSHAF AT TARTIL B5", "MUSHAF ATTARTIL B5"),
        ("Mushaf At Tartil B5", "MUSHAF ATTARTIL B5"),
        ("MUTABAAH DEWASA", "MUTABAAH DEWASA"),
        ("MUSHAF HAFALAN KECIL", "MUSHAF HAFALAN KECIL"),
        ("METODE KUNCI", "METODE KUNCI"),
        ("Metode Kunci", "METODE KUNCI"),
        ("Manset", None),
        ("Manset Panjang", None),
        ("MUSHAF SLETING", "MUSHAF BESAR SLETING BARU"),
        ("Mushaf Besar", "MUSHAF ATTARTIL B5"),
        ("Mushaf Kecil", "MUSHAF HAFALAN KECIL"),
        ("Mushaf Sleting", "MUSHAF HAFALAN KECIL"),
        ("Mushaf sleting Besar", "MUSHAF BESAR SLETING BARU"),
        ("Mushaf Waqaf & ibtida", None),
        ("Mushaf Hafalan Besar", "MUSHAF ATTARTIL B5"),
        ("Mushaf Hafalan Kecil", "MUSHAF HAFALAN KECIL"),
        ("MUSHAF SLETING BESAR", "MUSHAF BESAR SLETING BARU"),
        ("Mutabaah Dewasa", "MUTABAAH DEWASA"),
        ("Mutabaah Anak", "MUTABAAH ANAK"),
        ("Penghapus", None),
        ("Pin Besar", None),
        ("Pin Kecil", None),
        ("RAPOT ANAK", "RAPOT ANAK"),
        ("RAPOT DEWASA", "RAPOT DEWASA"),
        ("RISALAH 3", "RISALAH"),
        ("Raport Dewasa", "RAPOT DEWASA"),
        ("Raport Anak", "RAPOT ANAK"),
        ("RISALAH", "RISALAH"),
        ("Buku Risalah", "RISALAH"),
        ("Risalah", "RISALAH"),
        ("TAS AT-TARTIL", "TAS AT TARTIL"),
        ("Tas At-Tartil", "TAS AT TARTIL"),
    ],
    "Cikembar": [
        ("AYAT GHORIBAH", "BUKU AYAT GHORIB"),
        ("AL-MATSUROT SUGHRO", "AL-MATSUROT SUGHRO"),
        ("BUKU MENULIS HIJAIYAH", "BUKU HIJAIYAH"),
        ("BUKU MENULIS JUZ 30", "BUKU MENULIS JUZ 30"),
        ("FASILITAS KELAS ANDALUSIA", None),
        ("FASILITAS KELAS MEKAH", None),
        ("FASILITAS KELAS MADINAH", None),
        ("FASILITAS KELAS PALESTINA", None),
        ("FASILITAS KELAS PEMULA", None),
        ("FASILITAS KELAS TALAQQI", None),
        ("FASILITAS KELAS TAHSIN TEORI", None),
        ("IQRA ANAK 6 JILID", "IQRA ANAK 6 JILID"),
        ("IQRA ANAK SATUAN", "IQRO SATUAN"),
        ("IQRA DEWASA", "IQRO DEWASA (HVS PUTIH WARNA)"),
        ("IQRA KLASIKAL", "IQRO KLASIK"),
        ("IQRA DEWASA KERTAS BURAM", "IQRO DEWASA KERTAS BURAM"),
        ("JUZ AMMA 28-30", "JUZ AMMA (28-30)"),
        ("JUZ AMMA 26 - 30", "JUZ AMMA (26-30)"),
        ("KARTU INFAK", "Kartu Infak"),
        ("LEKKER KECIL", "MEJA LEKAR"),
        ("MUTABAAH ANAK", "MUTABAAH ANAK"),
        ("MUSHAF AT TARTIL B5", "MUSHAF ATTARTIL B5"),
        ("MUTABAAH DEWASA", "MUTABAAH DEWASA"),
        ("MUSHAF HAFALAH KECIL", "MUSHAF HAFALAN KECIL"),
        ("METODE KUNCI", "METODE KUNCI"),
        ("MUSHAF SLETING KECIL", "MUSHAF SLETING KECIL"),
        ("RAPOT ANAK", "RAPOT ANAK"),
        ("RAPOT DEWASA", "RAPOT DEWASA"),
        ("RISALAH", "RISALAH"),
        ("TAS AT TARTIL", "TAS AT TARTIL"),
    ],
    "Karang Tengah": [
        ("AYAT GHORIBAH", "BUKU AYAT GHORIB"),
        ("Al Matsurot", "AL-MATSUROT SUGHRO"),
        ("AL-MATSURAT SUGHRO", "AL-MATSUROT SUGHRO"),
        ("Buku Ilmu Tajwid Bergambar", "Buku Ilmu Tajwid DR. Aiman"),
        ("BUKU BAHASA ARAB 1", "BUKU BAHASA ARAB"),
        ("BUKU BAHASA ARAB 2", "BUKU BAHASA ARAB"),
        ("Buku Bahasa Arab BAru", "BUKU BAHASA ARAB"),
        ("Buku Attibyan", None),
        ("Buku Gambar", "BUKU ABU HAWARIY"),
        ("Buku Rumah Tahfidz", None),
        ("Buku Tulis", None),
        ("Buku Tabungan", "Kartu Infak"),
        ("BUKU MENULIS JUZ 30", "BUKU MENULIS JUZ 30"),
        ("BUKU MENULIS HIJAIYAH", "BUKU HIJAIYAH"),
        ("BUKU MEWARNAI", None),
        ("CD Murottal", None),
        ("FASILITAS KELAS ANDALUSIA", None),
        ("FASILITAS KELAS MEKAH", None),
        ("FASILITAS KELAS MADINAH", None),
        ("FASILITAS KELAS PALESTINA", None),
        ("FASILITAS KELAS PEMULA", None),
        ("FASILITAS KELAS TALAQQI", None),
        ("FASILITAS KELAS TAHSIN TEORI", None),
        ("Hadits Arbain", None),
        ("IQRO KLASIKAL", "IQRO KLASIK"),
        ("Iqro Satuan", "IQRO SATUAN"),
        ("Iqro Klasikal", "IQRO KLASIK"),
        ("Iqro Anak 6 Jilid Terbaru", "IQRA ANAK 6 JILID"),
        ("IQRO DEWASA (BURAM)", "IQRO DEWASA KERTAS BURAM"),
        ("Iqro Dewasa Baru", "IQRO DEWASA (HVS PUTIH WARNA)"),
        ("iqro paket harga lama", "IQRA ANAK 6 JILID"),
        ("IQRO SATUAN", "IQRO SATUAN"),
        ("Juz Amma 26-30", "JUZ AMMA (26-30)"),
        ("Juz Amma 28-30", "JUZ AMMA (28-30)"),
        ("KARTU INFAK", "Kartu Infak"),
        ("LEKKER KECIL", "MEJA LEKAR"),
        ("MUTABAAH ANAK", "MUTABAAH ANAK"),
        ("MUSHAF AT TARTIL B5", "MUSHAF ATTARTIL B5"),
        ("MUTABAAH DEWASA", "MUTABAAH DEWASA"),
        ("MUSHAF HAFALAN KECIL", "MUSHAF HAFALAN KECIL"),
        ("METODE KUNCI", "METODE KUNCI"),
        ("Mushaf Sleting", "MUSHAF BESAR SLETING BARU"),
        ("Mushaf Besar", "MUSHAF ATTARTIL B5"),
        ("Mushaf Kecil", "MUSHAF SLETING KECIL"),
        ("Mushaf sleting Besar", "MUSHAF BESAR SLETING BARU"),
        ("Mushaf Waqaf & ibtida", None),
        ("Mushaf Hafalan Besar", "MUSHAF ATTARTIL B5"),
        ("Mushaf Hafalan Kecil", "MUSHAF HAFALAN KECIL"),
        ("Mushaf Besar Baru", "MUSHAF BESAR SLETING BARU"),
        ("Penghapus", None),
        ("Pin Besar", None),
        ("Pin Kecil", None),
        ("RAPOT ANAK", "RAPOT ANAK"),
        ("RAPOT DEWASA", "RAPOT DEWASA"),
        ("RISALAH", "RISALAH"),
        ("Risalah", "RISALAH"),
        ("TAS AT-TARTIL", "TAS AT TARTIL"),
        ("Pustaka Ekraf", None),
    ],
}

UNREGISTERED_LABEL = "‼️ TIDAK TERDAFTAR (perlu didaftarkan manual)"




# ---------------------------------------------------------------------------
# Header lengkap sesuai template Accurate (urutan HARUS persis seperti ini).
# Kolom yang memang diisi datanya cuma: CUSTOMER NO, NUMBER, BRANCH, DATE,
# ITEM:ITEM NO, ITEM:QUANTITY, ITEM:UNITPRICE, ITEM:UNIT,
# ITEM:WAREHOUSE NAME  -> sisanya sengaja dikosongkan.
# ---------------------------------------------------------------------------
OUTPUT_HEADERS = [
    "CUSTOMER NO", "NUMBER", "BRANCH", "DATE", "TAXABLE", "ADDRESS",
    "TOTAL INCLUDING VAT", "TAX INVOICE NUMBER", "ADVANCE INVOICE",
    "INVOICE DISCOUNT (%)", "INVOICE DISCOUNT (Rp)", "DESCRIPTION", "PO NO",
    "SHIPPING", "SHIPPING DATE", "FOB", "PAYMENT TERMS", "DUE DATE",
    "PAYING BANK", "PAYMENT VALUE",
    "CUSTOM CHARACTER 1", "CUSTOM CHARACTER 2", "CUSTOM CHARACTER 3",
    "CUSTOM CHARACTER 4", "CUSTOM CHARACTER 5", "CUSTOM CHARACTER 6",
    "CUSTOM CHARACTER 7", "CUSTOM CHARACTER 8", "CUSTOM CHARACTER 9",
    "CUSTOM CHARACTER 10",
    "CUSTOM NUMBER 1", "CUSTOM NUMBER 2", "CUSTOM NUMBER 3",
    "CUSTOM NUMBER 4", "CUSTOM NUMBER 5", "CUSTOM NUMBER 6",
    "CUSTOM NUMBER 7", "CUSTOM NUMBER 8", "CUSTOM NUMBER 9",
    "CUSTOM NUMBER 10",
    "CUSTOM DATE 1", "CUSTOM DATE 2",
    "VA NUMBER", "ACCOUNT RECEIVABLE NUMBER", "PAYMENT WITH UNIQUE CODE",
    "SUB COMPANY CODE",
    "ITEM:ITEM NO", "ITEM:QUANTITY", "ITEM:UNITPRICE", "ITEM:UNIT",
    "ITEM:WAREHOUSE NAME ",
    "ITEM:NAME", "ITEM:ITEM DISCOUNT (%)", "ITEM:ITEM DISCOUNT (RP)",
    "ITEM:ITEM NOTES", "ITEM:SALESMAN ID", "ITEM:DEPT NAME",
    "ITEM:PROJECT NO",
    "ITEM:CUSTOM CHARACTER 1", "ITEM:CUSTOM CHARACTER 2",
    "ITEM:CUSTOM CHARACTER 3", "ITEM:CUSTOM CHARACTER 4",
    "ITEM:CUSTOM CHARACTER 5", "ITEM:CUSTOM CHARACTER 6",
    "ITEM:CUSTOM CHARACTER 7", "ITEM:CUSTOM CHARACTER 8",
    "ITEM:CUSTOM CHARACTER 9", "ITEM:CUSTOM CHARACTER 10",
    "ITEM:CUSTOM CHARACTER 11", "ITEM:CUSTOM CHARACTER 12",
    "ITEM:CUSTOM CHARACTER 13", "ITEM:CUSTOM CHARACTER 14",
    "ITEM:CUSTOM CHARACTER 15",
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
]

UNREGISTERED_HEADERS = [
    "Tanggal", "Nama Item", "Qty", "Harga Satuan", "Catatan",
]


def normalize(s: str) -> str:
    s = str(s).upper()
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

# ---------------------------------------------------------------------------
# Mapping Nama Produk Accurate -> Kode Produk Accurate
# Sumber: sheet "Produk Accurate" di file mapingan_kode_buat_template.xlsx
# Key di-normalize supaya matching tidak peduli huruf besar/kecil.
# ---------------------------------------------------------------------------
PRODUCT_CODE_TABLE = [
    ("AL-MATSUROT SUGHRO", "AL-MATSUROT SUGHRO"),
    ("Almatsurat Sughro", "100165"),
    ("BACALAH", "BACALAH"),
    ("BUKTI PENERIMAAN KAS KECIL", "BUKTI PENERIMAAN KAS KECIL"),
    ("BUKTI PENGELUARAN KAS", "BUKTI PENGELUARAN KAS"),
    ("BUKU ABU HAWARIY", "BUKU ABU HAWARIY"),
    ("BUKU BAHASA ARAB", "BUKU BAHASA ARAB"),
    ("BUKU DIROSAH TKIT", "BUKU DIROSAH TKIT"),
    ("BUKU KETIKA PENGHAFAL JATUH CINTA", "BUKU KETIKA PENGHAFAL JATUH CINTA"),
    ("BUKU MENULIS JUZ 30", "BUKU MENULIS JUZ 30"),
    ("BUKU PINTAR MEMBACA", "BUKU PINTAR MEMBACA"),
    ("BUKU PINTAR TAUD", "BUKU PINTAR TAUD"),
    ("Buku Kelompok A ( Hijaiyah, Angka, Abjad, Garis)", "100172"),
    ("Buku Kelompok B ( Hijaiyah, Abjad )", "100173"),
    ("Buku Paket", "100171"),
    ("Buku Panduan Doa dan Hadis", "100175"),
    ("Buku Pintar", "100174"),
    ("Buku Prestasi", "100176"),
    ("CEPAT MEMBACA", "CEPAT MEMBACA"),
    ("HADITS DOA SD BAGIAN 2", "HADITS DOA SD BAGIAN 2"),
    ("MEJA LEKAR", "MEJA LEKAR"),
    ("MUTABAAH KC", "MUTABAAH KC"),
    ("MUTABAAH TAUD", "MUTABAAH TAUD"),
    ("SAMPUL RAPOT TK1", "SAMPUL RAPOT TK1"),
    ("SAMPUL RAPOT TK2", "SAMPUL RAPOT TK2"),
    ("TAS AT TARTIL", "TAS AT TARTIL"),
    ("Buku Ilmu Tajwid DR. Aiman", "Buku Ilmu Tajwid DR. Aiman"),
    ("IQRA ANAK 6 JILID", "IQRA ANAK 6 JILID"),
    ("IQRO DEWASA (HVS PUTIH WARNA)", "IQRO DEWASA (HVS PUTIH WARNA)"),
    ("IQRO DEWASA KERTAS BURAM", "IQRO DEWASA KERTAS BURAM"),
    ("IQRO KLASIK", "IQRO KLASIK"),
    ("IQRO SATUAN", "IQRO SATUAN"),
    ("JUZ AMMA (1-2)", "JUZ AMMA (1-2)"),
    ("JUZ AMMA (26-30)", "JUZ AMMA (26-30)"),
    ("JUZ AMMA (28-30)", "JUZ AMMA (28-30)"),
    ("METODE KUNCI", "METODE KUNCI"),
    ("MUSHAF ATTARTIL B5", "MUSHAF ATTARTIL B5"),
    ("MUSHAF BESAR SLETING BARU", "MUSHAF BESAR SLETING BARU"),
    ("MUSHAF HAFALAN KECIL", "MUSHAF HAFALAN KECIL"),
    ("MUSHAF SLETING KECIL", "MUSHAF SLETING KECIL"),
    ("MUTABAAH ANAK", "MUTABAAH ANAK"),
    ("MUTABAAH DEWASA", "MUTABAAH DEWASA"),
    ("RAPOT ANAK", "RAPOT ANAK"),
    ("RAPOT DEWASA", "RAPOT DEWASA"),
    ("RISALAH", "RISALAH"),
    ("Buku Ayat Ghorib", "100094"),
    ("Buku Hijaiyah", "100045"),
    ("Kartu Infak", "100101"),
    ("Juz amma 30", "Juz amma 30"),
]

PRODUCT_CODE_MAP = {normalize(nama): kode for nama, kode in PRODUCT_CODE_TABLE}

def build_lookup(area: str) -> dict:
    """
    Bangun dict lookup dari CONVERSION_TABLES untuk 1 cabang:
    {normalize(nama_sipp): {"sipp": nama_sipp_asli, "accurate": nama_accurate_atau_None}}
    """
    pairs = CONVERSION_TABLES.get(area, [])
    lookup = {}
    for sipp, acc in pairs:
        lookup[normalize(sipp)] = {"sipp": sipp, "accurate": acc}
    return lookup


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


st.title("📚 Konversi Transaksi Accurate Attartil")
st.caption(
    "Upload laporan penjualan fasilitas belajar (Excel) untuk 1 cabang LTQ, "
    "lalu unduh file siap-import ke Accurate."
)

# ---------------------------------------------------------------------------
# Step 1: pilih LTQ
# ---------------------------------------------------------------------------
st.header("1. Pilih Cabang LTQ")
nama_list = [c["nama"] for c in CUSTOMERS]
selected_nama = st.selectbox("Nama Customer", nama_list)
selected_customer = next(c for c in CUSTOMERS if c["nama"] == selected_nama)
customer_no = selected_customer["id"]
area = selected_customer["area"]
warehouse_name = WAREHOUSE_MAP[area]
number_label = f"LTQ {area.lower()}"

col1, col2 = st.columns(2)
col1.metric("Customer No", customer_no)
col2.metric("Gudang", warehouse_name)

lookup = build_lookup(area)
n_terdaftar = sum(1 for v in lookup.values() if v["accurate"])
n_kosong = len(lookup) - n_terdaftar
st.info(
    f"Tabel konversi cabang **{area}** dimuat: {len(lookup)} nama produk SIPP "
    f"({n_terdaftar} sudah punya padanan Accurate, {n_kosong} masih kosong)."
)

with st.expander(f"Lihat tabel konversi lengkap — {area}"):
    preview_df = pd.DataFrame(
        [(sipp, acc if acc else "(kosong)") for sipp, acc in CONVERSION_TABLES.get(area, [])],
        columns=["Nama Produk di SIPP", "Nama Produk di Accurate"],
    )
    st.dataframe(preview_df, use_container_width=True, hide_index=True)

# ---------------------------------------------------------------------------
# Step 2: upload file laporan penjualan
# ---------------------------------------------------------------------------
st.header("2. Upload File Laporan Penjualan (.xlsx)")
uploaded = st.file_uploader("Pilih file Excel laporan penjualan", type=["xlsx", "xls"])

if uploaded is not None:
    try:
        raw_df = pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        st.stop()

    st.write("Preview data mentah:")
    st.dataframe(raw_df.head(20), use_container_width=True)

    col_tanggal = find_column(raw_df, ["Tanggal", "No Transaksi"])
    col_item = find_column(raw_df, ["Nama Item"])
    col_qty = find_column(raw_df, ["Qty"])
    col_harga = find_column(raw_df, ["Harga Satuan"])

    missing = [
        name for name, col in [
            ("Tanggal / No Transaksi", col_tanggal), ("Nama Item", col_item),
            ("Qty", col_qty), ("Harga Satuan", col_harga),
        ] if col is None
    ]
    if missing:
        st.error(
            "Kolom berikut tidak ditemukan di file input: "
            + ", ".join(missing)
            + ". Pastikan nama kolom sesuai laporan (Tanggal, Nama Item, Qty, Harga Satuan)."
        )
        st.stop()

    df = raw_df[[col_tanggal, col_item, col_qty, col_harga]].copy()
    df.columns = ["Tanggal", "NamaItem", "Qty", "HargaSatuan"]

    def parse_qty(v):
        if pd.isna(v):
            return 0
        s = str(v)
        m = re.search(r"[\d.,]+", s)
        if not m:
            return 0
        return float(m.group(0).replace(",", "").replace(".", "") if "," in s else m.group(0))

    def parse_harga(v):
        if pd.isna(v):
            return 0
        s = str(v)
        s = re.sub(r"[^\d]", "", s)
        return float(s) if s else 0

    df["Qty"] = df["Qty"].apply(parse_qty)
    df["HargaSatuan"] = df["HargaSatuan"].apply(parse_harga)
    df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce", dayfirst=True)
    df = df.dropna(subset=["Tanggal"])
    df["NamaItem"] = df["NamaItem"].astype(str).str.strip()

    grouped = (
        df.groupby(["Tanggal", "NamaItem"], as_index=False)
        .agg(Qty=("Qty", "sum"), HargaSatuan=("HargaSatuan", "max"))
        .sort_values(["Tanggal", "NamaItem"])
    )

    # -----------------------------------------------------------------
    # Step 3: Cocokkan Nama Item (SIPP) -> Nama Accurate, pakai tabel cabang
    # -----------------------------------------------------------------
    st.header("3. Cocokkan Nama Item ➜ Nama Produk Accurate")
    st.caption(
        f"Sistem mencocokkan nama item di laporan dengan tabel konversi cabang **{area}** "
        "yang sudah ditanam di aplikasi. Item yang di tabel konversi belum punya "
        "'Nama Produk di Accurate' (kosong), atau tidak ditemukan sama sekali di "
        "tabel konversi, otomatis ditandai TIDAK TERDAFTAR dan masuk ke sheet Unregistered."
    )

    unique_items = sorted(grouped["NamaItem"].unique())

    guesses = {}
    guess_notes = {}
    for item in unique_items:
        key = normalize(item)
        entry = lookup.get(key)
        if entry is None:
            guesses[item] = None
            guess_notes[item] = f"Nama item tidak ditemukan di tabel konversi cabang {area}"
            continue
        accurate_name = entry["accurate"]
        if not accurate_name:
            guesses[item] = None
            guess_notes[item] = f"Belum ada 'Nama Produk di Accurate' di tabel konversi cabang {area}"
            continue
        guesses[item] = accurate_name
        guess_notes[item] = f"Otomatis via tabel konversi {area}: {entry['sipp']} → {accurate_name}"

    # opsi selectbox: semua Nama Produk Accurate unik yang muncul di tabel cabang ini
    accurate_options = sorted({acc for _, acc in CONVERSION_TABLES.get(area, []) if acc})
    mapping_options = [UNREGISTERED_LABEL] + accurate_options

    mapping_df = pd.DataFrame({
        "Nama Item (dari file)": unique_items,
        "Nama Produk Accurate": [guesses[item] if guesses[item] else UNREGISTERED_LABEL for item in unique_items],
        "Catatan Otomatis": [guess_notes[item] for item in unique_items],
    })

    not_found = [item for item in unique_items if guesses[item] is None]
    if not_found:
        st.warning(
            f"Ada {len(not_found)} item yang otomatis ditandai TIDAK TERDAFTAR "
            "(lihat kolom 'Catatan Otomatis'). Silakan pilih manual di tabel bila kodenya memang ada."
        )

    edited_mapping = st.data_editor(
        mapping_df,
        column_config={
            "Nama Produk Accurate": st.column_config.SelectboxColumn(
                "Nama Produk Accurate", options=mapping_options, required=True
            ),
            "Catatan Otomatis": st.column_config.TextColumn("Catatan Otomatis", disabled=True),
        },
        disabled=["Nama Item (dari file)", "Catatan Otomatis"],
        use_container_width=True,
        hide_index=True,
        key="mapping_editor",
    )

    item_to_accurate = {}
    unregistered_items = set()
    for _, row in edited_mapping.iterrows():
        val = row["Nama Produk Accurate"]
        if val == UNREGISTERED_LABEL:
            unregistered_items.add(row["Nama Item (dari file)"])
        else:
            item_to_accurate[row["Nama Item (dari file)"]] = val

    # ---------------------------------------------------------------------
    # Step 4: generate output utama (Template Accurate, kolom lengkap)
    # ---------------------------------------------------------------------
    st.header("4. Hasil Konversi")

    output_rows = []
    for tanggal, sub in grouped.groupby("Tanggal"):
        sub_terdaftar = sub[~sub["NamaItem"].isin(unregistered_items)]
        if sub_terdaftar.empty:
            continue

        mmYY = tanggal.strftime("%m%y")
        dd = str(tanggal.day)
        number = f"Pustaka.{number_label}.{mmYY}.{dd}"
        date_str = tanggal.strftime("%d-%m-%Y")

        # Setiap baris item dalam grup (tanggal) yang sama tetap diisi
        # CUSTOMER NO / NUMBER / BRANCH / DATE-nya (tidak dikosongkan lagi).
        for _, r in sub_terdaftar.iterrows():
            accurate_name = item_to_accurate.get(r["NamaItem"])
            kode_produk = PRODUCT_CODE_MAP.get(normalize(accurate_name)) if accurate_name else None

            row = {h: None for h in OUTPUT_HEADERS}
            row["CUSTOMER NO"] = customer_no
            row["NUMBER"] = number
            row["BRANCH"] = "EKRAF"
            row["DATE"] = date_str
            row["ITEM:ITEM NO"] = kode_produk if kode_produk else accurate_name
            row["ITEM:QUANTITY"] = r["Qty"]
            row["ITEM:UNITPRICE"] = r["HargaSatuan"]
            row["ITEM:UNIT"] = "pcs"
            row["ITEM:WAREHOUSE NAME "] = warehouse_name
            row["ITEM:DEPT NAME"] = "EKRAF"
            output_rows.append(row)
    out_df = pd.DataFrame(output_rows, columns=OUTPUT_HEADERS)

    st.subheader("4a. Sheet Template (siap import ke Accurate)")
    st.caption(
        "Catatan: kolom yang tampil kosong di bawah ini memang sengaja dikosongkan "
        "(bukan teks \"None\") — saat di-export ke Excel, sel-sel ini akan benar-benar kosong. "
        "CUSTOMER NO, NUMBER, BRANCH, dan DATE diisi di setiap baris item (tidak digabung/dikosongkan)."
    )
    st.dataframe(out_df.fillna(""), use_container_width=True)

    # ---------------------------------------------------------------------
    # Step 4b: sheet Unregistered
    # ---------------------------------------------------------------------
    unregistered_rows = []
    if unregistered_items:
        unreg_df = grouped[grouped["NamaItem"].isin(unregistered_items)].sort_values(
            ["Tanggal", "NamaItem"]
        )
        for _, r in unreg_df.iterrows():
            unregistered_rows.append({
                "Tanggal": r["Tanggal"].strftime("%d-%m-%Y"),
                "Nama Item": r["NamaItem"],
                "Qty": r["Qty"],
                "Harga Satuan": r["HargaSatuan"],
                "Catatan": guess_notes.get(r["NamaItem"], "Ditandai manual sebagai TIDAK TERDAFTAR"),
            })

    unreg_out_df = pd.DataFrame(unregistered_rows, columns=UNREGISTERED_HEADERS)

    st.subheader("4b. Sheet Unregistered (item belum terdaftar)")
    if unreg_out_df.empty:
        st.success("Semua item sudah terpetakan ke Nama Produk Accurate. Tidak ada item unregistered. 🎉")
    else:
        st.warning(
            f"Ada {len(unreg_out_df)} baris item yang belum terdaftar / belum dipetakan. "
            "Baris-baris ini TIDAK ikut masuk ke sheet Template di atas."
        )
        st.dataframe(unreg_out_df, use_container_width=True)

    # ---------------------------------------------------------------------
    # Step 5: download — satu file, dua sheet (Template + Unregistered)
    # ---------------------------------------------------------------------
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="Template")
        unreg_out_df.to_excel(writer, index=False, sheet_name="Unregistered")
    buf.seek(0)

    st.download_button(
        "⬇️ Download File Template Accurate + Unregistered (.xlsx)",
        data=buf,
        file_name=f"Template_Accurate_{area.replace(' ', '')}_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Silakan upload file laporan penjualan (.xlsx) untuk melanjutkan.")
