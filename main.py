import streamlit as st

st.set_page_config(page_title="Menu Konversi Accurate", layout="wide")

st.sidebar.title("📋 Pilih Menu")
menu = st.sidebar.radio(
    "Menu",
    ["LTQ", "Isma", "Payment Isma", "Payment Piutang Isma"],
    label_visibility="collapsed",
)

# Jalankan ulang script menu yang dipilih setiap kali user berinteraksi
# (pakai exec, bukan import, supaya kode top-level di file menu selalu
# di-run ulang mengikuti siklus rerun Streamlit)
if menu == "LTQ":
    with open("ltq_menu.py", encoding="utf-8") as f:
        exec(f.read(), {"__name__": "__main__"})
elif menu == "Isma":
    with open("isma_menu.py", encoding="utf-8") as f:
        code = f.read()
    if not code.strip():
        st.title("📚 Menu Isma")
        st.warning("isma_menu.py masih kosong — belum ada kode untuk menu ini.")
    else:
        namespace = {"__name__": "__main__"}
        exec(code, namespace)
        # isma_menu.py hanya MENDEFINISIKAN render_isma_menu(), jadi harus
        # dipanggil manual di sini supaya benar-benar tampil (bukan blank).
        namespace["render_isma_menu"]()
elif menu == "Payment Isma":
    with open("payment_menu.py", encoding="utf-8") as f:
        code = f.read()
    if not code.strip():
        st.title("💳 Menu Payment Isma")
        st.warning("payment_menu.py masih kosong — belum ada kode untuk menu ini.")
    else:
        namespace = {"__name__": "__main__"}
        exec(code, namespace)
        # payment_menu.py hanya MENDEFINISIKAN render_payment_menu(), jadi
        # harus dipanggil manual di sini supaya benar-benar tampil (bukan blank).
        namespace["render_payment_menu"]()
elif menu == "Payment Piutang Isma":
    with open("payment_piutang_menu.py", encoding="utf-8") as f:
        code = f.read()
    if not code.strip():
        st.title("💳 Menu Payment Piutang Isma")
        st.warning("payment_piutang_menu.py masih kosong — belum ada kode untuk menu ini.")
    else:
        namespace = {"__name__": "__main__"}
        exec(code, namespace)
        # payment_piutang_menu.py hanya MENDEFINISIKAN render_payment_piutang_menu(),
        # jadi harus dipanggil manual di sini supaya benar-benar tampil (bukan blank).
        namespace["render_payment_piutang_menu"]()
