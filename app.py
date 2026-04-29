import streamlit as st
import pandas as pd
from modul import DataFilterAndSelect, ConfigurationInput, PaymentCount, PaymentExcelBuilder

@st.cache_data(show_spinner=False)
def convert_for_download(df):
    return df.to_csv(index=False).encode("utf-8")

@st.cache_data(show_spinner=False)
def process_uploaded_csv(file):
    df = pd.read_csv(file, encoding="utf-8")
    clean_data = DataFilterAndSelect(df)
    return clean_data.filter_and_select()

def main():
    st.title("🛠️ Gajian Kompensasi dan Angkutan  App")
    tab1, tab2, tab3 = st.tabs(["📄 Inisialisasi", "🧱 Rekap Data", "📦 Download File Gajian"])

    var1 = None
    var2 = None

    with tab1:
        st.header("📥 Inisialisasi")
        uploaded_initial_file = st.file_uploader("📤 Upload Data (CSV UTF-8 Volker)", type=["csv"])

        st.header("💰 Pengaturan Harga")
        harga_kompensasi_input = st.number_input(
            "Harga Kompensasi (per titik)", 
            min_value=0, 
            value=100000,
            step=1000,
            help="Masukkan tarif kompensasi lahan yang akan digunakan pada perhitungan"
        )
        st.session_state["harga_kompensasi"] = harga_kompensasi_input

        if uploaded_initial_file:
            try:
                var1 = process_uploaded_csv(uploaded_initial_file)
                var2 = ConfigurationInput()

                if "stage1_result" not in st.session_state:
                    st.session_state["stage1_result"] = var2.process_stage1(var1)
                    st.session_state["merged_data"] = None

                st.success("✅ Data utama berhasil diproses.")

                st.download_button(
                    label="⬇️ Download Template Lokasi dan Tanggal",
                    data=convert_for_download(st.session_state["stage1_result"]),
                    file_name="template_lokasi_dan_tanggal.csv",
                    mime="text/csv",
                    key="download_stage1"
                )

                uploaded_file = st.file_uploader("Upload Template Lokasi dan Tanggal", type=["csv", "xlsx"])
                if uploaded_file:
                    try:
                        if uploaded_file.name.endswith(".csv"):
                            stage1_df = pd.read_csv(uploaded_file)
                        else:
                            stage1_df = pd.read_excel(uploaded_file)

                        date_cols = [
                            "Tanggal Mulai (2025-05-23)",
                            "Tanggal Selesai (2025-05-23)",
                            "Tanggal Gajian (2025-05-23)"
                        ]
                        for col in date_cols:
                            if col in stage1_df.columns:
                                stage1_df[col] = pd.to_datetime(stage1_df[col], dayfirst=True, errors='coerce').dt.normalize()

                        st.success("✅ Template berhasil diupload")
                        st.dataframe(stage1_df)

                        st.session_state["stage1_result"] = stage1_df
                        st.session_state["merged_data"] = var2.process_stage(var1, stage1_df)

                    except Exception as e:
                        st.error(f"❌ Gagal mengakses file: {e}")
                else:
                    st.info("Masih menggunakan data default.")
                    st.session_state["merged_data"] = var2.process_stage(var1, st.session_state["stage1_result"])

                st.header("🧪 Kelompok Data")
                if st.session_state["merged_data"] is not None and not st.session_state["merged_data"].empty:
                    st.dataframe(st.session_state["merged_data"])
                else:
                    st.warning("⚠️ Kelompok data belum tersedia.")

            except Exception as e:
                st.error(f"❌ Gagal memproses file utama: {e}")
        else:
            st.warning("⚠️ Silakan upload file data utama CSV terlebih dahulu.")

    with tab2:
        st.header("🧱 Rekap Data")
        if st.session_state.get("merged_data") is not None and not st.session_state["merged_data"].empty:
            st.dataframe(st.session_state["merged_data"])

            if st.button("▶️ Process Payment Calculation", key='Procces'):
                processor = PaymentCount()
                
                # Menggunakan harga_kompensasi dari session_state
                harga_kompensasi = st.session_state.get("harga_kompensasi", 90000)

                result_df = (
                    processor
                    .set_data(st.session_state["merged_data"])
                    .harga_kompensasi(tarif=harga_kompensasi)
                    .harga_angkutan()
                    .get_result()
                )

                st.session_state["payment_result"] = result_df
                st.session_state["payment_processor"] = processor
                st.success("✅ Perhitungan gajian berhasil dilakukan.")

            if "payment_result" in st.session_state:
                st.subheader("💰 Hasil Perhitungan")
                st.dataframe(st.session_state["payment_result"])

                st.download_button(
                    label="⬇️ Download CSV Perhitungan Gajian",
                    data=convert_for_download(st.session_state["payment_result"]),
                    file_name="CSV Perhitungan Gajian.csv",
                    mime="text/csv",
                    key="procces2"
                )

                payment_processor = st.session_state.get("payment_processor")
                if payment_processor is not None:
                    pivot_df = payment_processor.get_pivot_summary()

                    if not pivot_df.empty:
                        st.subheader("📊 Rekap Total Pembayaran")
                        st.dataframe(pivot_df)

                        st.download_button(
                            label="⬇️ Download Rekap Pembayaran per titik",
                            data=pivot_df.to_csv(index=False).encode("utf-8"),
                            file_name="rekap_pembayaran.csv",
                            mime="text/csv"
                        )
        else:
            st.warning("⚠️ Silakan unggah data di Tab 1 terlebih dahulu.")

    with tab3:
        st.header("📦 Download File Gajian")

        st.markdown("### Tanggal Gajian")
        date_input = st.date_input("Tanggal Dokumen", value=None)
        location_input = st.text_input("Lokasi", value="Meliau")

        date_text = f"{location_input}, {date_input.strftime('%d %B %Y')}" if date_input else ""

        st.markdown("### IUP")
        iup = st.text_input("IUP", value="MCU")

        st.markdown("### Penandatangan")
        signer_b_name = st.text_input("Admin - Nama", value="Dodi Prasetyo")
        signer_b_title = st.text_input("Jabatan", value="Keu. / Umum")

        signer_d_name = st.text_input("Geos - Nama", value="Prya Arif Rahman")
        signer_d_title = st.text_input("Jabatan", value="Geologist")

        signers = {
            "B": (signer_b_name, signer_b_title),
            "D": (signer_d_name, signer_d_title),
        }

        df = st.session_state.get("payment_result")

        if df is not None:
            df["Tanggal Sampling"] = pd.to_datetime(df["Tanggal Sampling"], errors='coerce')
            df["Tanggal Sampling"] = df["Tanggal Sampling"].dt.strftime('%Y-%m-%d')
            output_file = f'Gajian Kompensasi dan Angkutan  IUP OP {iup} {date_text}.xlsx'

            if st.button("Generate Excel", key="asd"):
                builder = PaymentExcelBuilder(df)
                builder.create_multi_payment_excel(
                    output_file=output_file,
                    date_text=date_text,
                    signers=signers
                )
                with open(output_file, "rb") as f:
                    st.download_button("Download Excel", f, file_name=output_file)
        else:
            st.warning("⚠️ Harap lakukan proses pembayaran di Tab 2 terlebih dahulu.")

if __name__ == "__main__":
    main()
