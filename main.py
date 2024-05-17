import streamlit as st
import pandas as pd

# paling baru


# Database untuk memetakan nama perusahaan ke sektor atau jenis usaha (LU)
company_lu_mapping = {
    'Subak Uma Dalem': 'Pertanian, Kehutanan dan Perikanan',
    'Subak Keraman': 'Pertanian, Kehutanan dan Perikanan',
    'Limajari Interbhuana': 'Transportasi dan Pergudangan',
    'Putra Bhineka Perkasa, PT': 'Industri Pengolahan',
    'Sumiati Ekspor Internasional': 'Perdagangan Besar dan Eceran',
    'Hotel Merusaka Nusa Dua': 'Akmamin',
    'Lina Jaya, CV': 'Konstruksi',
    'Anugerah Merta Sedana, PT': 'Industri Pengolahan',
    'Jasamarga Bali Tol, PT': 'Transportasi dan Pergudangan',
    'Lotte Grosir Bali': 'Perdagangan Besar dan Eceran',
    'The Mulia Hotels and Resorts': 'Akmamin',
    'The Laguna Resort': 'Akmamin',
    'Hotel Grand Hyatt': 'Akmamin',
    'Bank Pembangunan Daerah Bali': 'Jasa Keuangan',
    'Lion Mentari Airlines KC': 'Transportasi dan Pergudangan',
    'Sheraton Bali Kuta Resort': 'Akmamin',
    'Sea Six Energy Indonesia, PT': 'Industri Pengolahan',
    'Subak Aseman IV': 'Pertanian, Kehutanan dan Perikanan'
}

def process_excel_file(uploaded_file, is_second_excel=False):
    combined_df = pd.DataFrame(columns=['Nama Contact', 'Pertanyaan', 'Nilai'])

    sheets = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')

    for sheet_name, df in sheets.items():
        st.header(f'Data pada Sheet: {sheet_name}')
        st.write("Dataframe utuh sebelum penyaringan:")
        st.dataframe(df)  # Tampilkan dataframe utuh sebelum penyaringan

        try:
            filtered_df = df[['Unnamed: 2', 'Unnamed: 3']]  # Ambil kolom Unnamed: 2 dan Unnamed: 3
            filtered_df.columns = ['Pertanyaan', 'Nilai']  # Ganti nama kolom agar sesuai

            # Filter hanya baris yang mengandung pertanyaan yang diinginkan
            relevant_columns = ['Permintaan Domestik  - Likert Scale',
                                'Permintaan Ekspor  - Likert Scale',
                                'Kapasitas Utilisasi - Likert Scale',
                                'Persediaan - Likert Scale',
                                'Investasi - Likert Scale',
                                'Biaya Energi - Likert Scale',
                                'Biaya Tenaga Kerja (Upah) - Likert Scale',
                                'Harga Jual – Likert Scale',
                                'Margin Usaha - Likert Scale',
                                'Tenaga Kerja - Likert Scale',
                                'Perkiraan Penjualan – Likert Scale',
                                'Perkiraan Tingkat Upah – Likert Scale',
                                'Perkiraan Harga Jual – Likert Scale',
                                'Perkiraan Jumlah Tenaga Kerja – Likert Scale',
                                'Perkiraan Investasi – Likert Scale']

            filtered_df = filtered_df[filtered_df['Pertanyaan'].isin(relevant_columns)]

            # Ubah tanda minus (–) menjadi strip (-) pada kolom Pertanyaan
            filtered_df['Pertanyaan'] = filtered_df['Pertanyaan'].str.replace('–', '-')

            # Tambahkan kolom 'Nama Contact' dengan nama sheet
            filtered_df.insert(0, 'Nama Contact', sheet_name)

            # Mencocokkan nama perusahaan dengan LU dari database
            filtered_df['LU'] = filtered_df['Nama Contact'].apply(
                lambda x: next((lu for company, lu in company_lu_mapping.items() if company in x), 'Unknown'))

            # Gabungkan dataframe hasil filter ke dalam dataframe kombinasi
            combined_df = pd.concat([combined_df, filtered_df])

            # Mengganti nilai kosong dengan kata "kosong" di combined_df
            combined_df = combined_df.fillna('kosong')
        except KeyError:
            st.error("Kolom yang dibutuhkan tidak ditemukan dalam file Excel.")

    if not combined_df.empty:
        st.header('Dataframe setelah penyaringan dari semua sheet:')
        st.dataframe(combined_df)

        st.header('Rata-rata nilai untuk setiap kolom Likert Scale:')
        avg_values = {}
        # Iterasi melalui setiap kolom Likert Scale
        for col in combined_df['Pertanyaan'].unique():
            # Konversi nilai-nilai ke tipe data numerik
            numeric_values = pd.to_numeric(combined_df[combined_df['Pertanyaan'] == col]['Nilai'], errors='coerce')
            # Hitung rata-rata nilai untuk kolom tersebut
            avg_value = numeric_values.mean()
            avg_values[col] = avg_value if not pd.isna(avg_value) else None

        # Tampilkan rata-rata nilai untuk setiap kolom Likert Scale
        for col, avg_value in avg_values.items():
            # Ubah "- Likert Scale" menjadi " "
            col_name = col.replace(' - Likert Scale', ' ')
            st.write(f"{col_name}: {avg_value}")

        if is_second_excel:
            # Set up dictionary to store total LU for each sector
            lu_totals = {}

            # Iterate through each company in the uploaded Excel file
            for company in combined_df['Nama Contact'].unique():
                # Get LU for the current company from the mapping
                lu = company_lu_mapping.get(company, 'Unknown')

                # Add LU to total count for the corresponding sector
                if lu in lu_totals:
                    lu_totals[lu] += 1
                else:
                    lu_totals[lu] = 1

            # Print total LU for each sector
            st.header('Total LU untuk setiap sektor atau jenis usaha (LU):')
            for lu, total in lu_totals.items():
                st.write(f"{lu}: {total}")

            # Calculate LU dominance
            total_lu = sum(lu_totals.values())
            lu_dominance = {lu: (count / total_lu) * 100 for lu, count in lu_totals.items()}
            max_dominant_lu = max(lu_dominance, key=lu_dominance.get)

            st.header('LU yang Mendominasi:')
            st.write(f"{max_dominant_lu}: {lu_dominance[max_dominant_lu]:.2f}%")

    return combined_df, avg_values

def process_domestik_ekspor_df(uploaded_file):
    combined_df_domestik_ekspor = pd.DataFrame(columns=['Nama Contact', 'Pertanyaan', 'Nilai', 'LU'])
    sheets = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')

    for sheet_name, df in sheets.items():
        try:
            relevant_columns = ['Permintaan Domestik  - Likert Scale', 'Permintaan Ekspor  - Likert Scale']
            df_filtered = df[['Unnamed: 2', 'Unnamed: 3']]  # Ambil kolom Unnamed: 2 dan Unnamed: 3
            df_filtered.columns = ['Pertanyaan', 'Nilai']  # Ganti nama kolom agar sesuai
            df_filtered = df_filtered[df_filtered['Pertanyaan'].isin(relevant_columns)]

            # Tambahkan kolom 'Nama Contact' dengan nama sheet
            df_filtered.insert(0, 'Nama Contact', sheet_name)

            # Mencocokkan nama perusahaan dengan LU dari database
            df_filtered['LU'] = df_filtered['Nama Contact'].apply(
                lambda x: next((lu for company, lu in company_lu_mapping.items() if company in x), 'Unknown'))

            # Gabungkan dataframe hasil filter ke dalam dataframe kombinasi
            combined_df_domestik_ekspor = pd.concat([combined_df_domestik_ekspor, df_filtered])

            # Mengganti nilai kosong dengan kata "kosong"
            combined_df_domestik_ekspor = combined_df_domestik_ekspor.fillna('kosong')
        except KeyError:
            st.error("Kolom yang dibutuhkan tidak ditemukan dalam file Excel.")

    # Menambahkan perhitungan jumlah orientasi domestik, ekspor, dan domestik & ekspor
    domestic_count = 0
    export_count = 0
    domestic_export_count = 0

    lu_domestic = []
    lu_export = []
    lu_domestic_export = []

    for contact in combined_df_domestik_ekspor['Nama Contact'].unique():
        df_contact = combined_df_domestik_ekspor[combined_df_domestik_ekspor['Nama Contact'] == contact]
        domestic_value = df_contact[df_contact['Pertanyaan'] == 'Permintaan Domestik  - Likert Scale']['Nilai'].values
        export_value = df_contact[df_contact['Pertanyaan'] == 'Permintaan Ekspor  - Likert Scale']['Nilai'].values

        lu_contact = df_contact['LU'].values[0]

        if len(domestic_value) > 0 and len(export_value) > 0:
            if domestic_value[0] != 'kosong' and export_value[0] == 'kosong':
                domestic_count += 1
                lu_domestic.append(lu_contact)
            elif domestic_value[0] == 'kosong' and export_value[0] != 'kosong':
                export_count += 1
                lu_export.append(lu_contact)
            elif domestic_value[0] != 'kosong' and export_value[0] != 'kosong':
                domestic_export_count += 1
                lu_domestic_export.append(lu_contact)

    total_count = domestic_count + export_count + domestic_export_count

    if total_count > 0:
        domestic_percentage = (domestic_count / total_count) * 100
        export_percentage = (export_count / total_count) * 100
        domestic_export_percentage = (domestic_export_count / total_count) * 100
    else:
        domestic_percentage = export_percentage = domestic_export_percentage = 0

    st.write(f"Jumlah orientasi Domestik: {domestic_count}")
    st.write(f"Jumlah orientasi Ekspor: {export_count}")
    st.write(f"Jumlah orientasi Domestik dan Ekspor: {domestic_export_count}")

    st.write(f"Persen orientasi Domestik: {domestic_percentage:.2f}%")
    st.write(f"Persen orientasi Ekspor: {export_percentage:.2f}%")
    st.write(f"Persen orientasi Domestik dan Ekspor: {domestic_export_percentage:.2f}%")

    lu_domestic = set(lu_domestic)
    lu_export = set(lu_export)
    lu_domestic_export = set(lu_domestic_export)

    st.write(f"LU yang berorientasi Domestik: {', '.join(lu_domestic)}")
    st.write(f"LU yang berorientasi Ekspor: {', '.join(lu_export)}")
    st.write(f"LU yang berorientasi Domestik dan Ekspor: {', '.join(lu_domestic_export)}")

    return combined_df_domestik_ekspor, total_count, domestic_count, export_count, domestic_export_count, lu_domestic, lu_export, lu_domestic_export

def main():
    st.set_page_config(page_title='ML Summary Liaison')
    st.title('Summary Liaison 📊')

    uploaded_file_1 = st.file_uploader('Upload file Excel Pertama (XLSX) untuk triwulan sebelumnya', type='xlsx')

    if uploaded_file_1:
        st.markdown('---')
        combined_df_1, avg_values_1 = process_excel_file(uploaded_file_1)

        st.write("\n---\n")
        st.write("Upload file Excel Kedua (XLSX) untuk triwulan sekarang")
        uploaded_file_2 = st.file_uploader('Upload file Excel Kedua (XLSX)', type='xlsx')

        if uploaded_file_2:
            st.markdown('---')
            combined_df_2, avg_values_2 = process_excel_file(uploaded_file_2, is_second_excel=True)

            # Tambahkan pemrosesan dataframe domestik ekspor
            st.header('Dataframe Permintaan Domestik dan Ekspor:')
            combined_df_domestik_ekspor, total_count, domestic_count, export_count, domestic_export_count, lu_domestic, lu_export, lu_domestic_export = process_domestik_ekspor_df(uploaded_file_2)
            st.dataframe(combined_df_domestik_ekspor)

            st.header('Perbandingan Rata-rata nilai antara dua upload:')
            if combined_df_1.empty or combined_df_2.empty:
                st.warning("Salah satu atau kedua dataframe kosong, perbandingan tidak dapat dilakukan.")
            else:
                changes = {'naik': 0, 'turun': 0}

                # Bandingkan nilai rata-rata antara dua dataframe
                for col in avg_values_1.keys():
                    if col in avg_values_2.keys():
                        change = avg_values_2[col] - avg_values_1[col]
                        if change > 0:
                            st.write(f"{col.replace(' - Likert Scale', ' ')}: Naik sebesar {change}")
                            changes['naik'] += 1
                        elif change < 0:
                            st.write(f"{col.replace(' - Likert Scale', ' ')}: Turun sebesar {abs(change)}")
                            changes['turun'] += 1
                        else:
                            st.write(f"{col.replace(' - Likert Scale', ' ')}: Tidak mengalami perubahan")

                # Generate comparison sentence
                quarter_now = st.text_input("Triwulan Sekarang (e.g., I 2024, II 2023, III 2022, IV 2024):")
                quarter_before = st.text_input("Triwulan Sebelumnya (e.g., I 2024, II 2023, III 2022, IV 2024):")
                phenomenon_reason = st.text_input(
                    "Alasan Fenomena yang terjadi (e.g., dinamika pasar, kondisi politik, dll):")

                if quarter_now and quarter_before and phenomenon_reason:
                    if changes['naik'] > changes['turun']:
                        trend = "percepatan"
                    elif changes['naik'] < changes['turun']:
                        trend = "perlambatan"
                    else:
                        trend = "tidak ada perubahan"

                    turun_indicators = [col.replace(' - Likert Scale', ' ') for col in avg_values_1.keys() if
                                        (col in avg_values_2.keys()) and (avg_values_2[col] - avg_values_1[col] < 0)]
                    st.header('Kesimpulan pertama:')
                    st.write(
                        f"Kinerja perekonomian Provinsi pada triwulan {quarter_now} terindikasi tumbuh melambat dibandingkan triwulan {quarter_before}. Hal ini sebagaimana tercermin dari hasil likert pada triwulan {quarter_now} yang mengalami {trend} pada {changes['turun']} dari {len(avg_values_1)} indikator likert liaison, yaitu {', '.join(turun_indicators)}. Fenomena ini terjadi karena {phenomenon_reason}.")

                    # Calculate total number of contacts
                    total_contacts = combined_df_domestik_ekspor['Nama Contact'].nunique()

                    # Calculate LU dominance
                    max_dominant_lu = combined_df_domestik_ekspor['LU'].value_counts().idxmax()
                    max_dominant_lu_percentage = (combined_df_domestik_ekspor['LU'].value_counts(normalize=True).max() * 100)

                    # Calculate orientation percentages
                    domestic_percentage = (domestic_count / total_count) * 100 if total_count > 0 else 0
                    export_percentage = (export_count / total_count) * 100 if total_count > 0 else 0
                    domestic_export_percentage = (domestic_export_count / total_count) * 100 if total_count > 0 else 0

                    # Generate second conclusion
                    st.header('Kesimpulan Kedua:')
                    st.write(
                        f"Jumlah total perusahaan yang di liaison KPw Bank Indonesia Provinsi Bali periode triwulan {quarter_now} adalah {total_contacts} kontak. Liaison pada triwulan laporan didominasi oleh LU {max_dominant_lu} sebesar {max_dominant_lu_percentage:.2f}% dari total kontak. Kemudian, {domestic_percentage:.2f}% berorientasi domestik, {export_percentage:.2f}% berorientasi ekspor dan {domestic_export_percentage:.2f}% berorientasi domestik dan ekspor. Perusahaan yang sepenuhnya berorientasi domestik adalah LU {', '.join(lu_domestic)}. Kontak yang sepenuhnya berorientasi domestik dan ekspor adalah LU {', '.join(lu_domestic_export)}. Sedangkan, kontak yang sepenuhnya berorientasi ekspor adalah beberapa kontak pada LU {', '.join(lu_export)}.")

if __name__ == "__main__":
    main()
