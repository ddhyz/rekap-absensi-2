# ===============================
# app.py (Index di dalam folder templates) - versi sheet 1 & 2
# ===============================
from flask import Flask, render_template, request, send_file
import pandas as pd
from datetime import datetime, date
import re
import os
from werkzeug.utils import secure_filename
from docxtpl import DocxTemplate

app = Flask(__name__)

# --- Konfigurasi upload ---
UPLOAD_FOLDER = "uploads"
ALLOWED_EXTENSIONS = {"xlsx", "xls"}
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# --- Helper: validasi ekstensi ---
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

# --- Helper: bersihkan ID ---
def clean_id(id_value):
    if pd.isna(id_value):
        return ""
    id_str = str(id_value).strip()
    if id_str.endswith('.0'):
        id_str = id_str[:-2]
    return id_str

# --- Helper: sort alfanumerik ---
def sort_nicely(l):
    def convert(text):
        return int(text) if text.isdigit() else text.lower()
    def alphanum_key(key):
        return [convert(c) for c in re.split('([0-9]+)', key)]
    return sorted(l, key=alphanum_key)

@app.route("/", methods=["GET", "POST"])
def index():
    df_telat_html = df_tidak_hadir_html = df_jumlah_absen_html = None
    download_file = None
    hasil_rekap_filename = None
    panggilan_files = []

    jumlah_karyawan_telat = 0
    jumlah_karyawan_tidak_hadir = 0
    jumlah_total_karyawan = 0

    if request.method == "POST":
        file = request.files.get("file_excel")
        if not file or file.filename == '':
            return render_template("index.html")
        if not allowed_file(file.filename):
            return render_template("index.html", error_message="File harus .xlsx atau .xls")

        # Simpan file upload
        safe_name = secure_filename(file.filename)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        uploaded_filename = f"{timestamp}_{safe_name}"
        file_path = os.path.join(UPLOAD_FOLDER, uploaded_filename)
        file.save(file_path)
        download_file = uploaded_filename

        # === PROSES EXCEL: baca sheet 1 & 2 ===
        all_sheets = pd.read_excel(file_path, sheet_name=None)  # dict {nama_sheet: DataFrame}
        df_list = []
        for sheet_name, df_sheet in all_sheets.items():
            # Ambil kolom sesuai format
            max_cols = len(df_sheet.columns)
            column_mapping = {}
            column_names = ['Perusahaan', 'Nama', 'ID', 'Tgl/Waktu', 'Mesin_ID', 'Kolom6', 'Status', 'Kolom8']
            for i in range(min(max_cols, len(column_names))):
                column_mapping[column_names[i]] = df_sheet.iloc[:, i]
            df_fix_sheet = pd.DataFrame(column_mapping)
            df_list.append(df_fix_sheet)

        # Gabung semua sheet menjadi satu DataFrame
        df_fix = pd.concat(df_list, ignore_index=True)

        # ID unik
        semua_id_dari_file = [clean_id(idv) for idv in df_fix["ID"] if clean_id(idv) != ""]
        semua_id_unik = sort_nicely(list(set(semua_id_dari_file)))
        jumlah_total_karyawan = len(semua_id_unik)

        # Normalisasi
        df_fix["Nama"] = df_fix["Nama"].astype(str).str.strip()
        df_fix["ID"] = df_fix["ID"].apply(clean_id)
        df_fix = df_fix[df_fix["Nama"].notna() & (df_fix["Nama"] != "nan") & (df_fix["Nama"] != "")]
        df_fix = df_fix[df_fix["ID"].notna() & (df_fix["ID"] != "nan") & (df_fix["ID"] != "")]
        df_fix["Tgl/Waktu"] = pd.to_datetime(df_fix["Tgl/Waktu"], dayfirst=True, errors='coerce')
        df_fix = df_fix.dropna(subset=["Tgl/Waktu"])
        df_fix["Tanggal_Saja"] = df_fix["Tgl/Waktu"].dt.date
        df_fix = df_fix.drop_duplicates(subset=["ID", "Tanggal_Saja"])

        # Telat pagi
        jam_telat = datetime.strptime("07:50:00", "%H:%M:%S").time()
        df_pagi = df_fix[(df_fix["Tgl/Waktu"].dt.hour >=5) & (df_fix["Tgl/Waktu"].dt.hour <=9)]
        id_to_nama = dict(zip(df_fix["ID"], df_fix["Nama"]))

        # Rentang tanggal kerja (exclude Minggu)
        if not df_fix["Tgl/Waktu"].empty:
            tanggal_awal = df_fix["Tgl/Waktu"].dt.date.min()
            tanggal_akhir = df_fix["Tgl/Waktu"].dt.date.max()
            semua_tanggal = [tgl for tgl in pd.date_range(tanggal_awal, tanggal_akhir).date if pd.Timestamp(tgl).weekday() !=6]
        else:
            semua_tanggal = []

        # Rekap telat
        rekap_telat = []
        for id_karyawan in semua_id_unik:
            nama_karyawan = id_to_nama.get(id_karyawan, "Unknown")
            data_id = df_pagi[df_pagi["ID"]==id_karyawan]
            telat_id = data_id[data_id["Tgl/Waktu"].dt.time > jam_telat]
            for _, row in telat_id.iterrows():
                rekap_telat.append({"ID": id_karyawan, "Nama": nama_karyawan, "Tgl/Waktu Telat": row["Tgl/Waktu"]})
        df_telat = pd.DataFrame(rekap_telat)

        # Rekap tidak hadir & jumlah absen
        rekap_tidak_hadir = []
        jumlah_absen_total = []
        for id_karyawan in semua_id_unik:
            nama_karyawan = id_to_nama.get(id_karyawan, "Unknown")
            data_id = df_fix[df_fix["ID"]==id_karyawan]
            hadir_tanggal = set(data_id["Tgl/Waktu"].dt.date) if not data_id.empty else set()
            tidak_hadir_tanggal = [tgl for tgl in semua_tanggal if tgl not in hadir_tanggal]
            for tgl in tidak_hadir_tanggal:
                rekap_tidak_hadir.append({"ID": id_karyawan, "Nama": nama_karyawan, "Tanggal Tidak Hadir": tgl})
            hadir_per_tanggal = len([tgl for tgl in semua_tanggal if tgl in hadir_tanggal])
            jumlah_absen_total.append({"ID": id_karyawan, "Nama": nama_karyawan, "Jumlah Absen Total": hadir_per_tanggal})
        df_tidak_hadir = pd.DataFrame(rekap_tidak_hadir)
        df_jumlah_absen = pd.DataFrame(jumlah_absen_total)

        # Hitung jumlah telat & tidak hadir
        if not df_telat.empty:
            jumlah_telat = df_telat.groupby("ID").size().reset_index(name="Jumlah Telat")
            df_jumlah_absen = pd.merge(df_jumlah_absen, jumlah_telat, on="ID", how="left")
        else:
            df_jumlah_absen["Jumlah Telat"] = 0

        if not df_tidak_hadir.empty:
            jumlah_tidak_hadir = df_tidak_hadir.groupby("ID").size().reset_index(name="Jumlah Tidak Hadir")
            df_jumlah_absen = pd.merge(df_jumlah_absen, jumlah_tidak_hadir, on="ID", how="left")
        else:
            if "Jumlah Tidak Hadir" not in df_jumlah_absen.columns:
                df_jumlah_absen["Jumlah Tidak Hadir"] = 0

        df_jumlah_absen[["Jumlah Telat","Jumlah Tidak Hadir"]] = df_jumlah_absen[["Jumlah Telat","Jumlah Tidak Hadir"]].fillna(0).astype(int)

        # Statistik
        jumlah_karyawan_telat = len(set(df_telat["ID"].unique())) if not df_telat.empty else 0
        jumlah_karyawan_tidak_hadir = len(set(df_tidak_hadir["ID"].unique())) if not df_tidak_hadir.empty else 0

        # Filter tidak hadir >3 hari
        df_tidak_hadir_lebih3 = df_jumlah_absen[df_jumlah_absen["Jumlah Tidak Hadir"]>3].copy()

        # Buat file Excel rekap
        hasil_rekap_filename = f"hasil_rekap_{uploaded_filename}"
        hasil_rekap_path = os.path.join(UPLOAD_FOLDER, hasil_rekap_filename)
        with pd.ExcelWriter(hasil_rekap_path) as writer:
            if not df_telat.empty:
                df_telat.to_excel(writer, sheet_name="Karyawan Telat", index=False)
            if not df_tidak_hadir.empty:
                df_tidak_hadir.to_excel(writer, sheet_name="Karyawan Tidak Hadir", index=False)
            df_jumlah_absen.to_excel(writer, sheet_name="Jumlah Kehadiran", index=False)
            if not df_tidak_hadir_lebih3.empty:
                df_tidak_hadir_lebih3.to_excel(writer, sheet_name=">3 Hari Tidak Hadir", index=False)

        # === Surat Panggilan Menggunakan Template Word + Hari Nama ===
        hari_list = ["Senin","Selasa","Rabu","Kamis","Jumat","Sabtu","Minggu"]
        for _, row in df_tidak_hadir_lebih3.iterrows():
            spg_filename = f"surat_panggilan_{row['ID']}_{uploaded_filename.rsplit('.',1)[0]}.docx"
            spg_path = os.path.join(UPLOAD_FOLDER, spg_filename)
            template_path = os.path.join("templates", "template_surat_panggilan.docx")
            doc = DocxTemplate(template_path)

            df_absen_id = df_tidak_hadir[df_tidak_hadir["ID"]==row['ID']]
            if not df_absen_id.empty:
                semua_tgl = df_absen_id["Tanggal Tidak Hadir"].apply(lambda x: x.strftime("%d-%m-%Y")).tolist()
                tanggal_terakhir = ", ".join(semua_tgl)
                jumlah_hari = len(semua_tgl)
            else:
                tanggal_terakhir = date.today().strftime("%d-%m-%Y")
                jumlah_hari = row['Jumlah Tidak Hadir']

            tanggal_surat = date.today()
            nama_hari = hari_list[tanggal_surat.weekday()]
            context = {
                "NAMA": row['Nama'],
                "ID": row['ID'],
                "JUMLAH_HARI": jumlah_hari,
                "TANGGAL_ABSEN": tanggal_terakhir,
                "TANGGAL_SURAT": f"{nama_hari}, {tanggal_surat.strftime('%d-%m-%Y')}"
            }

            doc.render(context)
            doc.save(spg_path)

            panggilan_files.append({
                "ID": row['ID'],
                "Nama": row['Nama'],
                "Jumlah Tidak Hadir": row['Jumlah Tidak Hadir'],
                "Surat Panggilan": spg_filename
            })

        # Konversi ke HTML untuk web
        df_telat_html = df_telat.to_html(classes="table table-striped", index=False) if not df_telat.empty else "<p class='text-center'>Tidak ada data karyawan telat</p>"
        df_tidak_hadir_html = df_tidak_hadir.to_html(classes="table table-striped", index=False) if not df_tidak_hadir.empty else "<p class='text-center'>Tidak ada data karyawan tidak hadir</p>"
        df_jumlah_absen_html = df_jumlah_absen.to_html(classes="table table-striped", index=False)

    return render_template(
        "index.html",
        df_telat=df_telat_html,
        df_tidak_hadir=df_tidak_hadir_html,
        df_jumlah_absen=df_jumlah_absen_html,
        download_file=download_file,
        hasil_rekap_filename=hasil_rekap_filename,
        jumlah_karyawan_telat=jumlah_karyawan_telat,
        jumlah_karyawan_tidak_hadir=jumlah_karyawan_tidak_hadir,
        jumlah_total_karyawan=jumlah_total_karyawan,
        panggilan_files=panggilan_files,
        error_message=None,
    )

@app.route("/download/<filename>")
def download(filename):
    file_path = os.path.join(UPLOAD_FOLDER, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return "File tidak ditemukan!", 404

if __name__ == "__main__":
    app.run(debug=True)

