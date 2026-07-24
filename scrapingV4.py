import base64
import os
import re
import time
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from selenium import webdriver

URL_LOGIN = "https://haccp.kkp.go.id/h4/login/"
OUTPUT_FILE = "rekap_surveilan_detail.xlsx"
GAGAL_FILE = "gagal_dibuka.txt"
JEDA_ANTAR_HALAMAN = 1.5

UNDUH_DOKUMEN = True
FOLDER_UNDUHAN = "dokumen_terunduh"

MAX_PERCOBAAN = 4
JEDA_RETRY = 5
JEDA_RETRY_BERTAMBAH = True 

MIN_KOLOM_INSPEKTUR = 3

CHECKLIST_ITEMS_SURVEILAN = [
    "Sampul Berkas Inspeksi",
    "Audit Kecukupan",
    "Data Umum UPI",
    "Daftar Hadir Pertemuan",
    "Check List",
    "Daftar Temuan Ketidaksesuaian",
    "Laporan Inspektur Mutu",
    "Laporan Tindakan Perbaikan",
    "Pemeriksaan Tindakan Perbaikan",
    "Evaluasi Hasil Surveilan",
    "Surat Keterangan Hasil Surveilan",
]

CHECKLIST_ITEMS_INSPEKSI = [
    "Sampul Berkas Inspeksi",
    "Audit Kecukupan",
    "Data Umum UPI",
    "Survei Kepuasan Pelayanan",
    "Daftar Hadir Pertemuan",
    "Check List",
    "Daftar Temuan Ketidaksesuaian",
    "Laporan Inspektur Mutu",
    "Laporan Singkat Inspektur Mutu",
    "REKOMENDASI HASIL INSPEKSI",
    "SERTIFIKAT PENERAPAN HACCP",
    "Laporan Tindakan Perbaikan",
    "Pemeriksaan Tindakan Perbaikan",
]

ITEM_SKV = "Surat Keterangan Hasil Surveilan"
ITEM_SERTIFIKAT = "SERTIFIKAT PENERAPAN HACCP"

MARKER_INFO = {"Surveilan": "INFO SURVEILAN", "Inspeksi": "INFO INSPEKSI"}
MARKER_DOK = {"Surveilan": "DOKUMENTASI SURVEILAN", "Inspeksi": "DOKUMENTASI INSPEKSI"}

HEADER_WARNA_PER_JENIS = {"Surveilan": "#7232bd", "Inspeksi": "#3b5998"}
CHECKLIST_ITEMS_PER_JENIS = {
    "Surveilan": CHECKLIST_ITEMS_SURVEILAN,
    "Inspeksi": CHECKLIST_ITEMS_INSPEKSI,
}

ITEM_DOKUMEN_AKHIR_PER_JENIS = {
    "Surveilan": ITEM_SKV,
    "Inspeksi": ITEM_SERTIFIKAT,
}




def deteksi_jenis(html):

    for jenis, marker in MARKER_INFO.items():
        if marker in html:
            return jenis
    return None


def _clean(text):

    if text is None:
        return ""
    text = text.replace("\u00a0", " ")
    return re.sub(r"\s+", " ", text).strip()


def ambil_daftar_link(html, base_url):

    soup = BeautifulSoup(html, "html.parser")
    links = []
    seen = set()
    for a in soup.select("table tr td a[href]"):
        href = a["href"]
        if "idi=" not in href or "type=dtl" not in href:
            continue
        full_url = urljoin(base_url, href)
        if full_url in seen:
            continue
        seen.add(full_url)
        links.append(full_url)
    return links


def _cari_kotak(soup, teks_penanda):

    node = soup.find(string=re.compile(re.escape(teks_penanda)))
    if not node:
        return None
    box = node.find_parent("div")
    while box is not None:
        style = box.get("style", "")
        if re.search(r"border-radius:\s*25px", style):
            return box
        box = box.find_parent("div")
    return None


def halaman_valid(html):

    if not html:
        return False
    if len(html) < 1000:
        return False
    if deteksi_jenis(html) is None:
        return False
    return True


def ambil_html_dengan_retry(driver, url):

    for percobaan in range(1, MAX_PERCOBAAN + 1):
        try:
            driver.get(url)
            time.sleep(JEDA_ANTAR_HALAMAN)
            html = driver.page_source
        except Exception as e:
            print(f"  [!] Percobaan {percobaan}/{MAX_PERCOBAAN} error saat membuka: {e}")
            html = None

        if halaman_valid(html):
            if percobaan > 1:
                print(f"  [OK] Berhasil di percobaan ke-{percobaan}.")
            return html

        if percobaan < MAX_PERCOBAAN:
            jeda = JEDA_RETRY * percobaan if JEDA_RETRY_BERTAMBAH else JEDA_RETRY
            print(f"  [!] Halaman kosong/gagal render (percobaan {percobaan}/{MAX_PERCOBAAN}), "
                  f"coba lagi dalam {jeda} detik...")
            time.sleep(jeda)
            try:
                driver.refresh()
                time.sleep(JEDA_ANTAR_HALAMAN)
            except Exception:
                pass

    return None


def ekstrak_info(html, jenis):

    soup = BeautifulSoup(html, "html.parser")
    data = {
        "nama": "",
        "tanggal": "",
        "status": "",
        "alamat": "",
        "inspektur": [],
        "ruang_lingkup": [],
    }

    box = _cari_kotak(soup, MARKER_INFO[jenis])
    if box is None:
        return data

    divs = box.find_all("div", recursive=False)
    texts = [_clean(d.get_text()) for d in divs]

    try:
        idx = texts.index("Nama")
        data["nama"] = texts[idx + 3]
        data["tanggal"] = texts[idx + 4]
        data["status"] = texts[idx + 5]
    except (ValueError, IndexError):
        pass

    try:
        idx = texts.index("Alamat")
        data["alamat"] = texts[idx + 1]
    except (ValueError, IndexError):
        pass


    for i, t in enumerate(texts):
        if t.startswith("Inspektur Mutu yang Bertugas"):
            m = re.search(r"=\s*(\d+)\s*Orang", t)
            jumlah = int(m.group(1)) if m else None
            sisa = texts[i + 1:]
            try:
                mulai = sisa.index("Noreg. Inspektur") + 1
            except ValueError:
                mulai = 0
            nonkosong = [x for x in sisa[mulai:] if x]
            n = jumlah if jumlah is not None else len(nonkosong) // 3
            for k in range(n):
                bagian = nonkosong[k * 3:(k + 1) * 3]
                if len(bagian) < 3:
                    break
                nama_lokasi, posisi, noreg = bagian
                if " Balai" in nama_lokasi:
                    nama_org, lokasi = nama_lokasi.split(" Balai", 1)
                    lokasi = "Balai" + lokasi
                else:
                    nama_org, lokasi = nama_lokasi, ""
                data["inspektur"].append({
                    "nama": nama_org.strip(),
                    "lokasi": lokasi.strip(),
                    "posisi": posisi,
                    "noreg": noreg,
                })
            break

    for i, t in enumerate(texts):
        if t.startswith("Ruang Lingkup yang Diperiksa"):
            m = re.search(r"=\s*(\d+)\s*Produk", t)
            jumlah = int(m.group(1)) if m else None
            sisa = texts[i + 1:]
            try:
                idx_grade = sisa.index("Grade")
            except ValueError:
                idx_grade = -1
            ada_keterangan = (
                idx_grade != -1
                and idx_grade + 1 < len(sisa)
                and sisa[idx_grade + 1] == "Keterangan"
            )
            mulai = idx_grade + 2 if ada_keterangan else idx_grade + 1
            if mulai < 0:
                mulai = 0
            nonkosong = [x for x in sisa[mulai:] if x]
            lebar = 3 if ada_keterangan else 2
            n = jumlah if jumlah is not None else len(nonkosong) // lebar
            for k in range(n):
                bagian = nonkosong[k * lebar:(k + 1) * lebar]
                if len(bagian) < lebar:
                    break
                if ada_keterangan:
                    nama_produk, grade, keterangan = bagian
                    data["ruang_lingkup"].append(
                        {"nama": nama_produk, "grade": grade, "keterangan": keterangan}
                    )
                else:
                    nama_produk, grade = bagian
                    data["ruang_lingkup"].append({"nama": nama_produk, "grade": grade})
            break

    return data


def _segmen_dokumentasi(html, jenis):
    soup = BeautifulSoup(html, "html.parser")
    box = _cari_kotak(soup, MARKER_DOK[jenis])
    if box is None:
        return {}

    children = box.find_all("div", recursive=False)
    warna_header = HEADER_WARNA_PER_JENIS[jenis]

    header_pos = []
    for i, d in enumerate(children):
        style = d.get("style", "")
        if warna_header not in style:
            continue
        teks = _clean(d.get_text())
        m = re.match(r"^\d+\.\s*(.+)$", teks)
        if m:
            header_pos.append((i, m.group(1).strip()))

    segmen = {}
    for idx, (pos, nama_item) in enumerate(header_pos):
        akhir = header_pos[idx + 1][0] if idx + 1 < len(header_pos) else len(children)
        segmen[nama_item] = children[pos:akhir]
    return segmen


def ekstrak_checklist_dokumentasi(segmen, jenis):
    items = CHECKLIST_ITEMS_PER_JENIS[jenis]
    hasil = {item: False for item in items}

    for nama_item, potongan in segmen.items():
        segmen_html = "".join(str(c) for c in potongan)
        sudah_lengkap = "Belum Unggah" not in segmen_html

        if nama_item in hasil:
            hasil[nama_item] = sudah_lengkap
        else:
            for item in items:
                if item.lower() == nama_item.lower():
                    hasil[item] = sudah_lengkap
                    break
            else:
                print(f"  [!] Item dokumentasi tidak dikenali: '{nama_item}'")

    return hasil


def _ekstrak_detail_skv(segmen):
    hasil = {"nomor": "", "tanggal": "", "penandatangan": ""}

    potongan = segmen.get(ITEM_SKV)
    if not potongan:
        return hasil

    texts = [_clean(d.get_text()) for d in potongan]
    nonkosong = [t for t in texts if t]

    try:
        idx = nonkosong.index("Penandatangan")
    except ValueError:
        return hasil

    sisa = nonkosong[idx + 1:]
    try:
        mulai = sisa.index("Tanggal") + 1
    except ValueError:
        mulai = 0
    data = sisa[mulai:]

    if len(data) >= 1:
        hasil["penandatangan"] = data[0]
    if len(data) >= 2:
        hasil["nomor"] = data[1]
    if len(data) >= 3:
        hasil["tanggal"] = data[2]

    return hasil


def _ekstrak_detail_sertifikat(segmen):

    hasil = {"daftar": [], "produk": "", "nomor": "", "batas": ""}

    potongan = segmen.get(ITEM_SERTIFIKAT)
    if not potongan:
        return hasil

    texts = [_clean(d.get_text()) for d in potongan]
    nonkosong = [t for t in texts if t]

    try:
        idx_batas = nonkosong.index("Batas Berlaku")
    except ValueError:
        return hasil

    ada_status = idx_batas >= 1 and nonkosong[idx_batas - 1] == "Status"
    lebar = 4 if ada_status else 3
    data = nonkosong[idx_batas + 1:]

    n = len(data) // lebar
    for k in range(n):
        bagian = data[k * lebar:(k + 1) * lebar]
        if len(bagian) < lebar:
            break
        if ada_status:
            produk, nomor, status, batas = bagian
        else:
            produk, nomor, batas = bagian
            status = ""
        hasil["daftar"].append(
            {"produk": produk, "nomor": nomor, "status": status, "batas": batas}
        )

    hasil["produk"] = "; ".join(dict.fromkeys(x["produk"] for x in hasil["daftar"] if x["produk"]))
    hasil["nomor"] = "; ".join(dict.fromkeys(x["nomor"] for x in hasil["daftar"] if x["nomor"]))
    hasil["batas"] = next((x["batas"] for x in hasil["daftar"] if x["batas"]), "")
    return hasil


def _sanitasi_nama_file(teks):
    teks = _clean(teks)
    teks = re.sub(r'[\\/:*?"<>|]', "_", teks)
    teks = teks.strip(" .")
    return teks or "tanpa_nama"


def buat_session_dari_driver(driver):
    sess = requests.Session()
    for c in driver.get_cookies():
        sess.cookies.set(c["name"], c["value"])
    sess.headers.update({"User-Agent": "Mozilla/5.0"})
    return sess


def _cari_semua_link_unggah(potongan):
    frag = BeautifulSoup("".join(str(c) for c in potongan), "html.parser")
    hasil = []
    for a in frag.find_all("a", href=True):
        img = a.find("img")
        if img and "storage.png" in img.get("src", ""):
            hasil.append(a["href"])
    return hasil


def _decode_ekstensi_download(href):
    try:
        token = href.split("?", 1)[1].split("&")[0]
        token += "=" * (-len(token) % 4)
        decoded = base64.b64decode(token).decode("utf-8", errors="ignore")
        m = re.search(r"file=([^&]+)", decoded)
        if m:
            return os.path.splitext(m.group(1))[1]
    except Exception:
        pass
    return ""


def unduh_dokumen_item(sess, base_url, href_unggah, nama_item, folder_tujuan):
    url_unggah = urljoin(base_url, href_unggah)
    try:
        r = sess.get(url_unggah, timeout=30)
        r.raise_for_status()
    except Exception as e:
        print(f"    [!] Gagal membuka halaman unggah '{nama_item}': {e}")
        return

    soup = BeautifulSoup(r.text, "html.parser")
    link_arsip = [
        a for a in soup.find_all("a", href=True)
        if a["href"].startswith("download.php") and "Buka File" in _clean(a.get_text())
    ]
    if not link_arsip:
        return

    os.makedirs(folder_tujuan, exist_ok=True)
    for idx, a in enumerate(link_arsip, start=1):
        href = a["href"]
        ext = _decode_ekstensi_download(href)
        nama_dasar = _sanitasi_nama_file(nama_item)
        suffix = "" if idx == 1 else f" ({idx})"
        nama_file = f"{nama_dasar}{suffix}{ext}"
        url_file = urljoin(url_unggah, href)
        try:
            rf = sess.get(url_file, timeout=60)
            rf.raise_for_status()
            with open(os.path.join(folder_tujuan, nama_file), "wb") as f:
                f.write(rf.content)
            print(f"    [OK] Unduh dokumen: {nama_file}")
        except Exception as e:
            print(f"    [!] Gagal unduh '{nama_file}': {e}")


def unduh_semua_dokumen_record(sess, base_url, segmen, checklist, tanggal, nama, jenis):
    folder_tujuan = os.path.join(
        FOLDER_UNDUHAN,
        f"{_sanitasi_nama_file(tanggal)}_{_sanitasi_nama_file(nama)}",
    )
    for nama_item, potongan in segmen.items():
        if not checklist.get(nama_item):
            continue
        hrefs = _cari_semua_link_unggah(potongan)
        if not hrefs:
            continue

        if jenis == "Inspeksi" and nama_item == ITEM_SERTIFIKAT and len(hrefs) > 1:
            daftar_sertifikat = _ekstrak_detail_sertifikat(segmen).get("daftar", [])
            for idx, href in enumerate(hrefs):
                if idx < len(daftar_sertifikat) and daftar_sertifikat[idx].get("produk"):
                    label = daftar_sertifikat[idx]["produk"]
                else:
                    label = f"Produk {idx + 1}"
                nama_file_item = f"{nama_item} - {label}"
                unduh_dokumen_item(sess, base_url, href, nama_file_item, folder_tujuan)
        else:
            unduh_dokumen_item(sess, base_url, hrefs[0], nama_item, folder_tujuan)



def kumpulkan_data(driver, daftar_link):
    data_surveilan = []
    data_inspeksi = []
    gagal = []
    total = len(daftar_link)
    sess = buat_session_dari_driver(driver) if UNDUH_DOKUMEN else None
    for i, url in enumerate(daftar_link, start=1):
        print(f"[{i}/{total}] Membuka Page...")
        try:
            html = ambil_html_dengan_retry(driver, url)
            if html is None:
                print(f"  [X] Tetap gagal dibuka setelah {MAX_PERCOBAAN} percobaan, dilewati.")
                gagal.append(url)
                continue

            jenis = deteksi_jenis(html)
            if jenis is None:
                print("  [X] Jenis halaman tidak dikenali (bukan INFO SURVEILAN/INSPEKSI), dilewati.")
                gagal.append(url)
                continue

            info = ekstrak_info(html, jenis)
            segmen = _segmen_dokumentasi(html, jenis)
            checklist = ekstrak_checklist_dokumentasi(segmen, jenis)
            jumlah_lengkap = sum(1 for v in checklist.values() if v)
            total_item_jenis = len(CHECKLIST_ITEMS_PER_JENIS[jenis])

            baris = {
                "nama": info["nama"],
                "tanggal": info["tanggal"],
                "status": info["status"],
                "alamat": info["alamat"],
                "ruang_lingkup": "; ".join(
                    f"{x['nama']} [{x['grade']}]" for x in info["ruang_lingkup"]
                ),
                "kelengkapan_dokumen": f"{jumlah_lengkap}/{total_item_jenis}",
                "url": url,
            }

            for idx, insp in enumerate(info["inspektur"], start=1):
                baris[f"Inspektur_{idx}_Nama"] = insp["nama"]
                baris[f"Inspektur_{idx}_Noreg"] = insp["noreg"]

            for item in CHECKLIST_ITEMS_PER_JENIS[jenis]:
                baris[item] = "Sudah" if checklist.get(item) else "Belum"

            if jenis == "Surveilan":
                detail = _ekstrak_detail_skv(segmen)
                baris["SKV_Nomor"] = detail.get("nomor", "")
                baris["SKV_Tanggal"] = detail.get("tanggal", "")
                if UNDUH_DOKUMEN:
                    unduh_semua_dokumen_record(
                        sess, url, segmen, checklist, info["tanggal"], info["nama"], jenis
                    )
                data_surveilan.append(baris)
            else:
                detail = _ekstrak_detail_sertifikat(segmen)
                baris["Jumlah_Sertifikat"] = len(detail.get("daftar", []))
                baris["Sertifikat_Produk"] = detail.get("produk", "")
                baris["Sertifikat_Nomor"] = detail.get("nomor", "")
                baris["Sertifikat_Batas_Berlaku"] = detail.get("batas", "")
                if UNDUH_DOKUMEN:
                    unduh_semua_dokumen_record(
                        sess, url, segmen, checklist, info["tanggal"], info["nama"], jenis
                    )
                data_inspeksi.append(baris)

            print(f"  -> [{jenis}] {baris['nama']} | {baris['status']} | dokumen {baris['kelengkapan_dokumen']}")
        except Exception as e:
            print(f"  [X] Gagal memproses {url}: {e}")
            gagal.append(url)
    return data_surveilan, data_inspeksi, gagal


def _kolom_inspektur_terbanyak(data):
    maks_inspektur = MIN_KOLOM_INSPEKTUR
    for baris in data:
        for k in baris:
            m = re.match(r"^Inspektur_(\d+)_Nama$", k)
            if m:
                maks_inspektur = max(maks_inspektur, int(m.group(1)))
    kolom = []
    for idx in range(1, maks_inspektur + 1):
        kolom += [f"Inspektur_{idx}_Nama", f"Inspektur_{idx}_Noreg"]
    return kolom


def _lengkapi_kolom(data, fieldnames):
    return [{kol: baris.get(kol, "") for kol in fieldnames} for baris in data]


def simpan_excel(data_surveilan, data_inspeksi, filename):
    if not data_surveilan and not data_inspeksi:
        print("Tidak ada data untuk disimpan.")
        return

    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        kolom_insp_s = _kolom_inspektur_terbanyak(data_surveilan)
        fieldnames_s = (
            ["nama", "tanggal", "alamat"]
            + kolom_insp_s
            + ["ruang_lingkup", "status", "kelengkapan_dokumen", "url"]
            + CHECKLIST_ITEMS_SURVEILAN
            + ["SKV_Nomor", "SKV_Tanggal"]
        )
        df_s = pd.DataFrame(_lengkapi_kolom(data_surveilan, fieldnames_s), columns=fieldnames_s)
        df_s.to_excel(writer, index=False, sheet_name="Surveilan")

        kolom_insp_i = _kolom_inspektur_terbanyak(data_inspeksi)
        fieldnames_i = (
            ["nama", "tanggal", "alamat"]
            + kolom_insp_i
            + ["ruang_lingkup", "status", "kelengkapan_dokumen", "url"]
            + CHECKLIST_ITEMS_INSPEKSI
            + ["Jumlah_Sertifikat", "Sertifikat_Produk", "Sertifikat_Nomor", "Sertifikat_Batas_Berlaku"]
        )
        df_i = pd.DataFrame(_lengkapi_kolom(data_inspeksi, fieldnames_i), columns=fieldnames_i)
        df_i.to_excel(writer, index=False, sheet_name="Inspeksi")

    print(
        f"Tersimpan {len(data_surveilan)} baris Surveilan & "
        f"{len(data_inspeksi)} baris Inspeksi ke {filename}"
    )


def simpan_daftar_gagal(gagal, filename):
    if not gagal:
        return
    with open(filename, "w", encoding="utf-8") as f:
        f.write("\n".join(gagal))
    print(f"\n[!] {len(gagal)} URL gagal dibuka, daftarnya disimpan ke {filename} "
          f"(bisa dicoba ulang nanti).")


def main():
    driver = webdriver.Chrome()
    driver.get(URL_LOGIN)

    input(
        "Login manual & buka halaman list rekap kegiatan di browser, "
        "lalu tekan Enter di sini..."
    )

    html_list = driver.page_source
    daftar_link = ambil_daftar_link(html_list, driver.current_url)
    daftar_link = list(reversed(daftar_link))
    print(f"Ditemukan {len(daftar_link)} data pada halaman list.")

    if not daftar_link:
        print(
            "Tidak ada link detail yang ditemukan. Pastikan Anda sudah berada "
            "di halaman list rekap kegiatan (tabel dengan link detail per baris)."
        )
        driver.quit()
        return

    data_surveilan, data_inspeksi, gagal = kumpulkan_data(driver, daftar_link)
    simpan_excel(data_surveilan, data_inspeksi, OUTPUT_FILE)
    simpan_daftar_gagal(gagal, GAGAL_FILE)
    print(
        f"\n{len(data_surveilan)} baris Surveilan & {len(data_inspeksi)} baris "
        f"Inspeksi tersimpan ke {OUTPUT_FILE}"
    )

    driver.quit()


if __name__ == "__main__":
    main()