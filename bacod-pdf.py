# J-PDF Toolkit Final Version - Rename Bupot Masukan/Keluaran + Faktur + Logging + CSV

import os
import re
import csv
import threading
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter import ttk
from PyPDF2 import PdfReader

root = Tk()
root.title("J-PDF Toolkit @Dev-Jalanet FINAL")
root.geometry("400x200")

def log_rename_result(folder, log_data):
    
    log_path = os.path.join(folder, "rename_log.txt")
    with open(log_path, "w", encoding="utf-8") as f:
        for line in log_data:
            f.write(line + "\n")

    csv_path = os.path.join(folder, "rename_result.csv")
    with open(csv_path, "w", newline='', encoding="utf-8") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["Original Filename", "New Filename"])
        for line in log_data:
            if line.startswith("✔"):
                parts = line[2:].split(" → ")
                if len(parts) == 2:
                    writer.writerow(parts)

def run_rename_tool():
    print("Start rename result")
    rename_app = Toplevel(root)
    rename_app.title("Rename PDF: PPh Unifikasi & Faktur Pajak")
    rename_app.geometry("750x550")

    def extract_info_masukan(pdf_path):
        print("Start extract_info_masukan")
        try:
            reader = PdfReader(pdf_path)
            text = "".join(page.extract_text() or "" for page in reader.pages)
            text = re.sub(r"\s+", " ", text)

            a2_nama = re.search(r"A\.2\s+NAMA\s*:\s*(.+?)\s*A\.3", text)
            no_bukti = re.search(r"PEMUNGUTAN PPh PEMUNGUTAN\s+([A-Z0-9]{9})", text) or \
                       re.search(r"BPPU\s+NOMOR MASA PAJAK.*?([A-Z0-9]{9})", text)
            c3_nama = re.search(r"C\.3 NAMA PEMOTONG DAN/ATAU PEMUNGUT\s*PPh\s*:?.*?(.+?)\s*C\.4", text)
            dokumen = re.search(r"B\.9\s+Nomor Dokumen\s*:\s*(.+?)B\.10", text)
            dpp_pph = re.search(r"(24-\d{3}-\d{2}).*?([\d.]+)\s+\d{1,2}(?:\.\d+)?\s+([\d.]+)", text)

            if not all([a2_nama, c3_nama, no_bukti, dpp_pph, dokumen]):
                raise ValueError("Field tidak lengkap: "
                             f"A2={bool(a2_nama)}, C3={bool(c3_nama)}, Bukti={bool(no_bukti)}, "
                             f"DPP/PPH={bool(dpp_pph)}, Dokumen={bool(dokumen)}")
            
            a2_raw = a2_nama.group(1).strip()
            a2_clean = re.split(r"\s{2,}|\s*A\.3", a2_raw)[0]
            a2_final = re.sub(r"[^\w\s-]", "", a2_clean).replace(" ", "_")
            c3_final = re.sub(r"[^\w\s-]", "", c3_nama.group(1).strip()).replace(" ", "_")
            dokumen_final = re.sub(r"[^\w\s-]", "", dokumen.group(1).strip()).replace(" ", "_")
            return {
                    "nama_dipotong": a2_final,
                    "no_bukti": no_bukti.group(1).strip(),
                    "dpp": dpp_pph.group(2).replace(".", ""),
                    "pph": dpp_pph.group(3).replace(".", ""),
                    "nama_pemotong": c3_final,
                    "dokumen": dokumen_final
                }
        except Exception as e:
        # Raise the exception info so we can use it in the caller
            raise RuntimeError(f"{os.path.basename(pdf_path)} → {str(e)}")

    def extract_info_keluaran(pdf_path):
        try:
            reader = PdfReader(pdf_path)
            text = "".join(page.extract_text() or "" for page in reader.pages)
            text = text.replace('\n', ' ')
            text = re.sub(r'\s+', ' ', text)

            nama_pemotong = re.search(r"C\.3 NAMA PEMOTONG DAN/ATAU PEMUNGUT\s*PPh\s*:?.*?(.+?)\s*C\.4", text)
            nomor_bukti = re.search(r"(?:PEMUNGUTAN PPh PEMUNGUTAN|BPPU)\s+([A-Z0-9]{9})", text)
            dpp_pph = re.search(r"(28-\d{3}-\d{2}).*?([\d.]+)\s+\d{1,2}(?:\.\d+)?\s+([\d.]+)", text)
            nama_dipotong = re.search(r"A\.2\s+NAMA\s*:\s*(.+?)\s*A\.3", text)
            nomor_dokumen = re.search(r"B\.9\s+Nomor Dokumen\s*:\s*(.+?)B\.10", text)

            if all([nama_pemotong, nomor_bukti, dpp_pph, nama_dipotong, nomor_dokumen]):
                c3_nama = re.sub(r"[^\w\s-]", "", nama_pemotong.group(1).strip()).replace(" ", "_")
                a2_raw = nama_dipotong.group(1).strip()
                a2_clean = re.split(r"\s{2,}|\s*A\.3", a2_raw)[0]
                a2_nama = re.sub(r"[^\w\s-]", "", a2_clean).replace(" ", "_")
                dokumen = re.sub(r"[^\w\s-]", "", nomor_dokumen.group(1).strip()).replace(" ", "_")
                return {
                    "nama_pemotong": c3_nama,
                    "no_bukti": nomor_bukti.group(1).strip(),
                    "dpp": dpp_pph.group(2).replace(".", ""),
                    "pph": dpp_pph.group(3).replace(".", ""),
                    "nama_dipotong": a2_nama,
                    "dokumen": dokumen
                }
        except:
            return None

    def extract_info_faktur(pdf_path):
        try:
            reader = PdfReader(pdf_path)
            text = "".join(page.extract_text() or "" for page in reader.pages)
            text = text.replace('\n', ' ')
            invoice_match = re.search(r'(TGR/\S+)', text)
            pembeli_match = re.search(r'Pembeli Barang.*?Nama\s*:\s*([A-Z0-9\s&,.]+)', text, re.DOTALL)

            if invoice_match and pembeli_match:
                invoice_raw = invoice_match.group(1).replace("/", "")
                pembeli_clean = re.sub(r'[\\/*?:"<>|]', '', pembeli_match.group(1).strip()).splitlines()[0].strip()
                pembeli_nama = re.sub(r'\s+', '_', pembeli_clean.title())
                return invoice_raw, pembeli_nama
        except:
            return None

    def process_files():
        folder = filedialog.askdirectory()
        if not folder:
            return

        listbox.delete(0, END)
        pdf_files = [f for f in os.listdir(folder) if f.lower().endswith(".pdf")]
        total = len(pdf_files)
        if total == 0:
            messagebox.showinfo("Info", "Tidak ada file PDF di folder ini.")
            return

        log_data = []
        renamed = skipped = 0
        progress_bar["value"] = 0
        progress_bar["maximum"] = total
        progress_label.config(text=f"Memproses: 0 / {total}")
        rename_app.update_idletasks()

        for idx, filename in enumerate(pdf_files, start=1):
            full_path = os.path.join(folder, filename)
            new_name = None

            if mode_var.get() == 1:
                try:
                    info = extract_info_masukan(full_path)
                    if info:
                        new_name = f"{info['nama_dipotong']}_{info['no_bukti']}_{info['dpp']}_{info['pph']}_{info['nama_pemotong']}_{info['dokumen']}.pdf"
                                            
                    else:
                        raise ValueError("Parser mengembalikan None (data tidak lengkap)")
                        
                except Exception as err_msg:
                    log_line = f"✘ Gagal ekstrak: {filename} → {err_msg}"
                    listbox.insert(END, log_line)
                    messagebox.showwarning("Gagal Ekstrak Bupot Masukan", str(err_msg))
                    skipped += 1
                    progress_bar["value"] = idx
                    progress_label.config(text=f"Memproses: {idx} / {total}")
                    rename_app.update_idletasks()
                    continue                           
            elif mode_var.get() == 3:
                info = extract_info_keluaran(full_path)
                if info:
                    new_name = f"{info['nama_pemotong']}_{info['no_bukti']}_{info['dpp']}_{info['pph']}_{info['nama_dipotong']}_{info['dokumen']}.pdf"
            elif mode_var.get() == 2:
                info = extract_info_faktur(full_path)
                if info:
                    invoice, pembeli = info
                    new_name = f"{invoice}_{pembeli}.pdf"

            if new_name:
                try:
                    new_path = os.path.join(folder, new_name)
                    base, ext = os.path.splitext(new_name)
                    counter = 1

                    # Batas nama file maksimal (umumnya 255 karakter untuk nama file saja)
                    max_filename_length = 240  # Aman dengan buffer

                    # Potong jika nama terlalu panjang
                    if len(new_name) > max_filename_length:
                        base = base[:max_filename_length - len(ext) - 5]  # kurangi space untuk _1.pdf dsb
                        new_name = f"{base}{ext}"
                        new_path = os.path.join(folder, new_name)

                    # Cegah overwrite file
                    while os.path.exists(new_path):
                        new_name = f"{base}_{counter}{ext}"
                        new_path = os.path.join(folder, new_name)
                        counter += 1

                    os.rename(full_path, new_path)

                    log_line = f"✔ {filename} → {new_name}"
                    listbox.insert(END, log_line)
                    log_data.append(log_line)
                    renamed += 1

                except Exception as e:
                    log_line = f"✘ Rename gagal: {filename} → {e}"
                    listbox.insert(END, log_line)
                    log_data.append(log_line)
                    skipped += 1
            else:
                log_line = f"✘ Gagal ekstrak: {filename}"
                listbox.insert(END, log_line)
                log_data.append(log_line)
                skipped += 1

            progress_bar["value"] = idx
            progress_label.config(text=f"Memproses: {idx} / {total}")
            listbox.yview_moveto(1.0)
            rename_app.update_idletasks()

        progress_label.config(text="Selesai.")
        log_rename_result(folder, log_data)
        messagebox.showinfo("Selesai", f"{renamed} berhasil, {skipped} gagal.\nLog disimpan di folder.")

    Label(rename_app, text="Pilih jenis dokumen:", font=("Arial", 12)).pack(pady=10)
    mode_var = IntVar(value=1)
    Radiobutton(rename_app, text="🔹 Bupot Masukan", variable=mode_var, value=1).pack()
    Radiobutton(rename_app, text="🔹 Bupot Keluaran", variable=mode_var, value=3).pack()
    Radiobutton(rename_app, text="🔹 Faktur Pajak", variable=mode_var, value=2).pack()

    Button(rename_app, text="Pilih Folder PDF", command=process_files, width=20).pack(pady=10)
    progress_bar = ttk.Progressbar(rename_app, length=600, mode='determinate')
    progress_bar.pack(pady=5)
    progress_label = Label(rename_app, text="Menunggu proses...")
    progress_label.pack()
    scrollbar = Scrollbar(rename_app)
    scrollbar.pack(side="right", fill="y")
    listbox = Listbox(rename_app, width=100, height=20, yscrollcommand=scrollbar.set)
    listbox.pack(pady=10)
    scrollbar.config(command=listbox.yview)

menu_bar = Menu(root)
file_menu = Menu(menu_bar, tearoff=0)
file_menu.add_command(label="Rename PDF", command=run_rename_tool)
file_menu.add_separator()
file_menu.add_command(label="Keluar", command=root.quit)
menu_bar.add_cascade(label="File", menu=file_menu)
root.config(menu=menu_bar)

Label(root, text="Selamat datang di J-PDF Toolkit", font=("Arial", 14)).pack(pady=60)
Label(root, text="Gunakan menu File untuk memulai", font=("Arial", 10)).pack()

root.mainloop()
