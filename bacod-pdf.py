import os
import re
import shutil
import string
import threading
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter import ttk
from PyPDF2 import PdfReader, PdfMerger


# ======================== WINDOW UTAMA ========================
root = Tk()
root.title("J-PDF Toolkit @Dev-Jalanet ")
root.geometry("400x200")


# ======================== FITUR: Rename PDF ========================
def run_rename_tool():
    rename_app = Toplevel(root)
    rename_app.title("Rename PDF: PPh Unifikasi & Faktur Pajak")
    rename_app.geometry("750x550")

    def extract_info_unifikasi_masukan(pdf_path):
        try:
            reader = PdfReader(pdf_path)
            text = "".join(page.extract_text() or "" for page in reader.pages)
            text = re.sub(r'\s+', ' ', text)

            # C.3 Nama Pemotong
            nama_pemotong = re.search(r"A\.2\s+NAMA\s*:\s*(.+?)\s*A\.3", text)

            # No Bukti Potong
            nomor = re.search(r"PEMUNGUTAN PPh PEMUNGUTAN\s+([A-Z0-9]{9})\s+(?:0[1-9]|1[0-2])-\d{4}", text) \
                or re.search(r"BPPU\s+NOMOR MASA PAJAK.*?([A-Z0-9]{9})", text)
            
            # A.2 Nama Dipotong
            nama_match = re.search(r"C\.3 NAMA PEMOTONG DAN/ATAU PEMUNGUT\s*PPh\s*:?.*?([A-Za-z0-9\s,\.\-&;]+?)\s*C\.4", text)

            # B.9 Nomor Dokumen
            nomor_dokumen = re.search(r"B\.9\s+Nomor Dokumen\s*:\s*(.+?)B\.10", text)
            
            # B.5 DPP dan B.7 PPh
            dpp = pph = None
            inline_match = re.search(r"(24-\d{3}-\d{2}).*?([0-9]{1,3}(?:\.\d{3})+|[0-9]+)\s+\d{1,2}\s+([0-9]{1,3}(?:\.\d{3})+|[0-9]+)", text)

            if inline_match:
                dpp = inline_match.group(2).replace(".", "")
                pph = inline_match.group(3).replace(".", "")
            else:
                words = text.split()
                for i in range(len(words) - 6):
                    if re.match(r"24-\d{3}-\d{2}", words[i]):
                        try:
                            dpp_candidate = words[i + 4].replace(".", "")
                            pph_candidate = words[i + 6].replace(".", "")
                            if dpp_candidate.isdigit() and pph_candidate.isdigit():
                                dpp = dpp_candidate
                                pph = pph_candidate
                                break
                        except IndexError:
                            continue

            if nama_match and nomor and dpp and pph and nama_pemotong and nomor_dokumen:
                raw_nama = nama_match.group(1).strip()
                c3_nama = re.sub(r"[^\w\s-]", "", nama_pemotong.group(1).strip().replace(" ", "_"))
                a2_nama = raw_nama.replace("&amp;", "&").replace("  ", " ")
                dokumen = re.sub(r"[^\w\s-]", "", nomor_dokumen.group(1).strip().replace(" ", "_"))
                return {
                    "nama_pemotong": c3_nama,
                    "no_bukti": nomor.group(1).strip(),
                    "dpp": dpp,
                    "pph": pph,
                    "nama_dipotong": a2_nama,
                    "dokumen": dokumen
                }
        except:
            return None

    def extract_info_unifikasi(pdf_path):
        try:
            reader = PdfReader(pdf_path)
            text = "".join(page.extract_text() or "" for page in reader.pages)
            text = text.replace('\n', ' ')
            text = re.sub(r'\s+', ' ', text)

            # Ekstraksi data dari format PDF
            nama_pemotong = re.search(r"C\.3 NAMA PEMOTONG DAN/ATAU PEMUNGUT\s*PPh\s*:?.*?([A-Za-z0-9\s,\.\-&;]+?)\s*C\.4", text)
            nomor_bukti = re.search(r"(?:PEMUNGUTAN PPh PEMUNGUTAN|BPPU)\s+([A-Z0-9]{9})", text)
            dpp_pph = re.search(r"(28-\d{3}-\d{2}).*?([\d\.]+)\s+[0-9]{1,2}(\.\d+)?\s+([\d\.]+)", text)
            nama_pemotong_dipotong = re.search(r"A\.2\s+NAMA\s*:\s*(.+?)\s*A\.3", text)
            nomor_dokumen = re.search(r"B\.9\s+Nomor Dokumen\s*:\s*(.+?)B\.10", text)

            if nama_pemotong and nomor_bukti and dpp_pph and nama_pemotong_dipotong and nomor_dokumen:
                c3_nama = re.sub(r"[^\w\s-]", "", nama_pemotong.group(1).strip().replace(" ", "_"))
                a2_nama = re.sub(r"[^\w\s-]", "", nama_pemotong_dipotong.group(1).strip().replace(" ", "_"))
                dokumen = re.sub(r"[^\w\s-]", "", nomor_dokumen.group(1).strip().replace(" ", "_"))
                return {
                    "nama_pemotong": c3_nama,
                    "no_bukti": nomor_bukti.group(1).strip(),
                    "dpp": dpp_pph.group(2).replace(".", ""),
                    "pph": dpp_pph.group(4).replace(".", ""),
                    "nama_dipotong": a2_nama,
                    "dokumen": dokumen
                }
            else:
                print("[DEBUG] Gagal ekstrak satu atau lebih elemen.")
                return None

        except Exception as e:
            print(f"[ERROR] Gagal parsing PDF Bupot Keluaran: {e}")
            return None          
    


    def extract_info_faktur(pdf_path):

        
        try:
            reader = PdfReader(pdf_path)
            text = ""
            for page in reader.pages:
                text += page.extract_text() or ""

                text = text.replace('\n', ' ')
                invoice_match = re.search(r'(TGR/\S+)', text)
                pembeli_match = re.search(r'Pembeli Barang.*?Nama\s*:\s*([A-Z0-9\s&,.]+)', text, re.DOTALL)

            if invoice_match and pembeli_match:
                invoice_raw = invoice_match.group(1).replace("/", "")
                pembeli_raw = pembeli_match.group(1).strip()
                pembeli_clean = pembeli_raw.splitlines()[0]
                pembeli_clean = re.sub(r'[\\/*?:"<>|]', "", pembeli_clean).strip()
                pembeli_clean = re.sub(r'\bA$', '', pembeli_clean).strip()
                pembeli_nama = pembeli_clean.title()
                return invoice_raw, pembeli_nama
            else:
                return None
        
        except Exception as e:
            print(f"[ERROR] Faktur: {pdf_path} - {e}")
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

        renamed = skipped = 0
        progress_bar["value"] = 0
        progress_bar["maximum"] = total
        progress_label.config(text=f"Memproses: 0 / {total}")
        rename_app.update_idletasks()

        for idx, filename in enumerate(pdf_files, start=1):
            full_path = os.path.join(folder, filename)
            new_name = None

            if mode_var.get() == 1:
                info = extract_info_unifikasi_masukan(full_path)
                if info:
                    # new_name = f"{info['nama_pemotong']}_{info['no_bukti']}_{info['dpp']}_{info['pph']}_{info['nama_dipotong']}_{info['dokumen']}.pdf"
                    nama = re.sub(r"[^\w\s-]", "", info['nama_dipotong'].upper().replace(" ", "_"))
                    new_name = f"{nama}_{info['no_bukti']}_{info['dpp']}_{info['pph']}_{info['nama_pemotong']}_{info['dokumen']}.pdf"
            if mode_var.get() == 3:
                info = extract_info_unifikasi(full_path)
                if info:
                    new_name = f"{info['nama_pemotong']}_{info['no_bukti']}_{info['dpp']}_{info['pph']}_{info['nama_dipotong']}_{info['dokumen']}.pdf"
                    # nama = re.sub(r"[^\w\s-]", "", info['nama'].upper().replace(" ", "_"))
                    # new_name = f"{nama}_{info['nomor']}_{info['dpp']}_{info['pph']}.pdf"
            else:
                info = extract_info_faktur(full_path)
                if info:
                    invoice, pembeli = info
                    new_name = f"{invoice}_{pembeli}.pdf"

            if new_name:
                new_path = os.path.join(folder, new_name)
                base, ext = os.path.splitext(new_name)
                counter = 1
                while os.path.exists(new_path):
                    new_name = f"{base}_{counter}.pdf"
                    new_path = os.path.join(folder, new_name)
                    counter += 1
                os.rename(full_path, new_path)
                listbox.insert(END, f"✔ {filename} → {new_name}")
                renamed += 1
            else:
                listbox.insert(END, f"✘ Gagal ekstrak: {filename}")
                skipped += 1

            progress_bar["value"] = idx
            progress_label.config(text=f"Memproses: {idx} / {total}")
            listbox.yview_moveto(1.0)
            rename_app.update_idletasks()

        progress_label.config(text="Selesai.")
        messagebox.showinfo("Selesai", f"{renamed} file berhasil dinamai ulang.\n{skipped} file gagal diproses.")

    Label(rename_app, text="Pilih Jenis Dokumen:", font=("Arial", 12)).pack(pady=10)
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


# ======================== FITUR: Merge PDF ========================
def run_merge_tool():
    merge_app = Toplevel(root)
    merge_app.title("Gabungkan PDF")
    merge_app.geometry("700x580")

    selected_pdfs = []
    progress_var = IntVar()
    status_var = StringVar()

    def select_files():
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        for path in file_paths:
            if path not in selected_pdfs:
                selected_pdfs.append(path)
                pdf_listbox.insert(END, os.path.basename(path))

    def select_folder():
        folder_path = filedialog.askdirectory()
        if not folder_path:
            return
        pdfs = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
        for path in sorted(pdfs):
            if path not in selected_pdfs:
                selected_pdfs.append(path)
                pdf_listbox.insert(END, os.path.basename(path))

    def remove_selected():
        selected_indices = pdf_listbox.curselection()
        for index in reversed(selected_indices):
            selected_pdfs.pop(index)
            pdf_listbox.delete(index)

    def clear_list():
        selected_pdfs.clear()
        pdf_listbox.delete(0, END)
        progress_var.set(0)
        status_var.set("")

    def merge_pdfs_thread(output_path):
        try:
            merger = PdfMerger()
            total = len(selected_pdfs)
            for i, pdf_path in enumerate(selected_pdfs, start=1):
                merger.append(pdf_path)
                progress_var.set(int(i / total * 100))
                status_var.set(f"Menggabungkan file {i} / {total}...")
                merge_app.update_idletasks()
            merger.write(output_path)
            merger.close()
            messagebox.showinfo("Sukses", f"Berhasil digabungkan ke:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Gagal", str(e))
        finally:
            progress_var.set(0)
            status_var.set("")

    def merge_pdfs():
        if not selected_pdfs:
            messagebox.showwarning("Peringatan", "Silakan pilih file PDF terlebih dahulu.")
            return
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if output_path:
            threading.Thread(target=merge_pdfs_thread, args=(output_path,), daemon=True).start()

    Button(merge_app, text="Pilih File PDF", command=select_files, width=20).pack(pady=3)
    Button(merge_app, text="Pilih Folder PDF", command=select_folder, width=20).pack(pady=3)
    Button(merge_app, text="Hapus File Terpilih", command=remove_selected, width=20).pack(pady=3)
    Button(merge_app, text="Bersihkan Daftar", command=clear_list, width=20).pack(pady=3)

    frame = Frame(merge_app)
    frame.pack(pady=10, fill="both", expand=True)
    scrollbar = Scrollbar(frame)
    scrollbar.pack(side="right", fill="y")
    pdf_listbox = Listbox(frame, selectmode=MULTIPLE, width=90, height=18, yscrollcommand=scrollbar.set)
    pdf_listbox.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=pdf_listbox.yview)

    progress_bar = ttk.Progressbar(merge_app, orient="horizontal", length=600, mode="determinate", variable=progress_var)
    progress_bar.pack(pady=5)
    status_label = Label(merge_app, textvariable=status_var, fg="blue")
    status_label.pack()

    Button(merge_app, text="Gabungkan PDF", command=merge_pdfs, width=20).pack(pady=10)


# ======================== MENU ========================
menu_bar = Menu(root)
file_menu = Menu(menu_bar, tearoff=0)
file_menu.add_command(label="Rename PDF", command=run_rename_tool)
file_menu.add_command(label="Merge PDF", command=run_merge_tool)
file_menu.add_separator()
file_menu.add_command(label="Keluar", command=root.quit)
menu_bar.add_cascade(label="File", menu=file_menu)
root.config(menu=menu_bar)

Label(root, text="Selamat datang di J-PDF Toolkit", font=("Arial", 14)).pack(pady=60)
Label(root, text="Gunakan menu File untuk memulai", font=("Arial", 10)).pack()

root.mainloop()