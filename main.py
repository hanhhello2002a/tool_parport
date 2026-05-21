import os
import json
import io
import sys
import threading
import urllib.request
import random
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client
import pythoncom  
import customtkinter as ctk  

# ==================== ĐỒNG BỘ VERSION CHUẨN (KHÔNG ĐỂ LỆCH) ====================
CURRENT_VERSION = "4.1"  # Đảm bảo số này khớp chính xác với số trên version.json của GitHub
VERSION_URL = "https://raw.githubusercontent.com/hanhhello2002a/tool_parport/refs/heads/main/version.json"
RAW_REPO_URL = "https://raw.githubusercontent.com/hanhhello2002a/tool_parport/refs/heads/main/"
# ===============================================================================

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("green")

class PassportToolGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"⚡ UK PASSPORT - PROTOCOL TOOL v{CURRENT_VERSION}")
        self.geometry("950x800")
        self.psd_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "design.psd")
        self.output_dir_var = tk.StringVar()
        self.prefix_num_var = tk.StringVar()
        self.suffix_var = tk.StringVar()
        self.load_config()
        self.placeholder_sample = (
            "hovaten,ho,ho2,ten,ten2,so,so2,ntns,ntns2,ma1,ma1.2,ma2,ma2.2\n"
            '"HOWARD GABRIEL","HOWARD","HOWARD","GABRIEL","GABRIEL","640891663","640891663",'
            '"21 DEC /DEC 97","21 DEC /DEC 97","P<GBRHOWARD<<GABRIEL<<<<<<<<<<<<<<<<<<<<<<<<",'
            '"P<GBRHOWARD<<GABRIEL<<<<<<<<<<<<<<<<<<<<<<<<","6408916630GBR9712211M3503029<<<<<<<<<<<<<<04",'
            '"6408916630GBR9712211M3503029<<<<<<<<<<<<<<04"'
        )
        self.tab_view = ctk.CTkTabview(self)
        self.tab_view.pack(fill="both", expand=True, padx=20, pady=15)
        try:
            self.tab_view._segmented_button.configure(
                fg_color="#141414", selected_fg_color="#22c55e",
                selected_text_color="#000000", unselected_text_color="#22c55e"
            )
        except: pass
        self.tab_process = self.tab_view.add("  🚀 XỬ LÝ PHOTOSHOP  ")
        self.tab_generate = self.tab_view.add("  ⚙️ CẤU HÌNH TÙY CHỈNH  ")
        self.tab_gen_perfect = self.tab_view.add("  🎲 RANDOM REAL 100% (MRZ)  ")
        self.init_tab_process()
        self.init_tab_generate()
        self.init_tab_gen_perfect()
        self.prefix_num_var.trace_add("write", lambda *args: self.save_config())
        self.suffix_var.trace_add("write", lambda *args: self.save_config())
        self.output_dir_var.trace_add("write", lambda *args: self.save_config())
        self.remote_config_data = {} 
        self.after(200, self.check_for_updates)

    # 🟢 SỬA LẠI: ÉP XÓA CACHE INTERNET TUYỆT ĐỐI KHI CHECK VERSION 🟢
    def check_for_updates(self):
        def async_check():
            try:
                # Tạo chuỗi random siêu dài để đánh lừa mọi hệ thống cache mạng
                bypass_url = f"{VERSION_URL}?nocache={random.randint(1000000, 9999999)}"
                req = urllib.request.Request(bypass_url)
                
                # Bổ sung các Header ép buộc Header không lưu đệm file cũ
                req.add_header('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)')
                req.add_header('Cache-Control', 'no-cache, no-store, must-revalidate')
                req.add_header('Pragma', 'no-cache')
                req.add_header('Expires', '0')
                
                with urllib.request.urlopen(req, timeout=7) as response:
                    self.remote_config_data = json.loads(response.read().decode('utf-8'))
                
                remote_version = str(self.remote_config_data.get("version", CURRENT_VERSION)).strip()
                local_version = str(CURRENT_VERSION).strip()
                
                # Chỉ hiện bảng update nếu bản trên mạng thực sự lớn hơn bản ở máy
                if float(remote_version) > float(local_version):
                    self.show_update_overlay(remote_version)
            except:
                pass
        threading.Thread(target=async_check, daemon=True).start()

    def show_update_overlay(self, new_version):
        self.tab_view.pack_forget() 
        self.overlay = ctk.CTkFrame(self, fg_color="#0a0a0a")
        self.overlay.pack(fill="both", expand=True)
        lbl_title = ctk.CTkLabel(self.overlay, text="🚨 CẬP NHẬT BẮT BUỘC: PHÔI PSD & CẤU HÌNH MỚI", font=("Consolas", 15, "bold"), text_color="#ef4444")
        lbl_title.pack(pady=(220, 10))
        lbl_info = ctk.CTkLabel(self.overlay, text=f"Hệ thống yêu cầu đồng bộ cấu hình sang bản: v{new_version}\nTool sẽ tự động giải phóng Photoshop, tải Phôi Thiết Kế và ghi đè chuỗi MA2 mặc định mới.", font=("Consolas", 12), text_color="#a3a3a3")
        lbl_info.pack(pady=10)
        self.progress_lbl = ctk.CTkLabel(self.overlay, text="Trạng thái: Sẵn sàng đồng bộ...", font=("Consolas", 12, "italic"), text_color="#eab308")
        self.progress_lbl.pack(pady=5)
        self.btn_update = ctk.CTkButton(self.overlay, text="📥 ĐỒNG BỘ PHÔI & CONFIG NGAY", font=("Consolas", 14, "bold"), height=45, fg_color="#22c55e", hover_color="#16a34a", text_color="#000000", command=lambda: threading.Thread(target=self.execute_resource_update).start())
        self.btn_update.pack(pady=20, ipadx=20)

    def execute_resource_update(self):
        self.btn_update.configure(state="disabled", text="ĐANG ĐỒNG BỘ...")
        try:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            self.progress_lbl.configure(text="⏳ Đang đóng ứng dụng Photoshop ngầm để mở khóa file...", text_color="#eab308")
            os.system("taskkill /f /im Photoshop.exe >nul 2>&1")
            
            self.progress_lbl.configure(text="⏳ Đang tải phôi thiết kế design.psd mới từ Server...", text_color="#06b6d4")
            psd_url = f"{RAW_REPO_URL}design.psd?nocache={random.randint(100000, 999999)}"
            target_psd_path = os.path.join(current_dir, "design.psd")
            
            req = urllib.request.Request(psd_url)
            req.add_header('User-Agent', 'Mozilla/5.0')
            req.add_header('Cache-Control', 'no-cache')
            with urllib.request.urlopen(req) as response:
                with open(target_psd_path, "wb") as f:
                    f.write(response.read())
                    
            self.progress_lbl.configure(text="⚙️ Đang đồng bộ cấu hình chuỗi MA2 mặc định mới...", text_color="#06b6d4")
            new_prefix = self.remote_config_data.get("default_prefix", "0")
            new_suffix = self.remote_config_data.get("default_suffix", "1M3503029<<<<<<<<<<<<<<04")
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump({
                    "prefix_num": new_prefix,
                    "suffix": new_suffix,
                    "output_dir": self.output_dir_var.get().strip()
                }, f, ensure_ascii=False, indent=4)
            self.prefix_num_var.set(new_prefix)
            self.suffix_var.set(new_suffix)
            self.progress_lbl.configure(text="✅ Đồng bộ thành công!", text_color="#22c55e")
            messagebox.showinfo("Thành công", "Đã cập nhật Phôi thiết kế PSD và cấu hình chuỗi MA2 mới thành công!\nỨng dụng sẽ tự động mở lại giao diện.")
            self.overlay.pack_forget()
            self.tab_view.pack(fill="both", expand=True, padx=20, pady=15)
            self.update_status("Đã ghi đè trực tiếp file design.psd mới nhất thành công!", "#22c55e")
        except Exception as e:
            messagebox.showerror("Lỗi Đồng Bộ", f"Không thể tải đè phôi thiết kế.\nChi tiết: {e}")
            self.btn_update.configure(state="normal", text="THỬ LẠI ĐỒNG BỘ")

    def load_config(self):
        default_dir = os.path.dirname(os.path.abspath(__file__))
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.prefix_num_var.set(data.get("prefix_num", "0"))
                    self.suffix_var.set(data.get("suffix", "1M3503029<<<<<<<<<<<<<<04"))
                    saved_dir = data.get("output_dir", default_dir)
                    if os.path.exists(saved_dir):
                        self.output_dir_var.set(os.path.normpath(saved_dir))
                    else:
                        self.output_dir_var.set(default_dir)
                    return
            except: pass
        self.prefix_num_var.set("0")
        self.suffix_var.set("1M3503029<<<<<<<<<<<<<<04")
        self.output_dir_var.set(default_dir)

    def save_config(self):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump({
                    "prefix_num": self.prefix_num_var.get().strip(),
                    "suffix": self.suffix_var.get().strip(),
                    "output_dir": self.output_dir_var.get().strip()
                }, f, ensure_ascii=False, indent=4)
        except: pass

    def init_tab_process(self):
        frame_folder = ctk.CTkFrame(self.tab_process, fg_color="#141414", corner_radius=10)
        frame_folder.pack(fill="x", padx=15, pady=10)
        lbl_folder = ctk.CTkLabel(frame_folder, text="📁 THƯ MỤC LƯU FILE ẢNH (PNG):", font=("Consolas", 12, "bold"), text_color="#22c55e")
        lbl_folder.pack(anchor="w", padx=15, pady=(10, 2))
        sub_frame = ctk.CTkFrame(frame_folder, fg_color="transparent")
        sub_frame.pack(fill="x", padx=10, pady=(0, 10))
        self.ent_folder = ctk.CTkEntry(sub_frame, textvariable=self.output_dir_var, font=("Consolas", 12))
        self.ent_folder.pack(side="left", padx=5, fill="x", expand=True)
        btn_browse = ctk.CTkButton(sub_frame, text="BROWSE FOLDER", font=("Consolas", 12, "bold"), width=130, fg_color="#2b2b2b", hover_color="#3a3a3a", text_color="#22c55e", command=self.browse_output_folder)
        btn_browse.pack(side="right", padx=5)
        frame_data = ctk.CTkFrame(self.tab_process, fg_color="#141414", corner_radius=10)
        frame_data.pack(fill="both", expand=True, padx=15, pady=5)
        lbl_data = ctk.CTkLabel(frame_data, text="📝 DỮ LIỆU ĐẦU VÀO ĐỊNH DẠNG CSV:", font=("Consolas", 12, "bold"), text_color="#22c55e")
        lbl_data.pack(anchor="w", padx=15, pady=(10, 2))
        self.txt_data = ctk.CTkTextbox(frame_data, font=("Consolas", 12), fg_color="#090909", border_width=1, border_color="#222222")
        self.txt_data.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.txt_data.insert("1.0", self.placeholder_sample)
        self.txt_data.configure(text_color="#555555")
        self.txt_data.bind("<FocusIn>", self.on_textbox_focus_in)
        self.txt_data.bind("<FocusOut>", self.on_textbox_focus_out)
        self.lbl_status = ctk.CTkLabel(self.tab_process, text="[SYSTEM LOG]: Ready.", font=("Consolas", 12, "italic"), text_color="#a3a3a3")
        self.lbl_status.pack(anchor="w", padx=20, pady=5)
        self.btn_run = ctk.CTkButton(self.tab_process, text="⚡ BẮT ĐẦU AUTO RUN PHOTOSHOP", font=("Consolas", 14, "bold"), height=45, fg_color="#22c55e", hover_color="#16a34a", text_color="#000000", command=self.start_processing_thread)
        self.btn_run.pack(fill="x", padx=15, pady=10)

    def on_textbox_focus_in(self, event):
        current_text = self.txt_data.get("1.0", "end-1c").strip()
        if current_text == self.placeholder_sample.strip():
            self.txt_data.delete("1.0", "end")
            self.txt_data.configure(text_color="#22c55e")

    def on_textbox_focus_out(self, event):
        current_text = self.txt_data.get("1.0", "end-1c").strip()
        if not current_text:
            self.txt_data.insert("1.0", self.placeholder_sample)
            self.txt_data.configure(text_color="#555555")

    def browse_output_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.output_dir_var.set(os.path.normpath(folder_selected))
            
    def update_status(self, text, color="#22c55e"):
        self.lbl_status.configure(text=f"[SYSTEM LOG]: {text}", text_color=color)
        self.update_idletasks()

    def start_processing_thread(self):
        t = threading.Thread(target=self.process_photoshop)
        t.start()

    def process_photoshop(self):
        pythoncom.CoInitialize()
        raw_text = self.txt_data.get("1.0", "end-1c").strip()
        base_dir = self.output_dir_var.get().strip()
        self.psd_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "design.psd")
        if not os.path.exists(self.psd_path):
            messagebox.showerror("Error", f"Không tìm thấy file 'design.psd' tại:\n{self.psd_path}")
            pythoncom.CoUninitialize()
            return
        if not base_dir or not os.path.exists(base_dir):
            messagebox.showerror("Error", "Thư mục chọn lưu ảnh không tồn tại!")
            pythoncom.CoUninitialize()
            return
        if not raw_text or raw_text == self.placeholder_sample.strip():
            messagebox.showerror("Error", "Vui lòng nhập dữ liệu CSV thực tế trước khi chạy!")
            pythoncom.CoUninitialize()
            return
        output_dir = os.path.join(base_dir, "PNG")
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        self.btn_run.configure(state="disabled", fg_color="#333333", text_color="#888888")
        self.update_status("Đang phân tích bảng dữ liệu CSV...", "#eab308")
        try:
            df = pd.read_csv(io.StringIO(raw_text))
        except Exception as e:
            messagebox.showerror("CSV Error", f"Định dạng CSV không hợp lệ!\nChi tiết: {e}")
            self.btn_run.configure(state="normal", fg_color="#22c55e", text_color="#000000")
            self.update_status("Ready", "#22c55e")
            pythoncom.CoUninitialize()
            return
        self.update_status("Đang mở ứng dụng Photoshop ngầm...", "#eab308")
        ps_app = None
        try:
            ps_app = win32com.client.Dispatch("Photoshop.Application")
            ps_app.Visible = False
        except Exception as e:
            messagebox.showerror("Photoshop Error", f"Không thể kết nối Photoshop.\nChi tiết: {e}")
            self.btn_run.configure(state="normal", fg_color="#22c55e", text_color="#000000")
            self.update_status("Ready", "#22c55e")
            pythoncom.CoUninitialize()
            return
        try:
            doc = ps_app.Open(self.psd_path)
            total_rows = len(df)
            for index, row in df.iterrows():
                hovaten_raw = str(row.get('hovaten', f'Unknow {index}')).strip()
                file_name = hovaten_raw
                for char in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']:
                    file_name = file_name.replace(char, '')
                save_path = os.path.join(output_dir, f"{file_name}.png")
                counter = 1
                while os.path.exists(save_path):
                    save_path = os.path.join(output_dir, f"{file_name} ({counter}).png")
                    counter += 1
                self.update_status(f"Render [{index + 1}/{total_rows}]: {hovaten_raw}...", "#06b6d4")
                layer_data = row.to_dict()
                layer_data_str = {str(k).strip(): str(v).strip() for k, v in layer_data.items()}
                jsx_script = """
                function updateLayersRecursive(layers, rowData) {
                    for (var i = 0; i < layers.length; i++) {
                        var layer = layers[i];
                        if (layer.typename === "LayerSet") {
                            updateLayersRecursive(layer.layers, rowData);
                        } else if (layer.kind === LayerKind.TEXT) {
                            if (rowData.hasOwnProperty(layer.name)) {
                                layer.textItem.contents = rowData[layer.name];
                            }
                        }
                    }
                }
                var doc = app.activeDocument;
                var data = {DATA_PLACEHOLDER};
                updateLayersRecursive(doc.layers, data);
                """
                jsx_script = jsx_script.replace("{DATA_PLACEHOLDER}", json.dumps(layer_data_str, ensure_ascii=False))
                ps_app.DoJavaScript(jsx_script)
                png_options = win32com.client.Dispatch("Photoshop.PNGSaveOptions")
                doc.SaveAs(save_path, png_options, True)
            doc.Close(2) 
            self.update_status("Quy trình xuất ảnh hoàn tất 100%!", "#22c55e")
            messagebox.showinfo("Success", f"Đã kết xuất hoàn tất {total_rows} file ảnh!\nThư mục: {output_dir}")
        except Exception as e:
            messagebox.showerror("System Error", f"Gặp sự cố khi render: {e}")
            self.update_status("Lỗi trong quá trình kết xuất", "#ef4444")
        finally:
            self.update_status("Đang đóng và giải phóng Photoshop ngầm...", "#eab308")
            if ps_app:
                try: ps_app.Quit()
                except: pass
            os.system("taskkill /f /im Photoshop.exe >nul 2>&1")
            self.btn_run.configure(state="normal", fg_color="#22c55e", text_color="#000000")
            self.update_status("Đã giải phóng RAM. Hệ thống sẵn sàng!", "#22c55e")
            pythoncom.CoUninitialize() 

    def init_tab_generate(self):
        frame_config = ctk.CTkFrame(self.tab_generate, fg_color="#141414", corner_radius=10)
        frame_config.pack(fill="x", padx=15, pady=10)
        ctk.CTkLabel(frame_config, text="⚙️ CẤU HÌNH CONFIG CHUỖI MA2 (MẶC ĐỊNH LƯU)", font=("Consolas", 12, "bold"), text_color="#22c55e").grid(row=0, column=0, columnspan=2, sticky="w", padx=15, pady=8)
        ctk.CTkLabel(frame_config, text="Số trước GBR (Ký tự số):", font=("Consolas", 12)).grid(row=1, column=0, sticky="w", padx=20, pady=5)
        self.ent_prefix = ctk.CTkEntry(frame_config, textvariable=self.prefix_num_var, width=150, font=("Consolas", 12))
        self.ent_prefix.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        ctk.CTkLabel(frame_config, text="Đuôi cố định mã MA2 (Suffix):", font=("Consolas", 12)).grid(row=2, column=0, sticky="w", padx=20, pady=5)
        self.ent_suffix = ctk.CTkEntry(frame_config, textvariable=self.suffix_var, width=450, font=("Consolas", 12))
        self.ent_suffix.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        frame_ctrl = ctk.CTkFrame(self.tab_generate, fg_color="#141414", corner_radius=10)
        frame_ctrl.pack(fill="x", padx=15, pady=5)
        ctk.CTkLabel(frame_ctrl, text="Giới tính:", font=("Consolas", 12)).pack(side="left", padx=15, pady=15)
        self.cbo_gender = ctk.CTkOptionMenu(frame_ctrl, values=["NAM", "NỮ"], width=100, font=("Consolas", 12))
        self.cbo_gender.set("NAM")
        self.cbo_gender.pack(side="left", padx=5)
        ctk.CTkLabel(frame_ctrl, text="Số dòng:", font=("Consolas", 12)).pack(side="left", padx=15)
        self.ent_count = ctk.CTkEntry(frame_ctrl, width=80, font=("Consolas", 12))
        self.ent_count.insert(0, "100")
        self.ent_count.pack(side="left", padx=5)
        btn_gen = ctk.CTkButton(frame_ctrl, text="⚡ GENERATE DATA", font=("Consolas", 12, "bold"), command=self.start_generate_thread)
        btn_gen.pack(side="right", padx=15)
        self.lbl_gen_status = ctk.CTkLabel(self.tab_generate, text="Trạng thái: Đang chờ...", font=("Consolas", 12, "italic"), text_color="#eab308")
        self.lbl_gen_status.pack(anchor="w", padx=20, pady=2)
        frame_out = ctk.CTkFrame(self.tab_generate, fg_color="#141414", corner_radius=10)
        frame_out.pack(fill="both", expand=True, padx=15, pady=5)
        self.txt_gen_out = ctk.CTkTextbox(frame_out, font=("Consolas", 12), fg_color="#090909", text_color="#22c55e", border_width=1, border_color="#222222")
        self.txt_gen_out.pack(fill="both", expand=True, padx=15, pady=15)
        btn_push = ctk.CTkButton(self.tab_generate, text="📥 CHUYỂN DỮ LIỆU SANG TAB PHOTOSHOP", font=("Consolas", 13, "bold"), height=45, fg_color="#00ff41", hover_color="#00aa2c", text_color="#000000", command=lambda: self.push_data_to_process(self.txt_gen_out))
        btn_push.pack(fill="x", padx=15, pady=10)

    def start_generate_thread(self):
        t = threading.Thread(target=self.generate_random_data)
        t.start()

    def generate_random_data(self):
        gender_map = {"NAM": "male", "NỮ": "female"}
        gender = gender_map[self.cbo_gender.get()]
        count = self.ent_count.get().strip()
        mid_num = self.prefix_num_var.get().strip()
        fixed_suffix = self.suffix_var.get().strip()
        if not count.isdigit():
            messagebox.showerror("Error", "Số lượng dòng phải là số nguyên!")
            return
        self.lbl_gen_status.configure(text="Connecting to Identity API Pool...", text_color="#06b6d4")
        try:
            url = f"https://randomuser.me/api/?results={count}&nat=gb&gender={gender}"
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as response:
                data = json.loads(response.read().decode('utf-8'))
            results = data.get("results", [])
            months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
            rows = ["hovaten,ho,ho2,ten,ten2,so,so2,ntns,ntns2,ma1,ma1.2,ma2,ma2.2"]
            used_pnums = set()
            for user in results:
                ho = str(user["name"]["last"]).upper()
                ten = str(user["name"]["first"]).upper()
                while True:
                    pNum = str(random.randint(100000000, 999999999))
                    if pNum not in used_pnums:
                        used_pnums.add(pNum)
                        break
                year_birth = random.randint(1946, 1999)
                m_idx = random.randint(0, 11)
                d_val = random.randint(1, 28)
                dob = f"{str(year_birth)[-2:]}{str(m_idx + 1).zfill(2)}{str(d_val).zfill(2)}"
                m2 = f"{pNum}{mid_num}GBR{dob}{fixed_suffix}"
                ho_part = ho.replace(" ", "<")
                ten_part = ten.replace(" ", "<")
                m1 = f"P<GBR{ho_part}<<{ten_part}".ljust(44, '<')[:44]
                ntns_format = f"{str(d_val).zfill(2)} {months[m_idx]} /{months[m_idx]} {str(year_birth)[-2:]}"
                rows.append(f'"{ho} {ten}","{ho}","{ho}","{ten}","{ten}","{pNum}","{pNum}","{ntns_format}","{ntns_format}","{ma1}","{ma1}","{m2}","{m2}"')
            self.txt_gen_out.delete("1.0", "end")
            self.txt_gen_out.insert("1.0", "\n".join(rows))
            self.lbl_gen_status.configure(text="Đã khởi tạo xong tổ hợp dữ liệu tùy chỉnh!", text_color="#22c55e")
        except:
            self.lbl_gen_status.configure(text="API Error!", text_color="#ef4444")

    def init_tab_gen_perfect(self):
        frame_ctrl = ctk.CTkFrame(self.tab_gen_perfect, fg_color="#141414", corner_radius=10)
        frame_ctrl.pack(fill="x", padx=15, pady=10)
        ctk.CTkLabel(frame_ctrl, text="Giới tính:", font=("Consolas", 12)).pack(side="left", padx=15, pady=15)
        self.cbo_sex_p = ctk.CTkOptionMenu(frame_ctrl, values=["NAM (M)", "NỮ (F)"], width=120, font=("Consolas", 12))
        self.cbo_sex_p.set("NAM (M)")
        self.cbo_sex_p.pack(side="left", padx=5)
        ctk.CTkLabel(frame_ctrl, text="Số dòng:", font=("Consolas", 12)).pack(side="left", padx=15)
        self.ent_count_p = ctk.CTkEntry(frame_ctrl, width=80, font=("Consolas", 12))
        self.ent_count_p.insert(0, "10")
        self.ent_count_p.pack(side="left", padx=5)
        btn_gen_p = ctk.CTkButton(frame_ctrl, text="🎲 GENERATE REAL 100% (MRZ)", font=("Consolas", 12, "bold"), command=self.start_gen_perfect_thread)
        btn_gen_p.pack(side="right", padx=15)
        self.lbl_status_p = ctk.CTkLabel(self.tab_gen_perfect, text="Trạng thái: Đang chờ...", font=("Consolas", 12, "italic"), text_color="#eab308")
        self.lbl_status_p.pack(anchor="w", padx=20, pady=2)
        frame_out = ctk.CTkFrame(self.tab_gen_perfect, fg_color="#141414", corner_radius=10)
        frame_out.pack(fill="both", expand=True, padx=15, pady=5)
        self.txt_out_p = ctk.CTkTextbox(frame_out, font=("Consolas", 12), fg_color="#090909", text_color="#22c55e", border_width=1, border_color="#222222")
        self.txt_out_p.pack(fill="both", expand=True, padx=15, pady=15)
        btn_push_p = ctk.CTkButton(self.tab_gen_perfect, text="📥 CHUYỂN DỮ LIỆU SANG TAB PHOTOSHOP", font=("Consolas", 13, "bold"), height=45, fg_color="#00ff41", hover_color="#00aa2c", text_color="#000000", command=lambda: self.push_data_to_process(self.txt_out_p))
        btn_push_p.pack(fill="x", padx=15, pady=10)

    def start_gen_perfect_thread(self):
        t = threading.Thread(target=self.generate_perfect_mrz_data)
        t.start()

    def mrz_val(self, c):
        if c == '<': return 0
        if '0' <= c <= '9': return ord(c) - 48
        return ord(c) - 55

    def mrz_cd(self, s):
        w = [7, 3, 1]
        total = 0
        for i, char in enumerate(s):
            total += self.mrz_val(char) * w[i % 3]
        return total % 10

    def generate_perfect_mrz_data(self):
        sex_sel = "M" if "NAM" in self.cbo_sex_p.get() else "F"
        count_str = self.ent_count_p.get().strip()
        if not count_str.isdigit():
            messagebox.showerror("Error", "Số lượng dòng phải là số nguyên!")
            return
        count = int(count_str)
        self.lbl_status_p.configure(text="Connecting to Identity Core API Pool...", text_color="#06b6d4")
        try:
            gender_api = "male" if sex_sel == "M" else "female"
            url = f"https://randomuser.me/api/?results={count}&nat=gb&gender={gender_api}"
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as response:
                data = json.loads(response.read().decode('utf-8'))
            results = data.get("results", [])
            name_pool = [{"ho": u["name"]["last"].upper(), "ten": u["name"]["first"].upper()} for u in results]
            months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
            rows = ["hovaten,ho,ho2,ten,ten2,so,so2,ntns,ntns2,ma1,ma1.2,ma2,ma2.2"]
            used_pnums = set()
            for _ in range(count):
                n1 = random.choice(name_pool)
                n2 = random.choice(name_pool)
                n3 = random.choice(name_pool)
                ho_parts = [n1["ho"]]
                if random.random() < 0.3: ho_parts.append(n2["ho"]) 
                name_parts = [n1["ten"]]
                if random.random() < 0.4: name_parts.append(n3["ten"]) 
                ho_full = " ".join(ho_parts)
                ten_full = " ".join(name_parts)
                full_name = f"{ho_full} {ten_full}"
                while True:
                    pNum = str(random.randint(100000000, 999999999))
                    if pNum not in used_pnums:
                        used_pnums.add(pNum)
                        break
                y = random.randint(1965, 2005)
                yy = str(y)[-2:]
                m_idx = random.randint(0, 11)
                mm = str(m_idx + 1).zfill(2)
                d_val = random.randint(1, 28)
                dd = str(d_val).zfill(2)
                text_format = f"{dd} {months[m_idx]} /{months[m_idx]} {yy}"
                dob = f"{yy}{mm}{dd}"
                exp = f"{str((int(yy) + 20) % 100).zfill(2)}{mm}{dd}" 
                c1 = self.mrz_cd(pNum)
                c2 = self.mrz_cd(dob)
                c3 = self.mrz_cd(exp)
                opt = "<<<<<<<<<<<<<<" 
                c4 = self.mrz_cd(opt)
                part2 = f"{pNum}{c1}GBR{dob}{c2}{sex_sel}{exp}{c3}{opt}{c4}"
                final_cd = self.mrz_cd(part2)
                ma2 = (part2 + str(final_cd))[:44]
                ho_mrz = "<".join(ho_parts)
                ten_mrz = "<".join(name_parts)
                ma1 = f"P<GBR{ho_mrz}<<{ten_mrz}".ljust(44, '<')[:44]
                rows.append(f'"{full_name}","{ho_full}","{ho_full}","{ten_full}","{ten_full}",{pNum},{pNum},"{text_format}","{text_format}","{ma1}","{ma1}","{ma2}","{ma2}"')
            self.txt_out_p.delete("1.0", "end")
            self.txt_out_p.insert("1.0", "\n".join(rows))
            self.lbl_status_p.configure(text="[SYSTEM]: UK PASSPORT MRZ DATA GENERATED SUCCESSFUL!", text_color="#22c55e")
        except:
            self.lbl_status_p.configure(text="Compile Error!", text_color="#ef4444")

    def push_data_to_process(self, txt_source):
        gen_data = txt_source.get("1.0", "end-1c").strip()
        if not gen_data or (gen_data.startswith("hovaten") and len(gen_data.split('\n')) <= 1):
            messagebox.showwarning("Cảnh báo", "Chưa có dữ liệu ngẫu nhiên nào được tạo!")
            return
        self.txt_data.delete("1.0", "end")
        self.txt_data.insert("1.0", gen_data)
        self.txt_data.configure(text_color="#22c55e")
        self.tab_view.set("  🚀 XỬ LÝ PHOTOSHOP  ")
        self.update_status("Import thành công dữ liệu ngẫu nhiên mới mã hóa!", "#06b6d4")

if __name__ == "__main__":
    app = PassportToolGUI()
    app.mainloop()
