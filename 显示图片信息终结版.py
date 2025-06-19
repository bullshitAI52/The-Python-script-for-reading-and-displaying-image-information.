import os
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import filedialog, messagebox, Frame
from tkinter import ttk 
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils import get_column_letter
from io import BytesIO
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- Helper Functions (�ޱ䶯) ---

def _get_image_info_and_thumbnail(file_path):
    """
    ��ȡ����ͼƬ�ļ�����ϸ��Ϣ����������ͼ��
    �� ThreadPoolExecutor �ں�̨���á�
    """
    info_dict = {'file_path': file_path, 'error': None, 'thumbnail_bytes': None}
    try:
        with Image.open(file_path) as img:
            width, height = img.size
            dpi = img.info.get('dpi', (72, 72))
            dpi_x = dpi[0] if dpi[0] > 0 else 72
            dpi_y = dpi[1] if dpi[1] > 0 else 72

            width_cm = (width / dpi_x) * 2.54
            height_cm = (height / dpi_y) * 2.54
            color_mode = img.mode
            img_format = img.format
            file_size_bytes = os.path.getsize(file_path)

            info_dict.update({
                'pixel_size': (width, height),
                'physical_size': (round(width_cm, 2), round(height_cm, 2)),
                'dpi': (dpi_x, dpi_y),
                'color_mode': color_mode,
                'format': img_format,
                'file_size': file_size_bytes,
            })

            try:
                thumb_copy = img.copy()
                thumb_copy.thumbnail((128, 128))
                img_byte_arr = BytesIO()
                thumb_copy.save(img_byte_arr, format='PNG')
                img_byte_arr.seek(0)
                info_dict['thumbnail_bytes'] = img_byte_arr
            except Exception as e_thumb:
                pass 

    except Exception as e_main:
        info_dict['error'] = str(e_main)
    
    return os.path.basename(file_path), info_dict


def format_file_size(size_bytes):
    if size_bytes is None: return "N/A"
    if size_bytes >= 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.2f} MB"
    else:
        return f"{size_bytes / 1024:.2f} KB"


class ImageToolApp:
    def __init__(self, root_window):
        self.root = root_window
        self.root.title("ͼƬ��Ϣ�봦���� v1.9 (��ɫģʽת��)")
        self.root.geometry("1100x800") # �����˴��ڿ���������°�ť

        self.folder_path = None
        self.cached_image_info = {} 
        self.preview_photo = None

        self._setup_ui()
        self.set_buttons_state(tk.DISABLED)
        self.select_button.config(state=tk.NORMAL)

    def _setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10 10 10 10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        paned_window = ttk.PanedWindow(main_frame, orient=tk.VERTICAL)
        paned_window.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        list_labelframe = ttk.Labelframe(paned_window, text="ͼƬ��Ϣ�б�", padding="5 5 5 5")
        paned_window.add(list_labelframe, weight=3)

        columns = ('filename', 'pixels', 'size_cm', 'dpi', 'mode', 'filesize')
        self.tree = ttk.Treeview(list_labelframe, columns=columns, show='headings', height=15)
        # ... ʡ���ж��� ...
        self.tree.heading('filename', text='�ļ�����')
        self.tree.heading('pixels', text='���سߴ�')
        self.tree.heading('size_cm', text='����ߴ� (cm)')
        self.tree.heading('dpi', text='DPI')
        self.tree.heading('mode', text='ģʽ')
        self.tree.heading('filesize', text='��С')
        
        self.tree.column('filename', width=250, stretch=tk.YES)
        self.tree.column('pixels', width=120, anchor='center')
        self.tree.column('size_cm', width=140, anchor='center')
        self.tree.column('dpi', width=100, anchor='center')
        self.tree.column('mode', width=80, anchor='center')
        self.tree.column('filesize', width=100, anchor='center')

        list_scrollbar_y = ttk.Scrollbar(list_labelframe, orient=tk.VERTICAL, command=self.tree.yview)
        list_scrollbar_x = ttk.Scrollbar(list_labelframe, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.config(yscrollcommand=list_scrollbar_y.set, xscrollcommand=list_scrollbar_x.set)
        list_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        list_scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        preview_labelframe = ttk.Labelframe(paned_window, text="ͼƬԤ��", padding="5")
        paned_window.add(preview_labelframe, weight=2) 

        self.preview_label = ttk.Label(preview_labelframe, text="\n\n�����Ϸ��б���ѡ��һ��ͼƬ��Ԥ��\n\n", anchor=tk.CENTER)
        self.preview_label.pack(fill=tk.BOTH, expand=True)
        
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        
        button_frame = ttk.Frame(main_frame, padding="5 0 5 0")
        button_frame.pack(fill=tk.X)
        
        self.select_button = ttk.Button(button_frame, text="ѡ��ͼƬ�ļ���", command=self.select_folder_and_load_info)
        self.select_button.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        self.rename_button = ttk.Button(button_frame, text="����������ͼƬ", command=self.start_rename_task)
        self.rename_button.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        self.export_button = ttk.Button(button_frame, text="������Ϣ��Excel", command=self.start_export_task)
        self.export_button.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        # --- [����] ��ɫת����ť ---
        self.convert_cmyk_button = ttk.Button(button_frame, text="ת��ΪCMYK", command=lambda: self.start_color_convert_task('CMYK'))
        self.convert_cmyk_button.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        self.convert_rgb_button = ttk.Button(button_frame, text="ת��ΪRGB", command=lambda: self.start_color_convert_task('RGB'))
        self.convert_rgb_button.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        # --- [�޸Ľ���] ---

        status_frame = ttk.Frame(main_frame, padding="5 5 5 5")
        status_frame.pack(fill=tk.X)
        self.status_label = ttk.Label(status_frame, text="��ѡ��һ������ͼƬ���ļ��С�")
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.progress_bar = ttk.Progressbar(status_frame, orient=tk.HORIZONTAL, length=200, mode='determinate')
        self.progress_bar.pack(side=tk.RIGHT, padx=5)

    def set_buttons_state(self, state):
        self.rename_button.config(state=state)
        self.export_button.config(state=state)
        # --- [����] �����°�ť��״̬ ---
        self.convert_cmyk_button.config(state=state)
        self.convert_rgb_button.config(state=state)
        # --- [�޸Ľ���] ---

    # --- [����] ��ɫת������ ---
    def start_color_convert_task(self, target_mode):
        if not self.folder_path:
            messagebox.showwarning("��������", "����ѡ�����ͼƬ���ļ��С�")
            return
        if not self.cached_image_info:
             messagebox.showinfo("��ʾ", "��ǰ�ļ���δ����ͼƬ��Ϣ���޷�ִ��ת����")
             return

        output_folder_name = f"{target_mode}_Converted"
        output_path = os.path.join(self.folder_path, output_folder_name)
        
        msg = (f"�˲��������԰����з� {target_mode} ģʽ��ͼƬת��Ϊ {target_mode} ģʽ��\n"
               f"ת������ļ����������µ����ļ����У�\n'{output_path}'\n\n"
               "ԭʼ�ļ����ᱻ�޸ġ��Ƿ������")

        if messagebox.askyesno("ȷ��ת������", msg):
            self.update_status(f"��ʼת��Ϊ {target_mode} ģʽ...")
            self.update_progress(0)
            self.set_buttons_state(tk.DISABLED)
            self.select_button.config(state=tk.DISABLED)
            
            thread = threading.Thread(target=self._convert_images_background_task,
                                      args=(target_mode, output_path,
                                            lambda p: self.root.after(0, self.update_progress, p),
                                            self._rename_completion_callback, # ���Ը����������Ļص�
                                            lambda err: self.root.after(0, self._task_error_callback, err, f"ת��Ϊ {target_mode}")),
                                      daemon=True)
            thread.start()

    def _convert_images_background_task(self, target_mode, output_path, progress_callback, completion_callback, error_callback):
        try:
            if not os.path.exists(output_path):
                os.makedirs(output_path)
            
            # ɸѡ����Ҫת����ͼƬ
            infos_to_convert = {
                fname: finfo for fname, finfo in self.cached_image_info.items() 
                if finfo and not finfo.get('error') and finfo.get('color_mode') != target_mode
            }
            
            total_images = len(infos_to_convert)
            processed_count = 0
            conversion_log = []

            if total_images == 0:
                if completion_callback:
                    completion_callback(f"û����Ҫת��Ϊ {target_mode} ģʽ��ͼƬ��", [])
                return

            for file_name, info in infos_to_convert.items():
                original_file_path = os.path.join(self.folder_path, file_name)
                # CMYK ģʽ��ô�Ϊ JPG
                output_ext = '.jpg' if target_mode == 'CMYK' else os.path.splitext(file_name)[1]
                output_file_name = os.path.splitext(file_name)[0] + output_ext
                output_file_path = os.path.join(output_path, output_file_name)

                try:
                    with Image.open(original_file_path) as img:
                        # ���ڴ���ɫ���ͼƬ��'P'ģʽ������תΪRGB
                        if img.mode == 'P':
                            img = img.convert('RGB')
                        
                        # ת����Ŀ��ģʽ
                        converted_img = img.convert(target_mode)
                        
                        # ���棬����JPG��ʽ����Ҫ��������
                        if output_ext.lower() == '.jpg':
                            converted_img.save(output_file_path, "jpeg", quality=95)
                        else:
                            converted_img.save(output_file_path)

                        conversion_log.append(f"�ɹ�: '{file_name}' -> '{output_file_name}'")
                except Exception as e_convert:
                    conversion_log.append(f"ʧ��: ת�� '{file_name}' ʱ���� - {e_convert}")

                processed_count += 1
                if progress_callback:
                    progress_callback(processed_count / total_images * 100)
            
            # ��ȡ�����ļ������������ڱ���
            total_all_files = len(self.cached_image_info)
            skipped_count = total_all_files - len(infos_to_convert)
            summary_message = f"ת����ɡ������� {total_images} ��ͼƬ������ {skipped_count} �� (����{target_mode}ģʽ)��"
            if completion_callback:
                completion_callback(summary_message, conversion_log)

        except Exception as e:
            if error_callback:
                error_callback(f"��ɫת�������з������ش���: {e}")
    # --- [��������] ---
    
    # ... on_tree_select, set_buttons_state �ȷ������ش��޸ģ��˴�ʡ���Ա��ּ�� ...
    # ... �������з����� _load_info_background_task, _rename_images_background_task �Ⱦ��޸Ķ� ...
    def on_tree_select(self, event):
        selection = self.tree.selection()
        if not selection:
            return

        item_id = selection[0]
        filename = self.tree.item(item_id, 'values')[0]
        info = self.cached_image_info.get(filename)
        
        if info and info.get('thumbnail_bytes'):
            try:
                thumbnail_data = info['thumbnail_bytes']
                thumbnail_data.seek(0)
                pil_image = Image.open(thumbnail_data)
                
                self.preview_photo = ImageTk.PhotoImage(pil_image)
                self.preview_label.config(image=self.preview_photo, text="")
            except Exception as e:
                self.preview_photo = None
                self.preview_label.config(image=None, text=f"�޷�����Ԥ��:\n{e}")
        else:
            self.preview_photo = None
            self.preview_label.config(image=None, text="\n\n�޿���Ԥ��\n\n")

    def update_status(self, message, is_error=False, is_warning=False):
        self.status_label.config(text=message)
        if is_error:
            self.status_label.config(foreground="red")
        elif is_warning:
            self.status_label.config(foreground="orange")
        else:
            self.status_label.config(foreground="black")
        self.root.update_idletasks()

    def update_progress(self, value):
        self.progress_bar['value'] = value
        self.root.update_idletasks()

    def select_folder_and_load_info(self):
        new_folder_path = filedialog.askdirectory()
        if new_folder_path:
            self.folder_path = new_folder_path
            self.update_status(f"���ڴ��ļ��м���ͼƬ��Ϣ: {new_folder_path}")
            for i in self.tree.get_children():
                self.tree.delete(i)
            self.cached_image_info = {}
            self.update_progress(0)
            
            self.set_buttons_state(tk.DISABLED)
            self.select_button.config(state=tk.DISABLED)

            thread = threading.Thread(target=self._load_info_background_task, daemon=True)
            thread.start()
        else:
            self.update_status("�ļ���ѡ����ȡ����", is_warning=True)

    def _load_info_background_task(self):
        try:
            image_files_basenames = [
                f for f in os.listdir(self.folder_path) 
                if os.path.isfile(os.path.join(self.folder_path, f)) and
                f.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff'))
            ]
            image_files_basenames.sort()
            
            temp_info_cache = {}
            total_files = len(image_files_basenames)
            processed_files_count = 0

            if total_files == 0:
                self.root.after(0, self.update_status, f"�ļ��� '{os.path.basename(self.folder_path)}' ��û���ҵ�֧�ֵ�ͼƬ�ļ���", True, True)
                self.root.after(0, self.update_progress, 100)
                self.root.after(0, lambda: self.select_button.config(state=tk.NORMAL))
                self.root.after(0, lambda: self.set_buttons_state(tk.DISABLED))
                return

            with ThreadPoolExecutor(max_workers=min(8, os.cpu_count() + 4 if os.cpu_count() else 8)) as executor:
                future_to_filename = {
                    executor.submit(_get_image_info_and_thumbnail, os.path.join(self.folder_path, fname)): fname 
                    for fname in image_files_basenames
                }

                for future in as_completed(future_to_filename):
                    try:
                        filename_key, info_data = future.result()
                        if info_data and not info_data.get('error'):
                            temp_info_cache[filename_key] = info_data
                        else:
                            temp_info_cache[filename_key] = None
                    except Exception as exc:
                        pass 

                    processed_files_count += 1
                    progress_percentage = (processed_files_count / total_files) * 100
                    self.root.after(0, self.update_progress, progress_percentage)

            self.cached_image_info = temp_info_cache
            self.root.after(0, self.display_image_info_from_cache)
            self.root.after(0, self.update_status, f"�ɹ����� {len(self.cached_image_info)} ��ͼƬ����Ϣ��")
            self.root.after(0, lambda: self.select_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.set_buttons_state(tk.NORMAL if self.cached_image_info else tk.DISABLED))

        except Exception as e:
            self.root.after(0, self.update_status, f"����ͼƬ��Ϣʱ��������: {e}", True)
            self.root.after(0, lambda: self.select_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.set_buttons_state(tk.DISABLED))


    def display_image_info_from_cache(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

        if not self.folder_path:
            self.update_status("����ѡ��һ���ļ��С�", is_warning=True)
            return

        if not self.cached_image_info:
            self.update_status(f"�ļ��� '{os.path.basename(self.folder_path)}' ��û��ͼƬ��Ϣ����ʾ��", is_warning=True)
            return

        for file_name, info in sorted(self.cached_image_info.items()):
            if info and not info.get('error'):
                pixel_size_str = f"{info['pixel_size'][0]}x{info['pixel_size'][1]}"
                physical_size_str = f"{info['physical_size'][0]:.2f}x{info['physical_size'][1]:.2f}"
                dpi_str = f"{info['dpi'][0]}x{info['dpi'][1]}"
                file_size_str = format_file_size(info['file_size'])
                
                values = (
                    file_name,
                    pixel_size_str,
                    physical_size_str,
                    dpi_str,
                    info['color_mode'],
                    file_size_str
                )
                self.tree.insert('', tk.END, values=values)
            else:
                values = (file_name, '�޷���ȡ��Ϣ', '', '', '', '')
                self.tree.insert('', tk.END, values=values)
        
        self.update_progress(100)
    
    def start_rename_task(self):
        if not self.folder_path:
            messagebox.showwarning("��������", "����ѡ�����ͼƬ���ļ��С�")
            return
        if not self.cached_image_info:
             messagebox.showinfo("��ʾ", "��ǰ�ļ���δ����ͼƬ��Ϣ���޷�ִ����������")
             return

        if messagebox.askyesno("ȷ�ϲ���", "ȷ��Ҫ����ͼƬ������ߴ�����������ѡ���ļ����е�ͼƬ��\n�˲���ͨ�������棬�����������"):
            self.update_status("��ʼ����������ͼƬ...")
            self.update_progress(0)
            self.set_buttons_state(tk.DISABLED)
            self.select_button.config(state=tk.DISABLED)
            
            thread = threading.Thread(target=self._rename_images_background_task,
                                      args=(lambda p: self.root.after(0, self.update_progress, p),
                                            self._rename_completion_callback,
                                            lambda err: self.root.after(0, self._task_error_callback, err, "����������")),
                                      daemon=True)
            thread.start()

    def _rename_images_background_task(self, progress_callback, completion_callback, error_callback):
        try:
            current_files_in_folder = {f for f in os.listdir(self.folder_path) if os.path.isfile(os.path.join(self.folder_path, f))}
            
            infos_to_rename = {}
            for fname, finfo in self.cached_image_info.items():
                if fname in current_files_in_folder and finfo and not finfo.get('error') and finfo.get('physical_size'):
                    infos_to_rename[fname] = finfo
            
            existing_names_on_disk = current_files_in_folder.copy()
            renamed_count = 0
            total_images_to_process = len(infos_to_rename)
            rename_log = []

            if total_images_to_process == 0:
                if completion_callback:
                    completion_callback("û���ҵ�����������������ͼƬ�ļ���", rename_log)
                return

            for file_original_name, info in infos_to_rename.items():
                original_file_path = os.path.join(self.folder_path, file_original_name)
                physical_size = info['physical_size']
                
                original_name_base, original_ext = os.path.splitext(file_original_name)
                
                new_name_base = f"{original_name_base}_{round(physical_size[0])}x{round(physical_size[1])}_cm"

                new_full_name = new_name_base + original_ext.lower()
                
                temp_existing_names_for_check = existing_names_on_disk.copy()
                if file_original_name in temp_existing_names_for_check:
                    temp_existing_names_for_check.remove(file_original_name)

                count = 0
                current_new_name_attempt = new_full_name
                while current_new_name_attempt in temp_existing_names_for_check:
                    count += 1
                    current_new_name_attempt = f"{new_name_base}_{count}{original_ext.lower()}"
                new_full_name = current_new_name_attempt
                
                if new_full_name != file_original_name:
                    new_file_path = os.path.join(self.folder_path, new_full_name)
                    try:
                        os.rename(original_file_path, new_file_path)
                        rename_log.append(f"�ɹ�: '{file_original_name}' -> '{new_full_name}'")
                        existing_names_on_disk.remove(file_original_name)
                        existing_names_on_disk.add(new_full_name)
                    except Exception as e_rename:
                        rename_log.append(f"ʧ��: ������ '{file_original_name}' ʱ���� - {e_rename}")
                else:
                    rename_log.append(f"����: '{file_original_name}' ������������")
            
                renamed_count += 1
                if progress_callback:
                    progress_callback(renamed_count / total_images_to_process * 100)
        
            if completion_callback:
                completion_callback("����������������ɡ�", rename_log)
        except Exception as e:
            if error_callback:
                error_callback(f"���������������з�������: {e}")

    def _rename_completion_callback(self, message, rename_log=None):
        self.update_status(message)
        log_summary_for_mb = ""
        if rename_log:
            log_summary_lines = rename_log[:10]
            log_summary_for_mb = "\n".join(log_summary_lines)
            if len(rename_log) > 10:
                log_summary_for_mb += f"\n...�ȹ� {len(rename_log)} ����¼��"
            
            full_message = f"{message}\n\n������־:\n{log_summary_for_mb}"
            if messagebox.askyesno("������־", f"{full_message}\n\n�Ƿ�Ҫ�������Ĳ�����־���浽�ļ���"):
                log_save_path = filedialog.asksaveasfilename(
                    defaultextension=".txt",
                    filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                    title="���������־",
                    initialdir=self.folder_path if self.folder_path else os.path.expanduser("~")
                )
                if log_save_path:
                    try:
                        with open(log_save_path, 'w', encoding='utf-8') as f:
                            f.write(f"�������: {message}\n\n��ϸ��־:\n")
                            for entry in rename_log:
                                f.write(entry + "\n")
                        messagebox.showinfo("��־�ѱ���", f"�����Ĳ�����־�ѱ��浽: {log_save_path}")
                    except Exception as e_save:
                        messagebox.showerror("����ʧ��", f"�޷�������־�ļ�: {e_save}")
            else:
                 messagebox.showinfo("�������", f"{message}\n\n������־:\n{log_summary_for_mb}")
        else:
            messagebox.showinfo("�������", message)

        # ת����ɺ��б���Ϣ�����ѹ�ʱ����ˢ�»���գ���������ֻ���°�ť״̬
        # �û������Ҫ�������ļ������ݣ���Ҫ�ֶ���ѡ��һ��
        self.select_button.config(state=tk.NORMAL)
        self.set_buttons_state(tk.NORMAL if self.cached_image_info else tk.DISABLED)


    def start_export_task(self):
        # ... ʡ�� ...

    def _save_to_excel_background_task(self, save_path, progress_callback, completion_callback, error_callback):
        # ... ʡ�� ...

    def _export_completion_callback(self, message, is_warning=False):
        self.update_status(message, is_warning=is_warning)
        self.select_button.config(state=tk.NORMAL)
        self.set_buttons_state(tk.NORMAL if self.cached_image_info else tk.DISABLED)
        if is_warning: messagebox.showwarning("������ʾ", message)
        else: messagebox.showinfo("�������", message)

    def _task_error_callback(self, error_message, task_name):
        self.update_status(f"{task_name}������ʧ��: {error_message}", is_error=True)
        self.select_button.config(state=tk.NORMAL)
        self.set_buttons_state(tk.NORMAL if self.cached_image_info else tk.DISABLED)
        self.update_progress(0)
        messagebox.showerror(f"{task_name}����", error_message)


if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style(root)
    available_themes = style.theme_names()
    if 'vista' in available_themes: style.theme_use('vista')
    elif 'clam' in available_themes: style.theme_use('clam')
    elif 'alt' in available_themes: style.theme_use('alt')
    app = ImageToolApp(root)
    root.mainloop()