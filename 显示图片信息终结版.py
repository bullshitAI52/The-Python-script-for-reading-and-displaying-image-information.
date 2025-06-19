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

# --- Helper Functions (无变动) ---

def _get_image_info_and_thumbnail(file_path):
    """
    获取单个图片文件的详细信息并生成缩略图。
    由 ThreadPoolExecutor 在后台调用。
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
        self.root.title("图片信息与处理工具 v1.9 (颜色模式转换)")
        self.root.geometry("1100x800") # 增加了窗口宽度以容纳新按钮

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

        list_labelframe = ttk.Labelframe(paned_window, text="图片信息列表", padding="5 5 5 5")
        paned_window.add(list_labelframe, weight=3)

        columns = ('filename', 'pixels', 'size_cm', 'dpi', 'mode', 'filesize')
        self.tree = ttk.Treeview(list_labelframe, columns=columns, show='headings', height=15)
        # ... 省略列定义 ...
        self.tree.heading('filename', text='文件名称')
        self.tree.heading('pixels', text='像素尺寸')
        self.tree.heading('size_cm', text='物理尺寸 (cm)')
        self.tree.heading('dpi', text='DPI')
        self.tree.heading('mode', text='模式')
        self.tree.heading('filesize', text='大小')
        
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
        
        preview_labelframe = ttk.Labelframe(paned_window, text="图片预览", padding="5")
        paned_window.add(preview_labelframe, weight=2) 

        self.preview_label = ttk.Label(preview_labelframe, text="\n\n请在上方列表中选择一张图片以预览\n\n", anchor=tk.CENTER)
        self.preview_label.pack(fill=tk.BOTH, expand=True)
        
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        
        button_frame = ttk.Frame(main_frame, padding="5 0 5 0")
        button_frame.pack(fill=tk.X)
        
        self.select_button = ttk.Button(button_frame, text="选择图片文件夹", command=self.select_folder_and_load_info)
        self.select_button.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        self.rename_button = ttk.Button(button_frame, text="批量重命名图片", command=self.start_rename_task)
        self.rename_button.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        self.export_button = ttk.Button(button_frame, text="导出信息到Excel", command=self.start_export_task)
        self.export_button.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        # --- [新增] 颜色转换按钮 ---
        self.convert_cmyk_button = ttk.Button(button_frame, text="转换为CMYK", command=lambda: self.start_color_convert_task('CMYK'))
        self.convert_cmyk_button.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        self.convert_rgb_button = ttk.Button(button_frame, text="转换为RGB", command=lambda: self.start_color_convert_task('RGB'))
        self.convert_rgb_button.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        # --- [修改结束] ---

        status_frame = ttk.Frame(main_frame, padding="5 5 5 5")
        status_frame.pack(fill=tk.X)
        self.status_label = ttk.Label(status_frame, text="请选择一个包含图片的文件夹。")
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.progress_bar = ttk.Progressbar(status_frame, orient=tk.HORIZONTAL, length=200, mode='determinate')
        self.progress_bar.pack(side=tk.RIGHT, padx=5)

    def set_buttons_state(self, state):
        self.rename_button.config(state=state)
        self.export_button.config(state=state)
        # --- [新增] 管理新按钮的状态 ---
        self.convert_cmyk_button.config(state=state)
        self.convert_rgb_button.config(state=state)
        # --- [修改结束] ---

    # --- [新增] 颜色转换任务 ---
    def start_color_convert_task(self, target_mode):
        if not self.folder_path:
            messagebox.showwarning("操作警告", "请先选择包含图片的文件夹。")
            return
        if not self.cached_image_info:
             messagebox.showinfo("提示", "当前文件夹未加载图片信息，无法执行转换。")
             return

        output_folder_name = f"{target_mode}_Converted"
        output_path = os.path.join(self.folder_path, output_folder_name)
        
        msg = (f"此操作将尝试把所有非 {target_mode} 模式的图片转换为 {target_mode} 模式。\n"
               f"转换后的文件将保存在新的子文件夹中：\n'{output_path}'\n\n"
               "原始文件不会被修改。是否继续？")

        if messagebox.askyesno("确认转换操作", msg):
            self.update_status(f"开始转换为 {target_mode} 模式...")
            self.update_progress(0)
            self.set_buttons_state(tk.DISABLED)
            self.select_button.config(state=tk.DISABLED)
            
            thread = threading.Thread(target=self._convert_images_background_task,
                                      args=(target_mode, output_path,
                                            lambda p: self.root.after(0, self.update_progress, p),
                                            self._rename_completion_callback, # 可以复用重命名的回调
                                            lambda err: self.root.after(0, self._task_error_callback, err, f"转换为 {target_mode}")),
                                      daemon=True)
            thread.start()

    def _convert_images_background_task(self, target_mode, output_path, progress_callback, completion_callback, error_callback):
        try:
            if not os.path.exists(output_path):
                os.makedirs(output_path)
            
            # 筛选出需要转换的图片
            infos_to_convert = {
                fname: finfo for fname, finfo in self.cached_image_info.items() 
                if finfo and not finfo.get('error') and finfo.get('color_mode') != target_mode
            }
            
            total_images = len(infos_to_convert)
            processed_count = 0
            conversion_log = []

            if total_images == 0:
                if completion_callback:
                    completion_callback(f"没有需要转换为 {target_mode} 模式的图片。", [])
                return

            for file_name, info in infos_to_convert.items():
                original_file_path = os.path.join(self.folder_path, file_name)
                # CMYK 模式最好存为 JPG
                output_ext = '.jpg' if target_mode == 'CMYK' else os.path.splitext(file_name)[1]
                output_file_name = os.path.splitext(file_name)[0] + output_ext
                output_file_path = os.path.join(output_path, output_file_name)

                try:
                    with Image.open(original_file_path) as img:
                        # 对于带调色板的图片（'P'模式），先转为RGB
                        if img.mode == 'P':
                            img = img.convert('RGB')
                        
                        # 转换到目标模式
                        converted_img = img.convert(target_mode)
                        
                        # 保存，对于JPG格式，需要设置质量
                        if output_ext.lower() == '.jpg':
                            converted_img.save(output_file_path, "jpeg", quality=95)
                        else:
                            converted_img.save(output_file_path)

                        conversion_log.append(f"成功: '{file_name}' -> '{output_file_name}'")
                except Exception as e_convert:
                    conversion_log.append(f"失败: 转换 '{file_name}' 时出错 - {e_convert}")

                processed_count += 1
                if progress_callback:
                    progress_callback(processed_count / total_images * 100)
            
            # 获取所有文件的总数，用于报告
            total_all_files = len(self.cached_image_info)
            skipped_count = total_all_files - len(infos_to_convert)
            summary_message = f"转换完成。共处理 {total_images} 张图片，跳过 {skipped_count} 张 (已是{target_mode}模式)。"
            if completion_callback:
                completion_callback(summary_message, conversion_log)

        except Exception as e:
            if error_callback:
                error_callback(f"颜色转换过程中发生严重错误: {e}")
    # --- [新增结束] ---
    
    # ... on_tree_select, set_buttons_state 等方法无重大修改，此处省略以保持简洁 ...
    # ... 其余所有方法如 _load_info_background_task, _rename_images_background_task 等均无改动 ...
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
                self.preview_label.config(image=None, text=f"无法加载预览:\n{e}")
        else:
            self.preview_photo = None
            self.preview_label.config(image=None, text="\n\n无可用预览\n\n")

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
            self.update_status(f"正在从文件夹加载图片信息: {new_folder_path}")
            for i in self.tree.get_children():
                self.tree.delete(i)
            self.cached_image_info = {}
            self.update_progress(0)
            
            self.set_buttons_state(tk.DISABLED)
            self.select_button.config(state=tk.DISABLED)

            thread = threading.Thread(target=self._load_info_background_task, daemon=True)
            thread.start()
        else:
            self.update_status("文件夹选择已取消。", is_warning=True)

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
                self.root.after(0, self.update_status, f"文件夹 '{os.path.basename(self.folder_path)}' 中没有找到支持的图片文件。", True, True)
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
            self.root.after(0, self.update_status, f"成功加载 {len(self.cached_image_info)} 张图片的信息。")
            self.root.after(0, lambda: self.select_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.set_buttons_state(tk.NORMAL if self.cached_image_info else tk.DISABLED))

        except Exception as e:
            self.root.after(0, self.update_status, f"加载图片信息时发生错误: {e}", True)
            self.root.after(0, lambda: self.select_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.set_buttons_state(tk.DISABLED))


    def display_image_info_from_cache(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

        if not self.folder_path:
            self.update_status("请先选择一个文件夹。", is_warning=True)
            return

        if not self.cached_image_info:
            self.update_status(f"文件夹 '{os.path.basename(self.folder_path)}' 中没有图片信息可显示。", is_warning=True)
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
                values = (file_name, '无法获取信息', '', '', '', '')
                self.tree.insert('', tk.END, values=values)
        
        self.update_progress(100)
    
    def start_rename_task(self):
        if not self.folder_path:
            messagebox.showwarning("操作警告", "请先选择包含图片的文件夹。")
            return
        if not self.cached_image_info:
             messagebox.showinfo("提示", "当前文件夹未加载图片信息，无法执行重命名。")
             return

        if messagebox.askyesno("确认操作", "确定要根据图片的物理尺寸批量重命名选定文件夹中的图片吗？\n此操作通常不可逆，请谨慎操作！"):
            self.update_status("开始批量重命名图片...")
            self.update_progress(0)
            self.set_buttons_state(tk.DISABLED)
            self.select_button.config(state=tk.DISABLED)
            
            thread = threading.Thread(target=self._rename_images_background_task,
                                      args=(lambda p: self.root.after(0, self.update_progress, p),
                                            self._rename_completion_callback,
                                            lambda err: self.root.after(0, self._task_error_callback, err, "批量重命名")),
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
                    completion_callback("没有找到符合重命名条件的图片文件。", rename_log)
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
                        rename_log.append(f"成功: '{file_original_name}' -> '{new_full_name}'")
                        existing_names_on_disk.remove(file_original_name)
                        existing_names_on_disk.add(new_full_name)
                    except Exception as e_rename:
                        rename_log.append(f"失败: 重命名 '{file_original_name}' 时出错 - {e_rename}")
                else:
                    rename_log.append(f"跳过: '{file_original_name}' 无需重命名。")
            
                renamed_count += 1
                if progress_callback:
                    progress_callback(renamed_count / total_images_to_process * 100)
        
            if completion_callback:
                completion_callback("批量重命名操作完成。", rename_log)
        except Exception as e:
            if error_callback:
                error_callback(f"批量重命名过程中发生错误: {e}")

    def _rename_completion_callback(self, message, rename_log=None):
        self.update_status(message)
        log_summary_for_mb = ""
        if rename_log:
            log_summary_lines = rename_log[:10]
            log_summary_for_mb = "\n".join(log_summary_lines)
            if len(rename_log) > 10:
                log_summary_for_mb += f"\n...等共 {len(rename_log)} 条记录。"
            
            full_message = f"{message}\n\n部分日志:\n{log_summary_for_mb}"
            if messagebox.askyesno("保存日志", f"{full_message}\n\n是否要将完整的操作日志保存到文件？"):
                log_save_path = filedialog.asksaveasfilename(
                    defaultextension=".txt",
                    filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                    title="保存操作日志",
                    initialdir=self.folder_path if self.folder_path else os.path.expanduser("~")
                )
                if log_save_path:
                    try:
                        with open(log_save_path, 'w', encoding='utf-8') as f:
                            f.write(f"操作结果: {message}\n\n详细日志:\n")
                            for entry in rename_log:
                                f.write(entry + "\n")
                        messagebox.showinfo("日志已保存", f"完整的操作日志已保存到: {log_save_path}")
                    except Exception as e_save:
                        messagebox.showerror("保存失败", f"无法保存日志文件: {e_save}")
            else:
                 messagebox.showinfo("操作完成", f"{message}\n\n部分日志:\n{log_summary_for_mb}")
        else:
            messagebox.showinfo("操作完成", message)

        # 转换完成后，列表信息可能已过时，但刷新会清空，所以这里只更新按钮状态
        # 用户如果需要看到新文件夹内容，需要手动再选择一次
        self.select_button.config(state=tk.NORMAL)
        self.set_buttons_state(tk.NORMAL if self.cached_image_info else tk.DISABLED)


    def start_export_task(self):
        # ... 省略 ...

    def _save_to_excel_background_task(self, save_path, progress_callback, completion_callback, error_callback):
        # ... 省略 ...

    def _export_completion_callback(self, message, is_warning=False):
        self.update_status(message, is_warning=is_warning)
        self.select_button.config(state=tk.NORMAL)
        self.set_buttons_state(tk.NORMAL if self.cached_image_info else tk.DISABLED)
        if is_warning: messagebox.showwarning("导出提示", message)
        else: messagebox.showinfo("导出完成", message)

    def _task_error_callback(self, error_message, task_name):
        self.update_status(f"{task_name}过程中失败: {error_message}", is_error=True)
        self.select_button.config(state=tk.NORMAL)
        self.set_buttons_state(tk.NORMAL if self.cached_image_info else tk.DISABLED)
        self.update_progress(0)
        messagebox.showerror(f"{task_name}错误", error_message)


if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style(root)
    available_themes = style.theme_names()
    if 'vista' in available_themes: style.theme_use('vista')
    elif 'clam' in available_themes: style.theme_use('clam')
    elif 'alt' in available_themes: style.theme_use('alt')
    app = ImageToolApp(root)
    root.mainloop()