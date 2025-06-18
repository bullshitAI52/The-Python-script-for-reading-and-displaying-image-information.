import os
from PIL import Image
import tkinter as tk
from tkinter import filedialog, Listbox, Button, messagebox

# 获取图片的像素尺寸、物理尺寸（厘米）、DPI、颜色模式、格式、文件大小
def get_image_info(file_path):
    try:
        with Image.open(file_path) as img:
            width, height = img.size  # 像素尺寸
            dpi = img.info.get('dpi', (72, 72))  # 默认 DPI 为 72x72
            width_cm = (width / dpi[0]) * 2.54
            height_cm = (height / dpi[1]) * 2.54
            color_mode = img.mode
            img_format = img.format
            file_size_bytes = os.path.getsize(file_path)
            return {
                'pixel_size': (width, height),
                'physical_size': (width_cm, height_cm),
                'dpi': dpi,
                'color_mode': color_mode,
                'format': img_format,
                'file_size': file_size_bytes
            }
    except Exception as e:
        print(f"无法处理 {file_path}: {e}")
        return None

# 修正物理尺寸：接近整数时四舍五入
def round_physical_size(size):
    rounded = []
    for value in size:
        fractional = value - int(value)
        if 0.01 <= fractional <= 0.02 or 0.98 <= fractional <= 0.99:
            value = round(value)
        else:
            value = round(value)
        rounded.append(value)
    return tuple(rounded)

# 生成新文件名并处理冲突
def generate_new_filename(folder, physical_size, ext, existing_names):
    rounded_size = round_physical_size(physical_size)
    size_str = f"{rounded_size[0]}x{rounded_size[1]}_cm"
    new_name = size_str
    count = 0
    while True:
        full_name = new_name + ext
        if full_name not in existing_names:
            existing_names.add(full_name)
            return full_name
        count += 1
        new_name = f"{size_str}_{count}"

# 重命名图片
def rename_images(folder_path):
    image_files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp'))]
    existing_names = set()
    for file in image_files:
        file_path = os.path.join(folder_path, file)
        info = get_image_info(file_path)
        if info:
            physical_size = info['physical_size']
            ext = os.path.splitext(file)[1]
            new_name = generate_new_filename(folder_path, physical_size, ext, existing_names)
            new_file_path = os.path.join(folder_path, new_name)
            os.rename(file_path, new_file_path)
            print(f"已将 {file} 重命名为 {new_name}")

# 图形界面类
class ImageToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("图片信息与重命名工具")
        self.listbox = Listbox(root, width=130, height=25)
        self.listbox.pack(pady=10)
        self.select_button = Button(root, text="选择文件夹", command=self.select_folder)
        self.select_button.pack(pady=5)
        self.show_info_button = Button(root, text="显示图片信息", command=self.show_image_info)
        self.show_info_button.pack(pady=5)
        self.rename_button = Button(root, text="重命名图片", command=self.rename_images)
        self.rename_button.pack(pady=5)

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.folder_path = folder_path
            self.listbox.delete(0, tk.END)
            self.listbox.insert(tk.END, f"已选择文件夹: {folder_path}")
            self.listbox.insert(tk.END, "请点击 '显示图片信息' 或 '重命名图片'")

    def format_file_size(self, size_bytes):
        if size_bytes >= 1024 * 1024:
            return f"{size_bytes / (1024 * 1024):.2f} MB"
        else:
            return f"{size_bytes / 1024:.2f} KB"

    def show_image_info(self):
        if not hasattr(self, 'folder_path'):
            messagebox.showwarning("警告", "请先选择文件夹")
            return
        self.listbox.delete(0, tk.END)
        image_files = [f for f in os.listdir(self.folder_path) if f.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp'))]
        image_files.sort()
        for file in image_files:
            file_path = os.path.join(self.folder_path, file)
            info = get_image_info(file_path)
            if info:
                psize = info['pixel_size']
                cm_size = info['physical_size']
                dpi = info['dpi']
                mode = info['color_mode']
                fmt = info['format']
                size_str = self.format_file_size(info['file_size'])
                self.listbox.insert(
                    tk.END,
                    f"{file}: {psize[0]}x{psize[1]} 像素, "
                    f"{cm_size[0]:.2f}x{cm_size[1]:.2f} 厘米, "
                    f"DPI: {dpi[0]}x{dpi[1]}, 模式: {mode}, 格式: {fmt}, 大小: {size_str}"
                )
            else:
                self.listbox.insert(tk.END, f"{file}: 无法获取信息")

    def rename_images(self):
        if not hasattr(self, 'folder_path'):
            messagebox.showwarning("警告", "请先选择文件夹")
            return
        confirm = messagebox.askyesno("确认", "确定要重命名图片吗？")
        if confirm:
            rename_images(self.folder_path)
            self.show_image_info()

# 启动程序
if __name__ == "__main__":
    root = tk.Tk()
    app = ImageToolApp(root)
    root.mainloop()
