import os
from PIL import Image
import tkinter as tk
from tkinter import filedialog, Listbox, Button, messagebox

# ��ȡͼƬ�����سߴ硢����ߴ磨���ף���DPI����ɫģʽ����ʽ���ļ���С
def get_image_info(file_path):
    try:
        with Image.open(file_path) as img:
            width, height = img.size  # ���سߴ�
            dpi = img.info.get('dpi', (72, 72))  # Ĭ�� DPI Ϊ 72x72
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
        print(f"�޷����� {file_path}: {e}")
        return None

# ��������ߴ磺�ӽ�����ʱ��������
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

# �������ļ����������ͻ
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

# ������ͼƬ
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
            print(f"�ѽ� {file} ������Ϊ {new_name}")

# ͼ�ν�����
class ImageToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ͼƬ��Ϣ������������")
        self.listbox = Listbox(root, width=130, height=25)
        self.listbox.pack(pady=10)
        self.select_button = Button(root, text="ѡ���ļ���", command=self.select_folder)
        self.select_button.pack(pady=5)
        self.show_info_button = Button(root, text="��ʾͼƬ��Ϣ", command=self.show_image_info)
        self.show_info_button.pack(pady=5)
        self.rename_button = Button(root, text="������ͼƬ", command=self.rename_images)
        self.rename_button.pack(pady=5)

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.folder_path = folder_path
            self.listbox.delete(0, tk.END)
            self.listbox.insert(tk.END, f"��ѡ���ļ���: {folder_path}")
            self.listbox.insert(tk.END, "���� '��ʾͼƬ��Ϣ' �� '������ͼƬ'")

    def format_file_size(self, size_bytes):
        if size_bytes >= 1024 * 1024:
            return f"{size_bytes / (1024 * 1024):.2f} MB"
        else:
            return f"{size_bytes / 1024:.2f} KB"

    def show_image_info(self):
        if not hasattr(self, 'folder_path'):
            messagebox.showwarning("����", "����ѡ���ļ���")
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
                    f"{file}: {psize[0]}x{psize[1]} ����, "
                    f"{cm_size[0]:.2f}x{cm_size[1]:.2f} ����, "
                    f"DPI: {dpi[0]}x{dpi[1]}, ģʽ: {mode}, ��ʽ: {fmt}, ��С: {size_str}"
                )
            else:
                self.listbox.insert(tk.END, f"{file}: �޷���ȡ��Ϣ")

    def rename_images(self):
        if not hasattr(self, 'folder_path'):
            messagebox.showwarning("����", "����ѡ���ļ���")
            return
        confirm = messagebox.askyesno("ȷ��", "ȷ��Ҫ������ͼƬ��")
        if confirm:
            rename_images(self.folder_path)
            self.show_image_info()

# ��������
if __name__ == "__main__":
    root = tk.Tk()
    app = ImageToolApp(root)
    root.mainloop()
