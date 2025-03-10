import os
import tkinter as tk
from tkinter import filedialog, messagebox
from ctypes import windll, POINTER, c_void_p, byref, c_int, create_string_buffer, Structure, sizeof
from PIL import Image
import requests
from urllib.parse import urlparse
import win32com.client

class BITMAP(Structure):
    _fields_ = [
        ("bmType", c_int),
        ("bmWidth", c_int),
        ("bmHeight", c_int),
        ("bmWidthBytes", c_int),
        ("bmPlanes", c_int),
        ("bmBitsPixel", c_int),
        ("bmBits", c_void_p)
    ]

def get_target_path(file_path, output_dir):
    if file_path.lower().endswith('.url'):
        def get_favicon_url(url):
            parsed_url = urlparse(url)
            base_url = f"{parsed_url.scheme}://{parsed_url.netloc}"
            favicon_url = f"{base_url}/favicon.ico"
            return favicon_url

        def download_favicon(url, output_path):
            response = requests.get(url, stream=True)
            if response.status_code == 200:
                with open(output_path, 'wb') as file:
                    for chunk in response.iter_content(1024):
                        file.write(chunk)
                return output_path
            else:
                return None

        with open(file_path, 'r') as file:
            for line in file:
                if line.startswith('URL='):
                    url = line.split('=', 1)[1].strip()
                    favicon_url = get_favicon_url(url)
                    output_path = os.path.join(output_dir, 'favicon.ico')
                    downloaded_favicon = download_favicon(favicon_url, output_path)
                    if downloaded_favicon:
                        return downloaded_favicon
                    else:
                        print(f"Failed to download favicon from {favicon_url}")
                        return None
    elif file_path.lower().endswith('.lnk'):
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(file_path)
        return shortcut.TargetPath
    
    elif file_path.lower().endswith('.png') or file_path.lower().endswith('.jpg') or file_path.lower().endswith('.jpeg'):
        return file_path

    return file_path

def convert_image_to_icon(image_path, output_dir):
    if not os.path.isfile(image_path):
        print(f"File not found: {image_path}")
        return None

    try:
        img = Image.open(image_path)
        img = img.convert("RGBA")
        icon_sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
        icon_path = os.path.join(output_dir, f"{os.path.splitext(os.path.basename(image_path))[0]}.ico")
        img.save(icon_path, format='ICO', sizes=icon_sizes)
        print(f"Icon created at {icon_path}")
        messagebox.showinfo("Success", f"Icon created at {icon_path}")
        return icon_path
    except Exception as e:
        print(f"Failed to convert image to icon: {e}")
        messagebox.showerror("Error", f"Failed to convert image to icon: {e}")
        return None

def extract_icon(file_path, output_dir, icon_index=0):
    target_path = get_target_path(file_path, output_dir)
    if target_path is None:
        return

    if target_path.lower().endswith('.png') or target_path.lower().endswith('.jpg') or target_path.lower().endswith('.jpeg'):
        convert_image_to_icon(target_path, output_dir)
        return

    if target_path.startswith('http://') or target_path.startswith('https://'):
        print(f"Cannot extract icon from URL: {target_path}")
        messagebox.showerror("Error", f"Cannot extract icon from URL: {target_path}")
        return

    if not os.path.isfile(target_path):
        print(f"File not found: {target_path}")
        return

    large_icon = (c_void_p * 1)()
    small_icon = (c_void_p * 1)()
    num_icons = windll.shell32.ExtractIconExW(target_path, icon_index, large_icon, small_icon, 1)
    if num_icons == 0:
        print(f"Failed to extract icon from {file_path}")
        return

    icon_path = os.path.join(output_dir, f"{os.path.basename(target_path)}.ico")

    # Save the icon to a file
    hicon = large_icon[0]
    hdc = windll.user32.GetDC(None)
    size = 256  # Size of the icon
    bmp = windll.gdi32.CreateCompatibleBitmap(hdc, size, size)
    memdc = windll.gdi32.CreateCompatibleDC(hdc)
    oldbmp = windll.gdi32.SelectObject(memdc, bmp)
    windll.user32.DrawIconEx(memdc, 0, 0, hicon, size, size, 0, None, 3)
    windll.gdi32.SelectObject(memdc, oldbmp)
    windll.gdi32.DeleteDC(memdc)
    windll.user32.ReleaseDC(None, hdc)

    # Convert the bitmap to an image and save it
    bmpinfo = BITMAP()
    windll.gdi32.GetObjectW(bmp, sizeof(BITMAP), byref(bmpinfo))
    bmpstr = create_string_buffer(bmpinfo.bmWidthBytes * bmpinfo.bmHeight * 4)
    windll.gdi32.GetBitmapBits(bmp, len(bmpstr), bmpstr)
    img = Image.frombuffer('RGBA', (bmpinfo.bmWidth, bmpinfo.bmHeight), bmpstr, 'raw', 'BGRA', 0, 1)
    img = img.convert("RGBA")  # Ensure the image is in RGBA format
    img.save(icon_path)

    print(f"Icon extracted to {icon_path}")
    messagebox.showinfo("Success", f"Icon extracted to {icon_path}")

def browse_file():
    file_path = filedialog.askopenfilename(title="Select File", filetypes=[("Executable Files", "*.exe *.lnk *.url *.png *.jpeg *.jpg")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)

def browse_directory():
    directory_path = filedialog.askdirectory(title="Select Output Directory")
    directory_entry.delete(0, tk.END)
    directory_entry.insert(0, directory_path)

def extract_icon_gui():
    file_path = file_entry.get()
    output_dir = directory_entry.get()

    if not file_path or not output_dir:
        messagebox.showerror("Error", "Please select a file and an output directory.")
        return

    extract_icon(file_path, output_dir)

app = tk.Tk()
app.title("Icon Extractor")

file_label = tk.Label(app, text="File:")
file_label.grid(row=0, column=0, padx=10, pady=5, sticky="e")

file_entry = tk.Entry(app, width=50)
file_entry.grid(row=0, column=1, padx=10, pady=5)

file_button = tk.Button(app, text="Browse", command=browse_file)
file_button.grid(row=0, column=2, padx=10, pady=5)

directory_label = tk.Label(app, text="Output Directory:")
directory_label.grid(row=1, column=0, padx=10, pady=5, sticky="e")

directory_entry = tk.Entry(app, width=50)
directory_entry.grid(row=1, column=1, padx=10, pady=5)

directory_button = tk.Button(app, text="Browse", command=browse_directory)
directory_button.grid(row=1, column=2, padx=10, pady=5)

extract_button = tk.Button(app, text="Extract Icon", command=extract_icon_gui)
extract_button.grid(row=2, column=0, columnspan=3, pady=10)

app.mainloop()