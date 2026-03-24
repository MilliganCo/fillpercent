import fitz  # PyMuPDF
from PIL import Image
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import os
import time
import threading

PHOTOSHOP_EXE = r"D:\Adobe\Adobe Photoshop 2025\Photoshop.exe"  # Укажите корректный путь
SCRIPT_PATH = r"D:\PyCharm 2024.2.3\PycharmProjects\fillpercent\blackfill.jsx"  # Путь к JSX-скрипту
OUTPUT_FOLDER = r"D:\PyCharm 2024.2.3\PycharmProjects\fillpercent\Prints"  # Куда сохраняются обработанные файлы
pdf_path = r"D:\PyCharm 2024.2.3\PycharmProjects\fillpercent\raw\Leopard_Color 27х18,5.pdf"
def extract_images_from_pdf(pdf_path):
    """Извлечение изображений из PDF."""
    doc = fitz.open(pdf_path)
    images = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        for img_index, img in enumerate(page.get_images(full=True)):
            xref = img[0]
            pixmap = fitz.Pixmap(doc, xref)

            try:
                if pixmap.alpha:
                    img = Image.frombytes("RGBA", (pixmap.width, pixmap.height), pixmap.samples)
                else:
                    img = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
                    img = img.convert("RGBA")
                images.append(img)
            except Exception as e:
                print(f"Ошибка при обработке изображения на странице {page_num + 1}, индекс {img_index}: {e}")
                continue
    return images

def calculate_fill_percentage(image: Image):
    """Рассчитать процент заполненности пикселями."""
    pixels = np.array(image)
    alpha_channel = pixels[:, :, 3]
    opaque_pixels = np.sum(alpha_channel > 0)
    total_pixels = alpha_channel.size
    fill_percentage = (opaque_pixels / total_pixels) * 100
    return fill_percentage

def save_to_excel(data):
    """Сохранение результатов в Excel."""
    file_path = "fill_results.xlsx"
    if not os.path.exists(file_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Results"
        ws.append(["PDF файл", "Изображение", "Заполненность (%)"])
        header_fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
        for col in range(1, 4):
            ws[f"{get_column_letter(col)}1"].fill = header_fill
            ws[f"{get_column_letter(col)}1"].font = Font(bold=True)
            ws[f"{get_column_letter(col)}1"].alignment = Alignment(horizontal="center")
        wb.save(file_path)

    wb = openpyxl.load_workbook(file_path)
    ws = wb["Results"]
    for row in data:
        ws.append(row)
    wb.save(file_path)

def process_pdf(pdf_path):
    """Обработка PDF-файла и расчет заполненности."""
    try:
        images = extract_images_from_pdf(pdf_path)
        if not images:
            messagebox.showerror("Ошибка", "В PDF не найдено изображений!")
            return

        total_fill_percentage = 0
        results = []
        excel_data = []

        for idx, img in enumerate(images):
            fill_percentage = calculate_fill_percentage(img)
            total_fill_percentage += fill_percentage
            results.append(f"Изображение {idx + 1}: Заполненность = {fill_percentage:.2f}%")
            excel_data.append([os.path.basename(pdf_path), f"Изображение {idx + 1}", fill_percentage])

        avg_fill_percentage = total_fill_percentage / len(images)
        results.append(f"Средняя заполненность: {avg_fill_percentage:.2f}%")

        save_to_excel(excel_data)
        messagebox.showinfo("Результаты", "\n".join(results))

    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка при обработке PDF: {e}")

command = f'start "" "{PHOTOSHOP_EXE}" "{pdf_path}" /r "{SCRIPT_PATH}"'
print("Запускаем команду:", command)

def process_all_pdfs():
    """Выбор PDF-файлов, обработка в Photoshop, расчет заполненности и сохранение в Excel."""
    def process_files():
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")], title="Выберите PDF‑файлы")
        if not file_paths:
            return

        processed_files = []

        for pdf_path in file_paths:
            try:
                pdf_path = os.path.abspath(pdf_path)  # Делаем путь абсолютным
                script_path = os.path.abspath(SCRIPT_PATH)

                # Формируем команду с `cmd /k start`
                command = f'start "" "{PHOTOSHOP_EXE}" "{pdf_path}" /r "{SCRIPT_PATH}"'
                print("Запускаем команду:", command)

                # Запускаем Photoshop и НЕ БЛОКИРУЕМ GUI
                result = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

                # Ждём появления обработанного файла в `Prints`
                processed_pdf = os.path.join(OUTPUT_FOLDER, os.path.basename(pdf_path))
                timeout = 30  # Максимальное время ожидания (секунды)
                start_time = time.time()
                while not os.path.exists(processed_pdf):
                    if time.time() - start_time > timeout:
                        messagebox.showerror("Ошибка", f"Файл {os.path.basename(pdf_path)} не появился в папке!")
                        break
                    time.sleep(1)

                if os.path.exists(processed_pdf):
                    processed_files.append(processed_pdf)

            except Exception as e:
                messagebox.showerror("Ошибка Photoshop", f"Ошибка при обработке {os.path.basename(pdf_path)}: {e}")

        # Теперь анализируем все обработанные файлы
        for processed_pdf in processed_files:
            process_pdf(processed_pdf)

    # Запускаем обработку файлов в отдельном потоке
    threading.Thread(target=process_files, daemon=True).start()


def main():
    """Графический интерфейс."""
    root = tk.Tk()
    root.title("PDF Analyzer")
    root.geometry("400x300")
    root.configure(bg="#2E2E2E")

    label = tk.Label(root, text="Анализ PDF заполненности", bg="#2E2E2E", fg="white", font=("Arial", 16))
    label.pack(pady=20)

    btn_process_all = tk.Button(root, text="Выбрать и обработать PDF", command=process_all_pdfs,
                                bg="#555555", fg="white", font=("Arial", 12), relief="ridge", borderwidth=2)
    btn_process_all.pack(pady=10, ipadx=10, ipady=5)

    btn_exit = tk.Button(root, text="Выход", command=root.destroy,
                         bg="#8B0000", fg="white", font=("Arial", 12), relief="ridge", borderwidth=2)
    btn_exit.pack(pady=10, ipadx=10, ipady=5)

    root.mainloop()

if __name__ == "__main__":
    main()
