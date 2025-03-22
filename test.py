import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk, ImageFilter, ImageEnhance
import cv2
import pytesseract
import shutil
import numpy as np

# Настройки по умолчанию
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
SCAN_AREA = (5, 1050, 589, 1224)  # (x1, y1, x2, y2)
OUTPUT_FOLDER = "output"
PREVIEW_SIZE = (625, 725)
DEBUG_MODE = True

# Параметры предобработки изображения
THRESHOLD_MIN = 38
THRESHOLD_MAX = 100
MEDIAN_FILTER_SIZE = 7
CONTRAST_FACTOR = 3.0
MORPH_KERNEL_SIZE = 3
MORPH_ITERATIONS = 1

def deskew(image):
    """Исправление наклона изображения (deskew)"""
    coords = np.column_stack(np.where(image > 0))
    angle = cv2.minAreaRect(coords)[-1]
    
    if angle < -45:
        angle = -(90 + angle)
    else:
        angle = -angle
    
    (h, w) = image.shape[:2]
    center = (w // 2, h // 2)
    M = cv2.getRotationMatrix2D(center, angle, 1.0)
    rotated = cv2.warpAffine(image, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    
    return rotated

def auto_threshold_and_recognize(img_array):
    """Пробуем разные значения THRESHOLD_MIN, пока не получим нормальный результат"""
    for threshold in range(21, 87, 1):  # Перебираем шагами по 1
        
        # Применяем двойной порог
        _, binary_img = cv2.threshold(img_array, threshold, 255, cv2.THRESH_BINARY)

        # Удаляем линии таблицы с помощью морфологических операций
        kernel = np.ones((MORPH_KERNEL_SIZE, MORPH_KERNEL_SIZE), np.uint8)

        # Операция открытия (эрозия+дилатация) для удаления тонких линий
        opened = cv2.morphologyEx(binary_img, cv2.MORPH_OPEN, kernel, iterations=MORPH_ITERATIONS)

        # Медианный фильтр для сглаживания
        filtered = cv2.medianBlur(opened, MEDIAN_FILTER_SIZE)

        # Увеличиваем контраст
        enhanced = cv2.convertScaleAbs(filtered, alpha=CONTRAST_FACTOR, beta=0)
        
        # Конвертируем обратно в PIL Image
        processed_img = Image.fromarray(enhanced)
        
        # Распознаем текст
        text = pytesseract.image_to_string(
            processed_img,
            config='--psm 11 --oem 3 -c tessedit_char_whitelist=0123456789'# -c tessedit_min_confidence=70'
        )
        
        batch_number = validate_batch_number(text)
        if batch_number:
            return batch_number, processed_img, threshold

    return None, img_array, threshold  # Если не нашли, вернем изображение для ручного ввода

def extract_batch_number_with_image(image_path):
    try:
        img = Image.open(image_path)
        cropped = img.crop(SCAN_AREA)
        
        # Конвертируем PIL Image в массив numpy для обработки в OpenCV
        img_array = np.array(cropped.convert('L'))

        # Конвертируем обратно в PIL Image
        _img = Image.fromarray(img_array)

        # Распознаем текст
        text = pytesseract.image_to_string(
            _img,
            config='--psm 7 --oem 3 -c tessedit_char_whitelist=0123456789'
        )
        
        batch_num = validate_batch_number(text)
        
        if DEBUG_MODE:
            debug_folder = os.path.join(OUTPUT_FOLDER, "debug")
            if not os.path.exists(debug_folder):
                os.makedirs(debug_folder)
            debug_filename = os.path.join(debug_folder, f"debug_{os.path.basename(image_path)}")
            _img.save(debug_filename)

        if DEBUG_MODE:
            print(f"Распознанный текст до   {os.path.basename(image_path)}: '{batch_num}'")

        if batch_num is None:
            
            # Применяем двойной порог
            _, binary_img = cv2.threshold(img_array, THRESHOLD_MIN, 255, cv2.THRESH_BINARY)
            
            # Увеличиваем контраст
            enhanced = cv2.convertScaleAbs(binary_img, alpha=CONTRAST_FACTOR, beta=0)

            ## Операция открытия (эрозия+дилатация) для удаления тонких линий
            #opened = cv2.morphologyEx(enhanced, cv2.MORPH_OPEN, enhanced, iterations=MORPH_ITERATIONS)

            # Медианный фильтр для сглаживания
            filtered = cv2.medianBlur(enhanced, MEDIAN_FILTER_SIZE)
            
            # Конвертируем обратно в PIL Image
            processed_img = Image.fromarray(filtered)
            
            # Увеличиваем для лучшей видимости в диалоге
            cropped_enlarged = processed_img.copy()
            new_width = processed_img.width * 2
            new_height = processed_img.height * 2
            cropped_enlarged = cropped_enlarged.resize((new_width, new_height), Image.LANCZOS)
            
            # Для отладки: сохраняем обработанную область
            if DEBUG_MODE:
                debug_folder = os.path.join(OUTPUT_FOLDER, "debug")
                if not os.path.exists(debug_folder):
                    os.makedirs(debug_folder)
                debug_filename = os.path.join(debug_folder, f"debug_{os.path.basename(image_path)}")
                processed_img.save(debug_filename)

            # Распознаем текст
            text = pytesseract.image_to_string(
                processed_img,
                config='--psm 7 --oem 3 -c tessedit_char_whitelist=0123456789'
            )
            
            batch_num = validate_batch_number(text)
            
            if DEBUG_MODE:
                print(f"Распознанный текст для  {os.path.basename(image_path)}: '{batch_num}'")

            if batch_num is not None:
                return batch_num, cropped_enlarged

            processed_img_auto = img_array
            batch_num, processed_img_auto, threshold = auto_threshold_and_recognize(processed_img_auto)

            if DEBUG_MODE:
                print(f"Распознанный текст auto {os.path.basename(image_path)}: '{batch_num}' '{threshold}'")
            return batch_num, cropped_enlarged

        return batch_num, _img
            
    except Exception as e:

        print(f"Ошибка обработки {image_path}: {str(e)}")
        return None, None

def validate_batch_number(text):
    pattern = r"\b\d{6}\b"
    cleaned_text = text.upper().replace(" ", "").replace("\n", "").strip()
    match = re.search(pattern, cleaned_text)
    return match.group(0) if match else None

class BatchRenamerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Распознавание номера партии")
        self.root.geometry("700x1000")

        self.files = []
        self.current_image = None
        self.results = {}

        if not os.path.exists(OUTPUT_FOLDER):
            os.makedirs(OUTPUT_FOLDER)

        self.create_gui()

        self.drag_data = {"x": 0, "y": 0, "rect": None}
        self.canvas.bind("<ButtonPress-1>", self.start_rect)
        self.canvas.bind("<B1-Motion>", self.drag_rect)
        self.canvas.bind("<ButtonRelease-1>", self.end_rect)

    def start_rect(self, event):
        if not self.current_image:
            return
            
        self.canvas.delete("temp_rect")
        
        # Сохраняем координаты на холсте
        self.drag_data["canvas_x_start"] = event.x
        self.drag_data["canvas_y_start"] = event.y
        
        self.drag_data["rect"] = self.canvas.create_rectangle(
            event.x, event.y, event.x, event.y,
            outline='yellow', width=2, tags="temp_rect"
        )

    def drag_rect(self, event):
        if self.drag_data["rect"]:
            self.canvas.coords(
                self.drag_data["rect"],
                self.drag_data["canvas_x_start"], 
                self.drag_data["canvas_y_start"], 
                event.x, 
                event.y
            )

    def end_rect(self, event):
        if not self.drag_data["rect"] or not self.current_image:
            return
        
        # Получаем размеры изображения и превью
        img_width, img_height = self.current_image.size
        preview = self.current_image.copy()
        preview.thumbnail(PREVIEW_SIZE)
        preview_width, preview_height = preview.size
        
        # Вычисляем смещение для центрирования
        offset_x = (self.canvas.winfo_width() - preview_width) / 2
        offset_y = (self.canvas.winfo_height() - preview_height) / 2
        
        # Вычисляем масштаб
        scale_x = img_width / preview_width
        scale_y = img_height / preview_height
        
        # Конвертируем координаты холста в координаты изображения
        x1 = int((self.drag_data["canvas_x_start"] - offset_x) * scale_x)
        y1 = int((self.drag_data["canvas_y_start"] - offset_y) * scale_y)
        x2 = int((event.x - offset_x) * scale_x)
        y2 = int((event.y - offset_y) * scale_y)
        
        # Проверяем границы
        x1 = max(0, min(x1, img_width))
        y1 = max(0, min(y1, img_height))
        x2 = max(0, min(x2, img_width))
        y2 = max(0, min(y2, img_height))
        
        # Сортируем координаты
        x1, x2 = min(x1, x2), max(x1, x2)
        y1, y2 = min(y1, y2), max(y1, y2)
        
        # Устанавливаем новую зону сканирования
        self.update_scan_area((x1, y1, x2, y2))
        
        # Обновляем предпросмотр
        selection = self.listbox.curselection()
        if selection:
            idx = selection[0]
            self.show_preview(self.files[idx])
        elif self.files:
            self.show_preview(self.files[0])
        
        self.canvas.delete("temp_rect")
        self.drag_data["rect"] = None

    def create_gui(self):
        # Верхняя панель с кнопками
        top_frame = tk.Frame(self.root)
        top_frame.pack(pady=10, fill=tk.X)
        
        self.btn_select = tk.Button(top_frame, text="Выбрать файлы", command=self.select_files)
        self.btn_select.pack(side=tk.LEFT, padx=5)
        
        self.btn_process = tk.Button(top_frame, text="Начать обработку", command=self.process_files, state=tk.DISABLED)
        self.btn_process.pack(side=tk.LEFT, padx=5)
        
        self.btn_settings = tk.Button(top_frame, text="Настройки зоны", command=self.adjust_scan_area)
        self.btn_settings.pack(side=tk.LEFT, padx=5)
        
        self.btn_processing = tk.Button(top_frame, text="Настройки обработки", command=self.adjust_processing)
        self.btn_processing.pack(side=tk.LEFT, padx=5)

        # Статусная строка
        self.status_var = tk.StringVar()
        self.status_var.set("Готов к работе. Выберите файлы.")
        self.status_label = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

        # Миниатюра с рамкой
        self.preview_frame = tk.Frame(self.root)
        self.preview_frame.pack(pady=10)
        
        self.canvas = tk.Canvas(self.preview_frame, width=PREVIEW_SIZE[0], height=PREVIEW_SIZE[1], bg='lightgrey')
        self.canvas.pack()
        
        # Список файлов
        list_frame = tk.Frame(self.root)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.listbox = tk.Listbox(list_frame, width=80)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.listbox.yview)
        
        self.listbox.bind('<<ListboxSelect>>', self.on_select_file)

    def select_files(self):
        self.files = list(filedialog.askopenfilenames(
            filetypes=[("Изображения", "*.png *.jpg *.jpeg *.bmp *.tiff")]
        ))
        
        if not self.files:
            return
            
        self.listbox.delete(0, tk.END)
        for file in self.files:
            self.listbox.insert(tk.END, os.path.basename(file))
        
        self.btn_process.config(state=tk.NORMAL)
        self.status_var.set(f"Выбрано {len(self.files)} файлов")
        
        if self.files:
            self.show_preview(self.files[0])

    def on_select_file(self, event):
        selection = self.listbox.curselection()
        if selection:
            idx = selection[0]
            if idx < len(self.files):
                self.show_preview(self.files[idx])
    
    def adjust_scan_area(self):
        global SCAN_AREA
        dialog = tk.Toplevel(self.root)
        dialog.title("Настройка зоны сканирования")
        dialog.geometry("300x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text="X1:").grid(row=0, column=0, padx=5, pady=5)
        x1_var = tk.StringVar(value=str(SCAN_AREA[0]))
        x1_entry = tk.Entry(dialog, textvariable=x1_var)
        x1_entry.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(dialog, text="Y1:").grid(row=1, column=0, padx=5, pady=5)
        y1_var = tk.StringVar(value=str(SCAN_AREA[1]))
        y1_entry = tk.Entry(dialog, textvariable=y1_var)
        y1_entry.grid(row=1, column=1, padx=5, pady=5)
        
        tk.Label(dialog, text="X2:").grid(row=2, column=0, padx=5, pady=5)
        x2_var = tk.StringVar(value=str(SCAN_AREA[2]))
        x2_entry = tk.Entry(dialog, textvariable=x2_var)
        x2_entry.grid(row=2, column=1, padx=5, pady=5)
        
        tk.Label(dialog, text="Y2:").grid(row=3, column=0, padx=5, pady=5)
        y2_var = tk.StringVar(value=str(SCAN_AREA[3]))
        y2_entry = tk.Entry(dialog, textvariable=y2_var)
        y2_entry.grid(row=3, column=1, padx=5, pady=5)
        
        def save_settings():
            try:
                x1 = int(x1_var.get())
                y1 = int(y1_var.get())
                x2 = int(x2_var.get())
                y2 = int(y2_var.get())
                
                if x1 >= x2 or y1 >= y2:
                    messagebox.showerror("Ошибка", "Координаты некорректны: X1 < X2, Y1 < Y2")
                    return
                
                self.update_scan_area((x1, y1, x2, y2))
                dialog.destroy()
                
                if self.files:
                    selection = self.listbox.curselection()
                    if selection:
                        idx = selection[0]
                        self.show_preview(self.files[idx])
                    else:
                        self.show_preview(self.files[0])
                
                self.status_var.set(f"Зона сканирования обновлена: {SCAN_AREA}")
                
            except ValueError as e:
                messagebox.showerror("Ошибка", f"Значения должны быть целыми числами: {str(e)}")
        
        save_button = tk.Button(dialog, text="Сохранить", command=save_settings)
        save_button.grid(row=4, column=0, columnspan=2, pady=10)
        
        def on_enter(event):
            save_settings()
        
        x1_entry.bind('<Return>', on_enter)
        y1_entry.bind('<Return>', on_enter)
        x2_entry.bind('<Return>', on_enter)
        y2_entry.bind('<Return>', on_enter)
        
        x1_entry.focus_set()
    
    def adjust_processing(self):
        global THRESHOLD_MIN, THRESHOLD_MAX, MEDIAN_FILTER_SIZE, CONTRAST_FACTOR, MORPH_KERNEL_SIZE, MORPH_ITERATIONS
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Настройки обработки изображения")
        dialog.geometry("450x350")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Создаем вкладки
        notebook = ttk.Notebook(dialog)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Вкладка для основных настроек
        basic_frame = ttk.Frame(notebook)
        notebook.add(basic_frame, text="Основные")
        
        # Вкладка для морфологических операций
        morph_frame = ttk.Frame(notebook)
        notebook.add(morph_frame, text="Морфология")
        
        # Основные настройки
        tk.Label(basic_frame, text="Нижний порог (черный):").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        thresh_min_var = tk.IntVar(value=THRESHOLD_MIN)
        thresh_min_scale = tk.Scale(basic_frame, from_=0, to=100, orient=tk.HORIZONTAL, variable=thresh_min_var)
        thresh_min_scale.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W+tk.E)
        
        tk.Label(basic_frame, text="Верхний порог (белый):").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        thresh_max_var = tk.IntVar(value=THRESHOLD_MAX)
        thresh_max_scale = tk.Scale(basic_frame, from_=100, to=255, orient=tk.HORIZONTAL, variable=thresh_max_var)
        thresh_max_scale.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W+tk.E)
        
        tk.Label(basic_frame, text="Размер медианного фильтра:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        median_var = tk.IntVar(value=MEDIAN_FILTER_SIZE)
        median_scale = tk.Scale(basic_frame, from_=1, to=15, orient=tk.HORIZONTAL, variable=median_var)
        median_scale.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W+tk.E)
        
        tk.Label(basic_frame, text="Фактор контраста:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        contrast_var = tk.DoubleVar(value=CONTRAST_FACTOR)
        contrast_scale = tk.Scale(basic_frame, from_=0.5, to=5.0, resolution=0.1, orient=tk.HORIZONTAL, variable=contrast_var)
        contrast_scale.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W+tk.E)
        
        # Морфологические операции
        tk.Label(morph_frame, text="Размер ядра морфологии:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        kernel_var = tk.IntVar(value=MORPH_KERNEL_SIZE)
        kernel_scale = tk.Scale(morph_frame, from_=1, to=9, orient=tk.HORIZONTAL, variable=kernel_var)
        kernel_scale.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W+tk.E)
        
        tk.Label(morph_frame, text="Количество итераций:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        iterations_var = tk.IntVar(value=MORPH_ITERATIONS)
        iterations_scale = tk.Scale(morph_frame, from_=1, to=5, orient=tk.HORIZONTAL, variable=iterations_var)
        iterations_scale.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W+tk.E)
        
        # Кнопка для тестирования настроек
        def test_settings():
            nonlocal thresh_min_var, thresh_max_var, median_var, contrast_var, kernel_var, iterations_var
            
            # Временно применяем настройки
            global THRESHOLD_MIN, THRESHOLD_MAX, MEDIAN_FILTER_SIZE, CONTRAST_FACTOR, MORPH_KERNEL_SIZE, MORPH_ITERATIONS
            old_settings = (THRESHOLD_MIN, THRESHOLD_MAX, MEDIAN_FILTER_SIZE, CONTRAST_FACTOR, MORPH_KERNEL_SIZE, MORPH_ITERATIONS)
            
            THRESHOLD_MIN = thresh_min_var.get()
            THRESHOLD_MAX = thresh_max_var.get()
            MEDIAN_FILTER_SIZE = median_var.get()
            if MEDIAN_FILTER_SIZE % 2 == 0:  # Медианный фильтр должен иметь нечетный размер
                MEDIAN_FILTER_SIZE += 1
                median_var.set(MEDIAN_FILTER_SIZE)
            CONTRAST_FACTOR = contrast_var.get()
            MORPH_KERNEL_SIZE = kernel_var.get()
            MORPH_ITERATIONS = iterations_var.get()
            
            # Проверка на текущем файле
            selection = self.listbox.curselection()
            if selection and self.files:
                idx = selection[0]
                file_path = self.files[idx]
                
                # Создаем временное окно для предпросмотра
                preview_dialog = tk.Toplevel(dialog)
                preview_dialog.title("Предпросмотр обработки")

                # Вычисляем размер окна на основе SCAN_AREA
                scan_width = SCAN_AREA[2] - SCAN_AREA[0]
                scan_height = SCAN_AREA[3] - SCAN_AREA[1]

                # Ограничиваем размер окна, чтобы оно не было слишком большим или маленьким
                preview_width = min(max(scan_width * 2, 300), 800)
                preview_height = min(max((scan_height) * 2, 250), 600)

                preview_dialog.geometry(f"{preview_width}x{preview_height+50}")
                
                # Получаем результат обработки
                batch_num, processed_img = extract_batch_number_with_image(file_path)
                
                if processed_img:
                    # Создаем и показываем обработанное изображение
                    photo = ImageTk.PhotoImage(processed_img)
                    img_label = tk.Label(preview_dialog, image=photo)
                    img_label.image = photo
                    img_label.pack(pady=10)
                    
                    result_text = f"Распознано: {batch_num}" if batch_num else "Не распознано"
                    tk.Label(preview_dialog, text=result_text, font=("Arial", 12, "bold")).pack(pady=5)
                else:
                    tk.Label(preview_dialog, text="Ошибка обработки изображения", fg="red").pack(pady=20)
                
                # Кнопка закрытия предпросмотра
                tk.Button(preview_dialog, text="Закрыть", command=preview_dialog.destroy).pack(pady=10)
            else:
                messagebox.showinfo("Информация", "Выберите файл для тестирования")
            
            # Восстанавливаем настройки
            THRESHOLD_MIN, THRESHOLD_MAX, MEDIAN_FILTER_SIZE, CONTRAST_FACTOR, MORPH_KERNEL_SIZE, MORPH_ITERATIONS = old_settings
        
        test_button = tk.Button(dialog, text="Тестировать", command=test_settings)
        test_button.pack(pady=10)
        
        # Кнопка сохранения настроек
        def save_settings():
            nonlocal thresh_min_var, thresh_max_var, median_var, contrast_var, kernel_var, iterations_var
            
            global THRESHOLD_MIN, THRESHOLD_MAX, MEDIAN_FILTER_SIZE, CONTRAST_FACTOR, MORPH_KERNEL_SIZE, MORPH_ITERATIONS
            
            THRESHOLD_MIN = thresh_min_var.get()
            THRESHOLD_MAX = thresh_max_var.get()
            MEDIAN_FILTER_SIZE = median_var.get()
            if MEDIAN_FILTER_SIZE % 2 == 0:  # Медианный фильтр должен иметь нечетный размер
                MEDIAN_FILTER_SIZE += 1
                median_var.set(MEDIAN_FILTER_SIZE)
            CONTRAST_FACTOR = contrast_var.get()
            MORPH_KERNEL_SIZE = kernel_var.get()
            MORPH_ITERATIONS = iterations_var.get()
            
            dialog.destroy()
            self.status_var.set("Настройки обработки обновлены")
        
        save_button = tk.Button(dialog, text="Сохранить", command=save_settings)
        save_button.pack(pady=10)

    def show_preview(self, image_path):
        try:
            img = Image.open(image_path)
            self.current_image = img
            
            # Создаем копию для превью
            preview = img.copy()
            preview.thumbnail(PREVIEW_SIZE)
            preview_width, preview_height = preview.size
            
            # Рассчитываем масштаб
            width, height = img.size
            scale_x = preview_width / width
            scale_y = preview_height / height
            
            # Показываем изображение
            tk_img = ImageTk.PhotoImage(preview)
            self.canvas.delete("all")
            self.canvas.image = tk_img
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()
            
            # Центрируем изображение
            offset_x = (canvas_width - preview_width) / 2
            offset_y = (canvas_height - preview_height) / 2
            self.canvas.create_image(offset_x, offset_y, anchor=tk.NW, image=tk_img)
            
            # Рисуем рамку зоны распознавания
            scaled_x1 = SCAN_AREA[0] * scale_x
            scaled_y1 = SCAN_AREA[1] * scale_y
            scaled_x2 = SCAN_AREA[2] * scale_x
            scaled_y2 = SCAN_AREA[3] * scale_y
            
            self.canvas.create_rectangle(
                offset_x + scaled_x1,
                offset_y + scaled_y1,
                offset_x + scaled_x2,
                offset_y + scaled_y2,
                outline='red',
                width=2
            )
            
            # Удаляем временные элементы
            self.canvas.delete("temp_rect")
            self.drag_data["rect"] = None

            # Если есть результат распознавания, показываем его
            filename = os.path.basename(image_path)
            if filename in self.results:
                batch_num = self.results[filename]
                self.canvas.create_text(
                    canvas_width // 2, 20,
                    text=f"Распознано: {batch_num}",
                    fill="green",
                    font=("Arial", 12, "bold")
                )
        except Exception as e:
            print(f"Ошибка отображения превью: {str(e)}")
            self.canvas.delete("all")
            self.canvas.create_text(
                PREVIEW_SIZE[0] // 2, PREVIEW_SIZE[1] // 2,
                text="Ошибка загрузки изображения",
                fill="red"
            )

    def process_files(self):
        if not self.files:
            messagebox.showinfo("Информация", "Нет выбранных файлов")
            return
            
        success = 0
        errors = []
        self.results = {}
        
        self.btn_select.config(state=tk.DISABLED)
        self.btn_process.config(state=tk.DISABLED)
        
        for idx, file_path in enumerate(self.files):
            filename = os.path.basename(file_path)
            self.status_var.set(f"Обработка {idx+1}/{len(self.files)}: {filename}")
            
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(idx)
            self.listbox.see(idx)
            self.show_preview(file_path)
            self.root.update()
            
            batch_num, processed_img = extract_batch_number_with_image(file_path)
            
            if not batch_num:
                manual_num = self.show_manual_input_dialog(filename, processed_img)
                if manual_num and re.match(r"^\d{6}$", manual_num):
                    batch_num = manual_num
                elif manual_num:
                    messagebox.showwarning("Предупреждение", "Введен некорректный номер партии. Требуется 6 цифр.")
            
            if batch_num:
                self.results[filename] = batch_num
                try:
                    ext = os.path.splitext(file_path)[1]
                    new_name = f"{batch_num}{ext}"
                    new_path = os.path.join(OUTPUT_FOLDER, new_name)
                    
                    counter = 1
                    while os.path.exists(new_path):
                        new_name = f"{batch_num}_{counter}{ext}"
                        new_path = os.path.join(OUTPUT_FOLDER, new_name)
                        counter += 1
                    
                    shutil.copy2(file_path, new_path)
                    success += 1
                    
                    self.show_preview(file_path)
                    
                except Exception as e:
                    errors.append(f"{filename}: {str(e)}")
            else:
                errors.append(f"{filename}: номер партии не распознан")
        
        self.btn_select.config(state=tk.NORMAL)
        self.btn_process.config(state=tk.NORMAL)
        
        self.show_report(success, errors)
        self.status_var.set(f"Обработка завершена. Успешно: {success}, с ошибками: {len(errors)}")
        self.btn_select.config(state=tk.NORMAL)
        self.btn_process.config(state=tk.NORMAL)
        
        #self.show_report(success, errors)
        #self.status_var.set(f"Обработка завершена. Успешно: {success}, с ошибками: {len(errors)}")

    def show_manual_input_dialog(self, filename, processed_img):
        dialog = tk.Toplevel(self.root)
        dialog.title("Ручной ввод номера партии")
        dialog.geometry("400x500")
        dialog.transient(self.root)
        dialog.grab_set()
        
        result = tk.StringVar()
        
        tk.Label(dialog, text=f"Номер партии не распознан для файла:", 
                font=("Arial", 10, "bold")).pack(pady=(10, 0))
        tk.Label(dialog, text=filename).pack(pady=(0, 10))
        
        if processed_img:
            img_display = processed_img.copy()
            img_display.thumbnail((350, 200))
            
            photo = ImageTk.PhotoImage(img_display)
            img_label = tk.Label(dialog, image=photo)
            img_label.image = photo
            img_label.pack(pady=10)
            
            tk.Label(dialog, text="Обработанное изображение области номера партии", 
                    font=("Arial", 9, "italic")).pack()
        
        tk.Label(dialog, text="Введите номер партии (6 цифр):", 
                font=("Arial", 10)).pack(pady=(20, 5))
        entry = tk.Entry(dialog, textvariable=result, font=("Arial", 12), width=15)
        entry.pack(pady=5)
        entry.focus_set()
        
        def on_ok():
            dialog.destroy()
        
        def on_cancel():
            result.set("")
            dialog.destroy()
        
        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=20)
        
        tk.Button(button_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="Пропустить", command=on_cancel, width=10).pack(side=tk.LEFT, padx=10)
        
        entry.bind("<Return>", lambda event: on_ok())
        
        dialog.wait_window()
        
        return result.get()

    def show_report(self, success, errors):
        if not success and not errors:
            messagebox.showinfo("Результат", "Ничего не обработано!")
            return
            
        report = []
        if success > 0:
            report.append(f"Успешно обработано: {success} файлов")
        
        if errors:
            report.append(f"Ошибки при обработке {len(errors)} файлов:")
            for error in errors[:10]:
                report.append(f"• {error}")
            
            if len(errors) > 10:
                report.append(f"...и еще {len(errors) - 10} ошибок")
        
        messagebox.showinfo("Результат обработки", "\n".join(report))
        
        if errors and messagebox.askyesno("Сохранить лог?", "Хотите сохранить полный лог ошибок в файл?"):
            import datetime
            
            log_path = os.path.join(OUTPUT_FOLDER, "error_log.txt")
            with open(log_path, "w", encoding="utf-8") as log_file:
                log_file.write(f"Дата обработки: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                log_file.write(f"Всего файлов: {len(self.files)}\n")
                log_file.write(f"Успешно обработано: {success}\n")
                log_file.write(f"Ошибок: {len(errors)}\n\n")
                log_file.write("Детали ошибок:\n")
                for error in errors:
                    log_file.write(f"- {error}\n")
            messagebox.showinfo("Лог сохранен", f"Лог ошибок сохранен в:\n{log_path}")

    def update_scan_area(self, new_area):
        global SCAN_AREA
        SCAN_AREA = new_area
        if DEBUG_MODE:
            print(f"Обновленная SCAN_AREA: {SCAN_AREA}")

def save_settings_to_file():
    settings = {
        'SCAN_AREA': SCAN_AREA,
        'THRESHOLD_MIN': THRESHOLD_MIN,
        'THRESHOLD_MAX': THRESHOLD_MAX,
        'MEDIAN_FILTER_SIZE': MEDIAN_FILTER_SIZE,
        'CONTRAST_FACTOR': CONTRAST_FACTOR,
        'MORPH_KERNEL_SIZE': MORPH_KERNEL_SIZE,
        'MORPH_ITERATIONS': MORPH_ITERATIONS
    }
    
    try:
        import json
        with open('settings.json', 'w') as f:
            json.dump(settings, f)
        print("Настройки сохранены в settings.json")
    except Exception as e:
        print(f"Ошибка сохранения настроек: {str(e)}")

def load_settings_from_file():
    try:
        import json
        with open('settings.json', 'r') as f:
            settings = json.load(f)
        
        global SCAN_AREA, THRESHOLD_MIN, THRESHOLD_MAX, MEDIAN_FILTER_SIZE, CONTRAST_FACTOR, MORPH_KERNEL_SIZE, MORPH_ITERATIONS
        
        SCAN_AREA = tuple(settings['SCAN_AREA'])
        THRESHOLD_MIN = settings['THRESHOLD_MIN']
        THRESHOLD_MAX = settings['THRESHOLD_MAX']
        MEDIAN_FILTER_SIZE = settings['MEDIAN_FILTER_SIZE']
        CONTRAST_FACTOR = settings['CONTRAST_FACTOR']
        MORPH_KERNEL_SIZE = settings['MORPH_KERNEL_SIZE']
        MORPH_ITERATIONS = settings['MORPH_ITERATIONS']
        
        print("Настройки загружены из settings.json")
    except FileNotFoundError:
        print("Файл настроек не найден, используются значения по умолчанию")
    except Exception as e:
        print(f"Ошибка загрузки настроек: {str(e)}")

if __name__ == "__main__":
    # Пробуем загрузить настройки
    load_settings_from_file()
    
    app = BatchRenamerApp()
    
    # Сохраняем настройки при закрытии
    app.root.protocol("WM_DELETE_WINDOW", lambda: [save_settings_to_file(), app.root.destroy()])
    
    app.root.mainloop()