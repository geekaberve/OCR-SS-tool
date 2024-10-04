import tkinter as tk
from tkinter import filedialog, ttk
from PIL import Image, ImageTk
import os
import logging
from paddleOCR import initialize_ocr_SLANet_LCNetV2, process_image, group_into_rows, save_as_xlsx, draw_bounding_boxes
import pandas as pd

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class OCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OCR to Excel Converter")
        self.root.geometry("1280x720")
        self.root.configure(bg="#4A0E4E")
        self.root.iconbitmap('icon.ico')

        self.ocr_engine = tk.StringVar()
        self.ocr_engine.set("PaddleOCR")

        self.setup_ui()

    def setup_ui(self):
        # Set up styles
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton', font=('Helvetica', 12), padding=6, background="#8E44AD")  # Light purple button
        style.configure('TLabel', font=('Helvetica', 12), foreground="#FFFFFF")  # White text
        style.configure('TCombobox', font=('Helvetica', 12))
        style.configure('TFrame', background="#4A0E4E")  # Dark purple background

        # Main frame
        self.main_frame = ttk.Frame(self.root, style='TFrame')
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Center frame for initial view
        self.center_frame = ttk.Frame(self.main_frame, style='TFrame')
        self.center_frame.pack(expand=True)

        # Logo
        logo_img = Image.open("icon.png")
        logo_img = logo_img.resize((100, 100), Image.LANCZOS)
        self.logo_photo = ImageTk.PhotoImage(logo_img)
        logo_label = ttk.Label(self.center_frame, image=self.logo_photo, style='TLabel')
        logo_label.pack(pady=(0, 20))

        # OCR Engine Selection
        ttk.Label(self.center_frame, text="Select OCR Engine:", style='TLabel').pack(pady=(0, 5))
        ocr_dropdown = ttk.Combobox(self.center_frame, textvariable=self.ocr_engine, state="readonly")
        ocr_dropdown['values'] = ('PaddleOCR', 'Tesseract', 'EasyOCR', 'Google Vision AI')
        ocr_dropdown.pack(pady=(0, 20))

        # Upload Button
        self.upload_button = ttk.Button(self.center_frame, text="Upload Image", command=self.select_image, style='TButton')
        self.upload_button.pack(pady=(0, 20))

        # Status Label
        self.status_label = ttk.Label(self.center_frame, text="", style='TLabel')
        self.status_label.pack(pady=(0, 10))

        # Prepare frames for results (hidden initially)
        self.top_frame = ttk.Frame(self.main_frame, style='TFrame')
        self.results_frame = ttk.Frame(self.main_frame, style='TFrame')

    def select_image(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Image files", ("*.png", "*.jpg", "*.jpeg", "*.bmp", "*.tiff"))]
        )
        if file_path:
            logger.info(f"Selected image: {file_path}")
            self.process_image(file_path)

    def process_image(self, file_path):
        # Start processing directly without a new thread
        try:
            ocr_engine = self.ocr_engine.get()

            if ocr_engine == "PaddleOCR":
                self.process_with_paddleocr(file_path)
            elif ocr_engine == "Tesseract":
                self.process_with_tesseract(file_path)
            elif ocr_engine == "EasyOCR":
                self.process_with_easyocr(file_path)
            elif ocr_engine == "Google Vision AI":
                self.process_with_google_vision(file_path)
            else:
                raise ValueError("Please select an OCR engine.")
        except Exception as e:
            logger.error(f"Error processing image: {str(e)}")
            self.status_label.config(text=f"Error: {str(e)}")

    def process_with_paddleocr(self, file_path):
        ocr = initialize_ocr_SLANet_LCNetV2()
        data = process_image(file_path, ocr)
        rows = group_into_rows(data)

        if not rows:
            raise ValueError("No data extracted from image.")

        output_xlsx = os.path.splitext(file_path)[0] + "_output.xlsx"
        save_as_xlsx(rows, output_xlsx)

        output_image_path = os.path.splitext(file_path)[0] + "_output_image.jpg"
        draw_bounding_boxes(file_path, data, output_image_path)

        self.status_label.config(text=f"Excel file saved: {output_xlsx}")
        self.display_results(output_image_path, output_xlsx)

    def process_with_tesseract(self, file_path):
        # Placeholder for Tesseract OCR processing
        self.status_label.config(text="Tesseract OCR not implemented yet")

    def process_with_easyocr(self, file_path):
        # Placeholder for EasyOCR processing
        self.status_label.config(text="EasyOCR not implemented yet")

    def process_with_google_vision(self, file_path):
        # Placeholder for Google Vision AI processing
        self.status_label.config(text="Google Vision AI not implemented yet")

    def display_results(self, image_path, excel_path):
        self.reorganize_layout()
        
        # Clear previous results
        for widget in self.results_frame.winfo_children():
            widget.destroy()

        # Display image with bounding boxes
        self.display_image(image_path, self.results_frame)

        # Display Excel content
        self.display_excel_content(excel_path, self.results_frame)

    def reorganize_layout(self):
        # Clear the center frame
        for widget in self.center_frame.winfo_children():
            widget.destroy()
        self.center_frame.pack_forget()

        # Set up top frame
        self.top_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(self.top_frame, text="Select OCR Engine:", style='TLabel').pack(side=tk.LEFT, padx=(0, 10))
        ocr_dropdown = ttk.Combobox(self.top_frame, textvariable=self.ocr_engine, state="readonly")
        ocr_dropdown['values'] = ('PaddleOCR', 'Tesseract', 'EasyOCR', 'Google Vision AI')
        ocr_dropdown.pack(side=tk.LEFT, padx=(0, 20))
        ttk.Button(self.top_frame, text="Upload Image", command=self.select_image, style='TButton').pack(side=tk.LEFT)

        # Set up results frame
        self.results_frame.pack(fill=tk.BOTH, expand=True)

    def display_image(self, image_path, panel):
        image = Image.open(image_path)
        panel_width = self.root.winfo_width() // 2
        panel_height = self.root.winfo_height() - 100

        # Calculate scaling factor to fit the image within the panel
        width_ratio = panel_width / image.width
        height_ratio = panel_height / image.height
        scale_factor = min(width_ratio, height_ratio)

        new_width = int(image.width * scale_factor)
        new_height = int(image.height * scale_factor)

        image = image.resize((new_width, new_height), Image.LANCZOS)
        photo = ImageTk.PhotoImage(image)
        
        image_label = ttk.Label(panel, image=photo, background="#4A0E4E")
        image_label.image = photo
        image_label.pack(side=tk.LEFT, expand=True)

    def display_excel_content(self, excel_path, panel):
        df = pd.read_excel(excel_path)
        text_widget = tk.Text(panel, wrap=tk.NONE, bg="#4A0E4E", fg="white")
        text_widget.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Display DataFrame content
        text_widget.insert(tk.END, df.to_string(index=False))

if __name__ == "__main__":
    root = tk.Tk()
    app = OCRApp(root)
    root.mainloop()