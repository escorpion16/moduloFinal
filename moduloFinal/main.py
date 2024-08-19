import customtkinter as ctk
import cv2
from PIL import Image, ImageTk
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from ultralytics import YOLO
import numpy as np

# Configuración del modelo de visión
model = YOLO("best.pt")

# Directorios de almacenamiento
CAPTURAS_DIR = "C:/Users/Dennis/Documents/moduloFinal/moduloFinal/Archivos"  # Cambia esta ruta a la de tu computadora
EXCEL_DIR = "C:/Users/Dennis/Documents/moduloFinal/moduloFinal/Archivos"  # Cambia esta ruta a la de tu computadora
os.makedirs(CAPTURAS_DIR, exist_ok=True)
os.makedirs(EXCEL_DIR, exist_ok=True)

# Lista de precios de los objetos
ITEMS = [
    {"name": "Base-fusible", "price": 1.36},
    {"name": "Borneras-electricas", "price": 12.82},
    {"name": "Bornes-bananas", "price": 1.23},
    {"name": "CM1241-RS-232-Siemens", "price": 316.50},
    {"name": "CM1243-5-Siemens", "price": 1002.00},
    {"name": "CSM1277-SIMATIC-NET-Siemens", "price": 319.00},
    {"name": "Disyuntor-1P", "price": 6.15},
    {"name": "Disyuntor-2P", "price": 15.70},
    {"name": "E-D-32DI-Siemens", "price": 387.00},
    {"name": "Fuente-A.-MeanWell", "price": 122.00},
    {"name": "Fuente-A.-PM1507-Siemens", "price": 414.00},
    {"name": "HMI-KTP700-Siemens", "price": 686.00},
    {"name": "Luz-piloto-roja", "price": 4.82},
    {"name": "Luz-piloto-verde", "price": 4.82},
    {"name": "Modulo-PM1207-Siemens", "price": 90.58},
    {"name": "Parada-de-emergencia", "price": 5.00},
    {"name": "Potenciometro", "price": 63.16},
    {"name": "Pulsador-N.A.-verde", "price": 12.17},
    {"name": "Relay-Camsco", "price": 4.11},
    {"name": "Relay-Schneider", "price": 12.00},
    {"name": "Relay-Siemens", "price": 8.40},
    {"name": "Router-tplink", "price": 20.00},
    {"name": "SCALANCE-XB005-Siemens", "price": 442.00},
    {"name": "SIMATIC-S7-1200-Siemens", "price": 335.00},
    {"name": "SIMATIC-S7-1500-Siemens", "price": 665.00},
    {"name": "Selector", "price": 22.36},
    {"name": "Siemens-LOGO-Power", "price": 101.50},
    {"name": "Variador-de-frecuencia-Siemens", "price": 170.00},
    {"name": "Voltimetro-Analogico", "price": 9.10},
    {"name": "Voltimetro-digital", "price": 5.60}
]


class App(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("Visión Artificial en Modulos Didácticos")
        self.geometry("800x600")
        self.modo_claro = True  # Empezar en modo claro por defecto
        self.setup_ui()

    def setup_ui(self):
        self.switch_theme_button = ctk.CTkButton(self, text="Cambiar Modo (Claro/Oscuro)", command=self.switch_theme)
        self.switch_theme_button.pack(pady=10, side=tk.TOP, anchor="ne")

        # Cargar el logo usando PIL
        try:
            logo_image = Image.open("logo4.png")  # Asegúrate de tener un archivo logo.png
            logo_image = logo_image.resize((200, 150), Image.LANCZOS)  # Ajustar tamaño del logo si es necesario
            self.logo_image_tk = ImageTk.PhotoImage(logo_image)
            self.logo_label = ctk.CTkLabel(self, image=self.logo_image_tk)
            self.logo_label.pack(pady=10)
        except FileNotFoundError:
            messagebox.showerror("Error", "No se encontró el archivo logo4.png. Asegúrate de que el archivo esté en la ruta correcta.")

        self.instructions_button = ctk.CTkButton(self, text="Instrucciones", command=self.show_instructions)
        self.instructions_button.pack(pady=10)

        self.start_button = ctk.CTkButton(self, text="Iniciar VisionAI", command=self.open_vision_screen)
        self.start_button.pack(pady=10)

        self.close_button = ctk.CTkButton(self, text="Cerrar App", command=self.quit)
        self.close_button.pack(pady=10, side=tk.BOTTOM)

    def switch_theme(self):
        if self.modo_claro:
            ctk.set_appearance_mode("dark")
            self.modo_claro = False
        else:
            ctk.set_appearance_mode("light")
            self.modo_claro = True

    def show_instructions(self):
        instructions = "1. Inicie VisionAI para ver la cámara en tiempo real.\n\n" \
                       "2. Use el botón Capturar para guardar una imagen y generar el archivo excel.\n\n" \
                       "3. Use Ver Archivos para comprobar que han generado los archivos.\n\n" \
                       "4. Use Ver Etiquetas para visualizar todas las clases de objetos que puede detectar el modelo.\n\n" \
                       "5. Use Cerrar App para finalizar esta aplicación."
        messagebox.showinfo("Instrucciones", instructions)

    def open_vision_screen(self):
        self.vision_screen = VisionScreen(self)
        self.vision_screen.grab_set()


class VisionScreen(ctk.CTkToplevel):

    def __init__(self, master):
        super().__init__(master)
        self.title("VisionAI")
        self.geometry("800x600")
        self.master = master
        self.modo_claro = master.modo_claro

        self.setup_ui()
        self.start_camera()

    def setup_ui(self):
        self.camera_frame = ctk.CTkFrame(self)
        self.camera_frame.place(relwidth=0.6, relheight=1, relx=0, rely=0)

        self.side_frame = ctk.CTkFrame(self)
        self.side_frame.place(relwidth=0.4, relheight=1, relx=0.6, rely=0)

        self.camera_label = ctk.CTkLabel(self.camera_frame)
        self.camera_label.pack(pady=10)

        self.capture_button = ctk.CTkButton(self, text="Capturar Imagen", command=self.capture_image)
        self.capture_button.pack(pady=10, side=tk.TOP)

        self.view_captures_button = ctk.CTkButton(self, text="Ver Archivos", command=self.view_captures)
        self.view_captures_button.pack(pady=10, side=tk.TOP)

        self.view_labels_button = ctk.CTkButton(self, text="Ver Etiquetas", command=self.view_labels)
        self.view_labels_button.pack(pady=10, side=tk.TOP)

        self.close_button = ctk.CTkButton(self, text="Cerrar App", command=self.quit)
        self.close_button.pack(pady=10, side=tk.TOP)

        self.back_button = ctk.CTkButton(self, text="Regresar", command=self.destroy)
        self.back_button.pack(pady=10, side=tk.BOTTOM)

        self.label_text = tk.StringVar()
        self.label_display = ctk.CTkLabel(self.side_frame, textvariable=self.label_text)
        self.label_display.pack(pady=10)

    def start_camera(self):
        self.video_capture = cv2.VideoCapture(1)
        self.update_frame()

    def update_frame(self):
        ret, frame = self.video_capture.read()
        if ret:
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            results = model(frame)
            annotated_frame = results[0].plot()  # Draw bounding boxes on the frame
            self.show_image_in_tkinter(annotated_frame)

            # Update labels
            self.update_labels(results)

        self.after(10, self.update_frame)

    def show_image_in_tkinter(self, frame):
        image = Image.fromarray(frame)
        image_tk = ImageTk.PhotoImage(image)
        self.camera_label.configure(image=image_tk)
        self.camera_label.image = image_tk

    def update_labels(self, results):
        labels = results[0].names
        counts = {}
        for det in results[0].boxes.cls:
            label = labels[int(det.item())]
            counts[label] = counts.get(label, 0) + 1
        self.label_text.set('\n'.join([f"{k}: {v}" for k, v in counts.items()]))

    def capture_image(self):
        ret, frame = self.video_capture.read()
        if ret:
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            results = model(frame)
            annotated_frame = results[0].plot()
            img_path = os.path.join(CAPTURAS_DIR, "captura.png")
            cv2.imwrite(img_path, cv2.cvtColor(annotated_frame, cv2.COLOR_RGB2BGR))
            self.save_detection_data(results)
            messagebox.showinfo("Captura de Imagen ", f"Imagen capturada y guardada en {img_path}")

    def save_detection_data(self, results):
        detection_data = []
        labels = results[0].names
        counts = {}
        for det in results[0].boxes.cls:
            label = labels[int(det.item())]
            counts[label] = counts.get(label, 0) + 1

        for label, count in counts.items():
            item = next((item for item in ITEMS if item["name"] == label), None)
            if item:
                total_price = count * item["price"]
                detection_data.append({
                    "Elemento detectado": label,
                    "Cantidad": count,
                    "Precio Unitario": item["price"],
                    "Total": total_price
                })

        df = pd.DataFrame(detection_data)
        total_price_sum = df["Total"].sum()
        df.loc[len(df.index)] = ["Precio total", "", "", total_price_sum]

        #total_objects_sum = df["Cantidad"].sum()
        #df.loc[len(df.index)] = ["Total elementos", "", "", total_objects_sum]

        capture_name = f"prefactura_{os.path.basename('captura.png').split('.')[0]}.xlsx"
        excel_path = os.path.join(CAPTURAS_DIR, capture_name)
        df.to_excel(excel_path, index=False)
        return df

    def view_captures(self):
        captures = os.listdir(CAPTURAS_DIR)
        if not captures:
            messagebox.showinfo("Ver Archivos", "No hay capturas para mostrar.")
            return

        self.captures_screen = CapturesScreen(self)
        self.captures_screen.grab_set()

    def view_labels(self):
        labels = [item["name"] for item in ITEMS]
        messagebox.showinfo("Etiquetas del Modelo\n", "\n".join(labels))


class CapturesScreen(ctk.CTkToplevel):

    def __init__(self, master):
        super().__init__(master)
        self.title("Ver Capturas")
        self.geometry("800x600")
        self.master = master
        self.setup_ui()

    def setup_ui(self):
        self.capture_listbox = tk.Listbox(self)
        self.capture_listbox.pack(pady=10, fill=tk.BOTH, expand=True)

        self.load_captures()

        #self.export_button = ctk.CTkButton(self, text="Exportar Excel", command=self.export_to_excel)
        #self.export_button.pack(pady=10, side=tk.BOTTOM)

        self.back_button = ctk.CTkButton(self, text="Regresar", command=self.destroy)
        self.back_button.pack(pady=10, side=tk.BOTTOM)

    def load_captures(self):
        captures = os.listdir(CAPTURAS_DIR)
        for capture in captures:
            self.capture_listbox.insert(tk.END, capture)

    def export_to_excel(self):
        captures = [file for file in os.listdir(CAPTURAS_DIR) if file.endswith('.png')]
        if not captures:
            messagebox.showinfo("Exportar a Excel", "No hay capturas para exportar.")
            return

        data = []
        for capture in captures:
            img_path = os.path.join(CAPTURAS_DIR, capture)
            frame = cv2.imread(img_path)
            results = model(frame)
            labels = results[0].names
            counts = {}
            for det in results[0].boxes.cls:
                label = labels[int(det.item())]
                counts[label] = counts.get(label, 0) + 1

            for label, count in counts.items():
                item = next((item for item in ITEMS if item["name"] == label), None)
                if item:
                    total_price = count * item["price"]
                    data.append({
                        "Elemento detectado": label,
                        "Cantidad": count,
                        "Precio Unitario": item["price"],
                        "Total": total_price
                    })

        df = pd.DataFrame(data)
        total_price_sum = df["Total"].sum()
        df.loc[len(df.index)] = ["Precio total", "", "", total_price_sum]
        excel_path = os.path.join(EXCEL_DIR, "prefactura_tablero.xlsx")
        df.to_excel(excel_path, index=False)
        messagebox.showinfo("Exportar a Excel", f"Datos exportados a {excel_path}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
