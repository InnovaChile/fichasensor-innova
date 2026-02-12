import os
import sys
import customtkinter as ctk
from PIL import Image, ImageTk
from tkinter import messagebox
from datetime import datetime
from src.excel.ExcelReader import ExcelReader
from src.generator.BitacoraGenerator import BitacoraGenerator
from src.generator.SensorSheetGenerator import SensorSheetGenerator
from src.utils.FileUtils import FileUtils
from src.utils.DateUtils import DateUtils
from src.models.ProjectInfo import ProjectInfo
from src.api.CorfoSoapClient import CorfoSoapClient

class MainApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Gestión de Fichas Sensor Innova Chile")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.geometry("620x690")

        # ===== ICONO VENTANA =====
        try:
            ico_path = FileUtils.get_ico_corfo_path()
            if os.path.exists(ico_path):
                self.iconbitmap(ico_path)
        except Exception as e:
            print("No se pudo cargar el ícono .ico:", e)

        self.codigo_var = ctk.StringVar()
        self.nombre_proyecto = ctk.StringVar()
        self.nombre_beneficiario = ctk.StringVar()
        self.nombre_responsable = ctk.StringVar()
        self.accion_var = ctk.StringVar(value="Generar bitácora de seguimiento")
        self.ficha_var = ctk.StringVar()

        self.fechas_actuales = []
        self.filas_proyecto_actual = []
        self.project_info = None

        # ---- ENCABEZADO ----
        header = ctk.CTkLabel(self, text="Gestión de Fichas Sensor Innova Chile",
                                font=("Arial", 26, "bold"), fg_color="#221E7C",
                                text_color="#FFFFFF", corner_radius=10)
        header.pack(pady=20, fill="x", padx=0)

        # ---- BUSCADOR ----
        buscador_frame = ctk.CTkFrame(self, fg_color="#3F3F3F", corner_radius=8)
        buscador_frame.pack(fill="x", padx=32, pady=(24, 8))
        ctk.CTkLabel(buscador_frame, text="Buscar proyecto por código:", font=("Arial", 14)).grid(row=0, column=0, padx=10, pady=14, sticky="w")
        ctk.CTkEntry(buscador_frame, textvariable=self.codigo_var, width=180).grid(row=0, column=1, padx=8, pady=14)
        ctk.CTkButton(buscador_frame, text="Buscar", command=self.buscar_ficha, fg_color="#221E7C").grid(row=0, column=2, padx=16, pady=14)

        # ---- SEPARADOR ----
        ctk.CTkLabel(self, text=" ").pack(pady=1)  # Espacio visual

        # ---- DATOS DE PROYECTO ----
        datos_frame = ctk.CTkFrame(self, fg_color="#2a2a2a", corner_radius=8)
        datos_frame.pack(fill="x", padx=32, pady=4)
        ctk.CTkLabel(datos_frame, text="Nombre del Proyecto:", font=("Arial", 13)).grid(row=0, column=0, sticky="e", padx=8, pady=6)
        ctk.CTkEntry(datos_frame, textvariable=self.nombre_proyecto, state="readonly", width=340).grid(row=0, column=1, padx=8, pady=6)
        ctk.CTkLabel(datos_frame, text="Beneficiario:", font=("Arial", 13)).grid(row=1, column=0, sticky="e", padx=8, pady=6)
        ctk.CTkEntry(datos_frame, textvariable=self.nombre_beneficiario, state="readonly", width=340).grid(row=1, column=1, padx=8, pady=6)
        ctk.CTkLabel(datos_frame, text="Responsable:", font=("Arial", 13)).grid(row=2, column=0, sticky="e", padx=8, pady=6)
        ctk.CTkEntry(datos_frame, textvariable=self.nombre_responsable, state="readonly", width=340).grid(row=2, column=1, padx=8, pady=6)

        # ---- ESTADO MENSAJE (carga) ----
        self.estado_msg = ctk.CTkLabel(self, text="", font=("Arial", 13, "bold"), text_color="#19BC82")
        self.estado_msg.pack(pady=(4, 0))

        # ---- SEPARADOR ----
        sep = ctk.CTkFrame(self, height=2, fg_color="#3F3F3F")
        sep.pack(fill="x", padx=32, pady=12)

        # ---- ACCIONES ----
        acciones_frame = ctk.CTkFrame(self, fg_color="#3F3F3F", corner_radius=8)
        acciones_frame.pack(fill="x", padx=32, pady=6)
        ctk.CTkLabel(acciones_frame, text="Acción:", font=("Arial", 14)).grid(row=0, column=0, padx=8, pady=12, sticky="e")
        self.combo_accion = ctk.CTkComboBox(acciones_frame, values=["Generar bitácora de seguimiento", "Generar registro de visitas"],
                                            variable=self.accion_var, width=260, command=self.on_accion_cambio)
        self.combo_accion.grid(row=0, column=1, padx=12, pady=12)
        ctk.CTkLabel(acciones_frame, text="Ficha sensor:", font=("Arial", 14)).grid(row=1, column=0, sticky="e", padx=8, pady=8)
        self.combo_fichas = ctk.CTkComboBox(acciones_frame, values=["Selecciona una ficha"], variable=self.ficha_var, width=260, command=self.on_ficha_cambio)
        self.combo_fichas.grid(row=1, column=1, padx=12, pady=8)
        self.combo_fichas.configure(state="disabled")

        # ---- BOTÓN FINAL ----
        self.btn_generar = ctk.CTkButton(self, text="GENERAR DOCUMENTO", font=("Arial", 16, "bold"),
                                            fg_color="#221E7C", corner_radius=10, height=40, command=self.generar_documento, state="disabled")
        self.btn_generar.pack(pady=28, padx=32, fill="x")
        
        # ---- LOGO DE CORFO (GRANDE) ----
        logo_path = FileUtils.get_logo_corfo_path()
        logo_img = ctk.CTkImage(light_image=Image.open(logo_path), size=(120, 60))
        ctk.CTkLabel(self, image=logo_img, text="").pack(pady=(14, 4))

        # ---- FOOTER ----
        ctk.CTkLabel(self, text="Innova Chile - Corfo", font=("Arial", 13, "italic"), text_color="#72C7D5").pack(side="bottom", pady=10)

    def get_soap_data(self, codigo):
        client = CorfoSoapClient()
        response = client.get_project_data(codigo)  # Debe retornar un dict o similar
        if response:
            return {
                "Nombre Proyecto": response.get("Nombre Proyecto", ""),
                "Nombre Beneficiario": response.get("Nombre Beneficiario", ""),
                "Ejecutivo Técnico": response.get("Ejecutivo Técnico", ""),
                "Representante Legal": response.get("Representante Legal", "")
            }
        return {}

    def fecha_key(self, meeting):
        valor = meeting.get("Fecha de reunión", "")
        try:
            if isinstance(valor, datetime):
                return valor
            for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
                try:
                    return datetime.strptime(str(valor), fmt)
                except Exception:
                    continue
            return datetime.strptime(DateUtils.format_date(valor), "%d-%m-%Y")
        except Exception:
            return datetime(1900, 1, 1)

    def buscar_ficha(self):
        self.estado_msg.configure(text="Cargando... Por favor espere.", text_color="#F2C744")
        self.update()
        codigo = self.codigo_var.get().strip()
        if not codigo:
            self.estado_msg.configure(text="")
            messagebox.showwarning("Advertencia", "Debes ingresar un código de proyecto.")
            return

        reader = ExcelReader()
        filas = reader.get_project_rows(codigo)
        # === ORDENAR FILAS POR FECHA ===
        filas_ordenadas = sorted(filas, key=self.fecha_key)
        self.filas_proyecto_actual = filas_ordenadas

        # --- INTEGRAR CABECERA REAL ---
        soap_data = self.get_soap_data(codigo)
        if filas_ordenadas:
            self.project_info = ProjectInfo(soap_data, filas_ordenadas)
            cabecera = self.project_info.to_dict().get("project", {})
            self.nombre_proyecto.set(cabecera.get("Nombre Proyecto", ""))
            self.nombre_beneficiario.set(cabecera.get("Nombre Beneficiario", ""))
            self.nombre_responsable.set(cabecera.get("Representante Legal", ""))

            # Nuevo: combo muestra fecha — tipo — ID
            self.fechas_actuales = [
                f"{DateUtils.format_date(fila.get('Fecha de reunión', ''))} — {str(fila.get('Tipo de reunión', '')).strip()} — {str(fila.get('Id', ''))}"
                for fila in filas_ordenadas
            ]
            if self.accion_var.get() == "Generar registro de visitas":
                self.combo_fichas.configure(values=self.fechas_actuales, state="normal")
            else:
                self.combo_fichas.set("Selecciona una ficha")
                self.combo_fichas.configure(state="disabled")
            self.btn_generar.configure(state="normal")
            self.estado_msg.configure(text="Archivos cargados correctamente.", text_color="#19BC82")
        else:
            self.project_info = None
            self.nombre_proyecto.set("")
            self.nombre_beneficiario.set("")
            self.nombre_responsable.set("")
            self.combo_fichas.set("Selecciona una ficha")
            self.combo_fichas.configure(state="disabled")
            self.btn_generar.configure(state="disabled")
            self.estado_msg.configure(text="")
            messagebox.showwarning("No encontrado", "El proyecto no tiene Fichas Sensor ingresadas.")

    def on_accion_cambio(self, *_):
        # Habilita/deshabilita combo y botón según la acción
        if self.accion_var.get() == "Generar registro de visitas" and self.fechas_actuales:
            self.combo_fichas.configure(values=self.fechas_actuales, state="normal")
            if self.ficha_var.get() in self.fechas_actuales:
                self.btn_generar.configure(state="normal")
            else:
                self.btn_generar.configure(state="disabled")
        else:
            self.combo_fichas.set("Selecciona una ficha")
            self.combo_fichas.configure(state="disabled")
            if self.project_info:
                self.btn_generar.configure(state="normal")
            else:
                self.btn_generar.configure(state="disabled")

    def on_ficha_cambio(self, *_):
        ficha_seleccionada = self.ficha_var.get()
        if ficha_seleccionada and ficha_seleccionada in self.fechas_actuales:
            self.btn_generar.configure(state="normal")
        else:
            self.btn_generar.configure(state="disabled")

    def generar_documento(self):
        accion = self.accion_var.get()
        output_dir = FileUtils.get_downloads_folder()
        codigo = self.codigo_var.get().strip()
        if not self.project_info:
            messagebox.showerror("Error", "Primero debes filtrar las fichas y seleccionar un proyecto válido.")
            return

        if accion == "Generar bitácora de seguimiento":
            template_path = FileUtils.get_template_path("PLANTILLA_BITACORA_DE_SEGUIMIENTO_DEL_PROYECTO.docx")
            generator = BitacoraGenerator(template_path)
            generator.generate(self.project_info, output_dir, codigo)
            messagebox.showinfo("Documento", "Bitácora generada exitosamente.")

        elif accion == "Generar registro de visitas":
            ficha_str = self.ficha_var.get()
            if not ficha_str or ficha_str == "Selecciona una ficha":
                messagebox.showwarning("Atención", "Debes seleccionar una ficha.")
                return
            # Parsear la fecha y el id desde la ficha seleccionada
            # Formato esperado: "2025-03-07 — Reunión virtual — 123"
            partes = ficha_str.split("—")
            fecha_val = partes[0].strip()
            id_val = partes[2].strip() if len(partes) > 2 else ""

            meeting = None
            for fila in self.filas_proyecto_actual:
                fecha_comp = DateUtils.format_date(fila.get("Fecha de reunión", ""))
                id_comp = str(fila.get("Id", "")).strip()
                if fecha_comp == fecha_val and id_comp == id_val:
                    meeting = fila
                    break
            if not meeting:
                messagebox.showwarning("No encontrado", "No se encontró la ficha seleccionada.")
                return
            template_path = FileUtils.get_template_path("PLANTILLA_REGISTRO_DE_VISITAS_Y_REUNIONES.docx")
            generator = SensorSheetGenerator(template_path)
            generator.generate(meeting, self.project_info.to_dict().get("project", {}), output_dir)
            messagebox.showinfo("Documento", f"Registro de visita generado para {fecha_val}.")
        else:
            messagebox.showerror("Error", "Acción no reconocida.")

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
