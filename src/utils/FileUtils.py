import os
import sys
from tkinter import filedialog, messagebox

class FileUtils:
    @staticmethod
    def get_base_dir():
        """
        Devuelve la ruta base correcta tanto en desarrollo como en ejecutable.
        """
        # Si ejecutado como EXE de PyInstaller
        if hasattr(sys, '_MEIPASS'):
            return sys._MEIPASS
        # En desarrollo: .../src/ (ajusta si tu estructura cambia)
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    @staticmethod
    def get_downloads_folder():
        """ Devuelve la ruta a la carpeta de Descargas del usuario.
        """
        # En Windows, la carpeta de Descargas suele estar en C:\Users\<Usuario>\Downloads
        # En Linux y macOS, suele estar en /home/<Usuario>/Downloads o /Users/<Usuario>/Downloads
        return os.path.join(os.path.expanduser('~'), 'Downloads')

    @staticmethod
    def get_fichasensor_excel_path():
        """
        Devuelve la ruta al archivo Excel institucional.
        Si no se encuentra, permite al usuario seleccionarlo manualmente.
        """
        user_folder = os.path.expanduser('~')
        ruta_fija = os.path.join(
            user_folder,
            'OneDrive - corfo.cl',
            'InnovaChile - General',
            'Base Ficha Sensor',
            'Ficha Sensor1.xlsx'
        )
        if os.path.exists(ruta_fija):
            return ruta_fija
        else:
            respuesta = messagebox.askyesno("Archivo no encontrado",
                f"No se encontró el archivo Ficha Sensor.xlsx en:\n\n{ruta_fija}\n\n¿Deseas buscarlo manualmente?")
            if respuesta:
                path = filedialog.askopenfilename(
                    title="Seleccionar archivo Ficha Sensor",
                    filetypes=[("Excel Files", "*.xlsx *.xls")]
                )
                if path:
                    return path
            raise FileNotFoundError("No se pudo localizar el archivo Ficha Sensor.xlsx")


    @staticmethod
    def get_template_path(template_name):
        """ Devuelve la ruta al archivo de plantilla especificado.
        """
        return FileUtils.resolve_resource_path(
            os.path.join("src", "templates", template_name),
            os.path.join("templates", template_name),
        )

    @staticmethod
    def get_assets_folder():
        """ Devuelve la ruta a la carpeta de assets.
        """
        return FileUtils.resolve_resource_path(
            os.path.join("src", "assets"),
            "assets",
        )

    @staticmethod
    def resolve_resource_path(*relative_paths):
        """
        Busca un recurso en rutas candidatas para desarrollo y ejecutable empaquetado.
        """
        base_dir = FileUtils.get_base_dir()
        for rel_path in relative_paths:
            candidate = os.path.join(base_dir, rel_path)
            if os.path.exists(candidate):
                return candidate
        return os.path.join(base_dir, relative_paths[0])

    @staticmethod
    def get_logo_corfo_path():
        """ Devuelve la ruta al logo de Corfo.
        """
        return os.path.join(FileUtils.get_assets_folder(), "logoCorfo.png")

    @staticmethod
    def get_icon_corfo_path():
        """ Devuelve la ruta al icono de Corfo.
        """
        return os.path.join(FileUtils.get_assets_folder(), "Corfo.jpg")
    
    @staticmethod
    def get_ico_corfo_path():
        """ Devuelve la ruta al icono de Corfo.
        """
        return os.path.join(FileUtils.get_assets_folder(), "favicon.ico")
    


# # FUNCIÓN INICIAL PARA LA LECTURA DEL ARCHIVO FICHA SENSOR
# @staticmethod
# def get_fichasensor_excel_path():
#     """
#     Devuelve la ruta al archivo Excel institucional.
#     Si no se encuentra, permite al usuario seleccionarlo manualmente.
#     """
#     user_folder = os.path.expanduser('~')
#     ruta_fija = os.path.join(
#         user_folder,
#         'OneDrive - corfo.cl',
#         'Documentos - SUBDIRECCIÓN DE MEJORA CONTINUA',
#         'Ficha Sensor.xlsx'
#     )
#     if os.path.exists(ruta_fija):
#         return ruta_fija
#     else:
#         respuesta = messagebox.askyesno("Archivo no encontrado",
#             f"No se encontró el archivo Ficha Sensor.xlsx en:\n\n{ruta_fija}\n\n¿Deseas buscarlo manualmente?")
#         if respuesta:
#             path = filedialog.askopenfilename(
#                 title="Seleccionar archivo Ficha Sensor",
#                 filetypes=[("Excel Files", "*.xlsx *.xls")]
#             )
#             if path:
#                 return path
#         raise FileNotFoundError("No se pudo localizar el archivo Ficha Sensor.xlsx")
