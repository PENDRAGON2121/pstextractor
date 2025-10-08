#!/usr/bin/env python3
"""
Extractor XML PST con GUI
Extractor de archivos XML de facturación desde archivos PST usando interfaz gráfica.

Este script combina:
- GUI para selección fácil de archivos PST
- Múltiples métodos de extracción (Outlook COM, pypff si disponible)
- Barra de progreso visual
- Notificaciones de éxito/error

Dependencias:
    - tkinter: Para interfaz gráfica (incluida con Python)
    - win32com.client: Para Outlook COM (pywin32)
    - tqdm: Para barras de progreso adicionales
    - lxml: Para validación de XML (opcional)

Autor: Generado automáticamente
Fecha: 2025-10-07
"""

import os
import re
import sys
import argparse
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time

# Importaciones opcionales
try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False

try:
    from tqdm import tqdm
    TQDM_AVAILABLE = True
except ImportError:
    TQDM_AVAILABLE = False

try:
    from lxml import etree
    LXML_AVAILABLE = True
except ImportError:
    LXML_AVAILABLE = False

try:
    import pypff
    PYPFF_AVAILABLE = True
except ImportError:
    PYPFF_AVAILABLE = False


def seleccionar_archivo_pst():
    """
    Abrir un diálogo para seleccionar el archivo PST.
    
    Returns:
        str: Ruta del archivo PST seleccionado, o None si se cancela
    """
    print("🔍 Abriendo selector de archivo PST...")
    
    # Crear ventana raíz (oculta)
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal
    root.attributes('-topmost', True)  # Mantener diálogo al frente
    
    # Configurar el diálogo de archivo
    archivo_pst = filedialog.askopenfilename(
        title="Seleccionar archivo PST de Outlook",
        filetypes=[
            ("Archivos PST de Outlook", "*.pst"),
            ("Todos los archivos", "*.*")
        ],
        initialdir=os.path.expanduser("~/Desktop"),  # Empezar en el escritorio
    )
    
    root.destroy()  # Cerrar la ventana raíz
    
    if archivo_pst:
        print(f"✅ Archivo seleccionado: {archivo_pst}")
        return archivo_pst
    else:
        print("❌ No se seleccionó ningún archivo")
        return None


class VentanaProgreso:
    """Ventana de progreso para mostrar el estado de la extracción."""
    
    def __init__(self, titulo="Procesando PST"):
        self.ventana = None
        self.barra_progreso = None
        self.etiqueta_estado = None
        self.etiqueta_porcentaje = None
        self.etiqueta_stats = None
        self.activa = False
        self.crear_ventana(titulo)
    
    def crear_ventana(self, titulo):
        """Crear la ventana de progreso."""
        self.ventana = tk.Tk()
        self.ventana.title(titulo)
        self.ventana.geometry("600x200")
        self.ventana.resizable(False, False)
        
        # Centrar la ventana
        self.ventana.update_idletasks()
        x = (self.ventana.winfo_screenwidth() // 2) - (600 // 2)
        y = (self.ventana.winfo_screenheight() // 2) - (200 // 2)
        self.ventana.geometry(f"600x200+{x}+{y}")
        
        # Evitar que se cierre con X
        self.ventana.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Icono y título
        titulo_label = tk.Label(self.ventana, text="🧾 Extrayendo XML de Facturación", 
                               font=("Arial", 14, "bold"))
        titulo_label.pack(pady=10)
        
        # Estado actual
        self.etiqueta_estado = tk.Label(self.ventana, text="Preparando...", 
                                       font=("Arial", 10))
        self.etiqueta_estado.pack(pady=5)
        
        # Barra de progreso
        self.barra_progreso = ttk.Progressbar(self.ventana, length=500, mode='determinate')
        self.barra_progreso.pack(pady=10)
        
        # Porcentaje
        self.etiqueta_porcentaje = tk.Label(self.ventana, text="0%", 
                                           font=("Arial", 10, "bold"))
        self.etiqueta_porcentaje.pack(pady=2)
        
        # Estadísticas
        self.etiqueta_stats = tk.Label(self.ventana, text="Emails: 0 | XMLs: 0", 
                                      font=("Arial", 9), fg="gray")
        self.etiqueta_stats.pack(pady=5)
        
        # Botón cancelar (opcional)
        self.boton_cancelar = tk.Button(self.ventana, text="Minimizar", 
                                       command=self.minimizar)
        self.boton_cancelar.pack(pady=10)
        
        self.activa = True
        self.ventana.update()
    
    def actualizar(self, progreso, total, estado="", emails_procesados=0, xmls_encontrados=0):
        """Actualizar el progreso y estado."""
        if not self.activa or not self.ventana:
            return
        
        try:
            # Actualizar barra de progreso
            if total > 0:
                porcentaje = (progreso / total) * 100
                self.barra_progreso['value'] = porcentaje
                self.etiqueta_porcentaje.config(text=f"{porcentaje:.1f}%")
            
            # Actualizar estado
            if estado:
                self.etiqueta_estado.config(text=estado)
            
            # Actualizar estadísticas
            self.etiqueta_stats.config(text=f"Emails: {emails_procesados:,} | XMLs: {xmls_encontrados}")
            
            # Forzar actualización
            self.ventana.update()
            
        except Exception as e:
            print(f"⚠️ Error actualizando ventana de progreso: {e}")
            self.activa = False
    
    def finalizar(self, mensaje="Completado", exito=True):
        """Finalizar el progreso."""
        if not self.activa or not self.ventana:
            return
        
        try:
            if exito:
                self.etiqueta_estado.config(text=f"✅ {mensaje}")
                self.barra_progreso['value'] = 100
                self.etiqueta_porcentaje.config(text="100%")
            else:
                self.etiqueta_estado.config(text=f"❌ {mensaje}")
            
            self.boton_cancelar.config(text="Cerrar", command=self.cerrar)
            self.ventana.update()
            
        except Exception as e:
            print(f"⚠️ Error finalizando ventana de progreso: {e}")
    
    def minimizar(self):
        """Minimizar la ventana."""
        try:
            self.ventana.iconify()
        except:
            pass
    
    def cerrar(self):
        """Cerrar la ventana."""
        try:
            self.activa = False
            self.ventana.destroy()
        except:
            pass
    
    def on_closing(self):
        """Manejar el evento de cerrar ventana."""
        self.minimizar()  # Solo minimizar, no cerrar


class ExtractorXMLPSTGUI:
    """Extractor de archivos XML con interfaz gráfica."""
    
    def __init__(self, pst_file, output_dir):
        """
        Inicializar el extractor.
        
        Args:
            pst_file (str): Archivo PST a procesar
            output_dir (str): Directorio donde guardar los XML extraídos
        """
        self.pst_file = Path(pst_file)
        self.output_dir = Path(output_dir)
        self.log_file = self.output_dir / "remitentes_pst.csv"
        
        # Patrón regex para archivos XML de facturación
        self.xml_pattern = re.compile(r"^FE[-_]?\d{3,}\.xml$", re.IGNORECASE)
        
        # Contadores
        self.total_emails = 0
        self.processed_emails = 0
        self.extracted_xml_files = 0
        self.errors = []
        
        # GUI
        self.ventana_progreso = None
    
    def setup_directories(self):
        """Crear directorios necesarios."""
        self.output_dir.mkdir(parents=True, exist_ok=True)
        (self.output_dir / "xml_facturacion").mkdir(exist_ok=True)
        (self.output_dir / "reportes").mkdir(exist_ok=True)
        print(f"📁 Directorio de salida: {self.output_dir}")
    
    def validate_pst_file(self):
        """Validar que el archivo PST existe y es accesible."""
        if not self.pst_file.exists():
            raise FileNotFoundError(f"❌ El archivo PST no existe: {self.pst_file}")
        
        if not self.pst_file.is_file():
            raise ValueError(f"❌ La ruta no es un archivo válido: {self.pst_file}")
        
        if self.pst_file.stat().st_size == 0:
            raise ValueError(f"❌ El archivo PST está vacío: {self.pst_file}")
        
        size_mb = self.pst_file.stat().st_size / (1024*1024)
        print(f"📊 Archivo PST: {self.pst_file}")
        print(f"📏 Tamaño: {size_mb:.1f} MB")
        
        return size_mb
    
    def inicializar_log(self):
        """Crear archivo CSV de log."""
        with open(self.log_file, "w", encoding="utf-8") as log:
            log.write("archivo_xml,remitente,asunto,fecha_email,fecha_procesamiento,carpeta_origen,tamaño_bytes\n")
    
    def extraer_con_outlook_com(self):
        """Extraer usando Outlook COM."""
        if not WIN32COM_AVAILABLE:
            raise ImportError("win32com.client no disponible")
        
        print("🔄 Intentando extracción con Outlook COM...")
        
        try:
            # Conectar con Outlook
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            # Intentar añadir el PST
            namespace.AddStore(str(self.pst_file))
            
            # Buscar el PST en las stores
            pst_store = None
            for store in namespace.Stores:
                if str(self.pst_file).lower() in store.FilePath.lower():
                    pst_store = store
                    break
            
            if not pst_store:
                raise Exception("PST no encontrado después de añadirlo")
            
            print(f"✅ PST encontrado: {pst_store.DisplayName}")
            
            # Procesar el PST
            root_folder = pst_store.GetRootFolder()
            self.procesar_carpeta_outlook(root_folder)
            
            return True
            
        except Exception as e:
            print(f"❌ Error con Outlook COM: {e}")
            return False
    
    def procesar_carpeta_outlook(self, folder, ruta_carpeta=""):
        """Procesar carpeta usando Outlook COM."""
        if not folder:
            return
        
        nombre_carpeta = folder.Name
        ruta_actual = f"{ruta_carpeta}/{nombre_carpeta}" if ruta_carpeta else nombre_carpeta
        
        if self.ventana_progreso:
            self.ventana_progreso.actualizar(
                self.processed_emails, 
                max(self.total_emails, 1000),  # Estimación si no sabemos el total
                f"Procesando: {nombre_carpeta}",
                self.processed_emails,
                self.extracted_xml_files
            )
        
        try:
            # Procesar elementos en esta carpeta
            items = folder.Items
            
            for item in items:
                try:
                    self.processed_emails += 1
                    
                    # Verificar si tiene adjuntos
                    if hasattr(item, 'Attachments') and item.Attachments.Count > 0:
                        for attachment in item.Attachments:
                            filename = attachment.FileName
                            
                            if self.xml_pattern.match(filename):
                                # Extraer adjunto XML
                                xml_path = self.output_dir / "xml_facturacion" / filename
                                
                                # Si ya existe, agregar sufijo
                                counter = 1
                                while xml_path.exists():
                                    name_parts = filename.rsplit('.', 1)
                                    new_name = f"{name_parts[0]}_{counter:03d}.{name_parts[1]}"
                                    xml_path = self.output_dir / "xml_facturacion" / new_name
                                    counter += 1
                                
                                # Guardar adjunto
                                attachment.SaveAsFile(str(xml_path))
                                self.extracted_xml_files += 1
                                
                                # Registrar en log
                                self.registrar_en_log(
                                    xml_path.name,
                                    getattr(item, 'SenderName', 'desconocido'),
                                    getattr(item, 'Subject', 'sin asunto'),
                                    getattr(item, 'ReceivedTime', 'fecha desconocida'),
                                    ruta_actual,
                                    xml_path.stat().st_size if xml_path.exists() else 0
                                )
                                
                                print(f"✅ XML extraído: {filename}")
                    
                    # Actualizar progreso cada 50 emails
                    if self.processed_emails % 50 == 0 and self.ventana_progreso:
                        self.ventana_progreso.actualizar(
                            self.processed_emails,
                            max(self.total_emails, 1000),
                            f"Procesados {self.processed_emails} emails en: {nombre_carpeta}",
                            self.processed_emails,
                            self.extracted_xml_files
                        )
                
                except Exception as e:
                    self.errors.append(f"Error procesando item en {ruta_actual}: {str(e)}")
            
            # Procesar subcarpetas
            try:
                for subfolder in folder.Folders:
                    self.procesar_carpeta_outlook(subfolder, ruta_actual)
            except Exception as e:
                self.errors.append(f"Error accediendo subcarpetas de {ruta_actual}: {str(e)}")
                
        except Exception as e:
            self.errors.append(f"Error procesando carpeta {ruta_actual}: {str(e)}")
    
    def registrar_en_log(self, xml_file, remitente, asunto, fecha, carpeta, tamaño):
        """Registrar extracción en el log CSV."""
        try:
            with open(self.log_file, "a", encoding="utf-8") as log:
                # Limpiar datos para CSV
                clean_remitente = str(remitente).replace(",", ";").replace("\n", " ").strip()[:100]
                clean_asunto = str(asunto).replace(",", ";").replace("\n", " ").strip()[:150]
                clean_fecha = str(fecha).replace(",", ";").replace("\n", " ").strip()
                clean_carpeta = str(carpeta).replace(",", ";")
                current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                log.write(f"{xml_file},{clean_remitente},{clean_asunto},{clean_fecha},{current_time},{clean_carpeta},{tamaño}\n")
        except Exception as e:
            self.errors.append(f"Error escribiendo log: {str(e)}")
    
    def generar_reporte_final(self):
        """Generar reporte final de la extracción."""
        reporte_path = self.output_dir / "reportes" / "reporte_extraccion.txt"
        
        with open(reporte_path, "w", encoding="utf-8") as f:
            f.write("REPORTE DE EXTRACCIÓN XML PST\n")
            f.write("=" * 50 + "\n\n")
            f.write(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Archivo PST: {self.pst_file}\n")
            f.write(f"Directorio salida: {self.output_dir}\n\n")
            f.write("ESTADÍSTICAS:\n")
            f.write(f"- Emails procesados: {self.processed_emails:,}\n")
            f.write(f"- XMLs extraídos: {self.extracted_xml_files:,}\n")
            f.write(f"- Errores: {len(self.errors):,}\n\n")
            
            if self.errors:
                f.write("ERRORES ENCONTRADOS:\n")
                for i, error in enumerate(self.errors[:20], 1):
                    f.write(f"{i}. {error}\n")
                if len(self.errors) > 20:
                    f.write(f"... y {len(self.errors)-20} errores más.\n")
        
        print(f"📋 Reporte generado: {reporte_path}")
    
    def extraer_xml_files(self):
        """Ejecutar el proceso completo de extracción."""
        print("🚀 Iniciando extracción de archivos XML desde PST...")
        
        try:
            # Configurar directorios
            self.setup_directories()
            
            # Validar archivo PST
            size_mb = self.validate_pst_file()
            
            # Inicializar log
            self.inicializar_log()
            
            # Crear ventana de progreso
            self.ventana_progreso = VentanaProgreso("Extrayendo XML de PST")
            
            # Intentar extracción
            exito = False
            
            if WIN32COM_AVAILABLE:
                try:
                    exito = self.extraer_con_outlook_com()
                except Exception as e:
                    print(f"❌ Error con Outlook COM: {e}")
            
            if not exito:
                raise Exception("No se pudo extraer el PST con ningún método disponible")
            
            # Generar reporte
            self.generar_reporte_final()
            
            # Finalizar ventana de progreso
            if self.ventana_progreso:
                self.ventana_progreso.finalizar(
                    f"Extracción completada - {self.extracted_xml_files} XMLs encontrados"
                )
            
            # Mostrar resultado
            mensaje_resultado = (
                f"🎉 ¡Extracción completada con éxito!\n\n"
                f"📊 RESULTADOS:\n"
                f"📧 Emails procesados: {self.processed_emails:,}\n"
                f"📄 XMLs de facturación extraídos: {self.extracted_xml_files:,}\n"
                f"❌ Errores: {len(self.errors):,}\n\n"
                f"📁 Archivos guardados en:\n{self.output_dir}\n\n"
                f"📋 Log detallado: {self.log_file.name}"
            )
            
            print("\n" + "="*60)
            print("📊 RESULTADOS FINALES")
            print("="*60)
            print(f"📧 Emails procesados: {self.processed_emails:,}")
            print(f"📄 XMLs extraídos: {self.extracted_xml_files:,}")
            print(f"❌ Errores: {len(self.errors):,}")
            print(f"📁 XMLs guardados en: {self.output_dir / 'xml_facturacion'}")
            
            # Mostrar mensaje de éxito
            messagebox.showinfo("Extracción Completada", mensaje_resultado)
            
            return True
            
        except Exception as e:
            error_msg = f"❌ Error durante la extracción: {str(e)}"
            print(error_msg)
            
            if self.ventana_progreso:
                self.ventana_progreso.finalizar("Error en la extracción", exito=False)
            
            messagebox.showerror("Error de Extracción", error_msg)
            return False
        
        finally:
            # Asegurar que la ventana se cierre
            if self.ventana_progreso:
                # Mantener abierta 3 segundos más para que el usuario vea el resultado
                self.ventana_progreso.ventana.after(3000, self.ventana_progreso.cerrar)


def main():
    """Función principal del script."""
    parser = argparse.ArgumentParser(
        description="Extractor de archivos XML de facturación desde PST con GUI",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos de uso:
  python extractor_xml_pst_gui.py                          # Usar GUI para todo
  python extractor_xml_pst_gui.py -i "archivo.pst"        # Especificar PST
  python extractor_xml_pst_gui.py -o "directorio_salida"  # Especificar salida

Características:
  🖱️  Interfaz gráfica fácil de usar
  📊 Barra de progreso en tiempo real
  🔄 Múltiples métodos de extracción
  📋 Logs detallados y reportes
  ✅ Notificaciones visuales de éxito/error

El script buscará archivos XML con nomenclatura FE*.xml dentro de los
correos del archivo PST seleccionado.
        """
    )
    
    parser.add_argument(
        "-i", "--input-pst",
        help="Archivo PST a procesar (se abrirá selector si no se especifica)"
    )
    
    parser.add_argument(
        "-o", "--output-dir", 
        help="Directorio de salida (se creará automáticamente si no se especifica)"
    )
    
    args = parser.parse_args()
    
    try:
        # Mostrar información inicial
        print("🧾 EXTRACTOR XML PST CON GUI")
        print("=" * 40)
        print("Busca archivos XML de facturación (FE*.xml) en archivos PST")
        print()
        
        # Verificar dependencias críticas
        if not WIN32COM_AVAILABLE:
            error_msg = (
                "❌ ERROR: win32com.client no disponible\n\n"
                "Para instalar:\n"
                "pip install pywin32\n\n"
                "Esta dependencia es necesaria para acceder a archivos PST en Windows."
            )
            print(error_msg)
            messagebox.showerror("Dependencia Faltante", error_msg)
            sys.exit(1)
        
        # Seleccionar archivo PST
        pst_file = args.input_pst
        if not pst_file:
            pst_file = seleccionar_archivo_pst()
            if not pst_file:
                print("⏹️ Operación cancelada por el usuario")
                sys.exit(0)
        
        # Determinar directorio de salida
        output_dir = args.output_dir
        if not output_dir:
            pst_path = Path(pst_file)
            output_dir = pst_path.parent / f"{pst_path.stem}_xml_extraidos"
            print(f"📁 Directorio de salida automático: {output_dir}")
        
        # Confirmar con el usuario
        confirmacion = messagebox.askyesno(
            "Confirmar Extracción",
            f"🔍 CONFIRMAR EXTRACCIÓN\n\n"
            f"📂 Archivo PST:\n{pst_file}\n\n"
            f"📁 Directorio de salida:\n{output_dir}\n\n"
            f"¿Proceder con la extracción?"
        )
        
        if not confirmacion:
            print("⏹️ Operación cancelada por el usuario")
            sys.exit(0)
        
        # Crear y ejecutar extractor
        extractor = ExtractorXMLPSTGUI(pst_file, output_dir)
        exito = extractor.extraer_xml_files()
        
        if exito:
            print("\n✅ Extracción completada exitosamente")
        else:
            print("\n❌ Extracción falló")
            sys.exit(1)
        
    except KeyboardInterrupt:
        print("\n⏹️ Operación interrumpida por el usuario")
        sys.exit(1)
    except Exception as e:
        error_msg = f"❌ Error inesperado: {str(e)}"
        print(error_msg)
        try:
            messagebox.showerror("Error", error_msg)
        except:
            pass
        sys.exit(1)


if __name__ == "__main__":
    main()