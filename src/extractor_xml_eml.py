#!/usr/bin/env python3
"""
Extractor de archivos XML de facturación desde correos EML.

Este script extrae archivos XML con nomenclatura FE#####.xml de correos .eml,
optimizado para procesar grandes volúmenes de archivos.

Autor: Generado automáticamente
Fecha: 2025-10-07
"""

import os
import re
import email
import sys
import argparse
from email import policy
from datetime import datetime
from pathlib import Path
from tqdm import tqdm
from lxml import etree


class ExtractorXMLEML:
    """Extractor de archivos XML de facturación desde correos EML."""
    
    def __init__(self, input_dir, output_dir):
        """
        Inicializar el extractor.
        
        Args:
            input_dir (str): Directorio con archivos EML
            output_dir (str): Directorio donde guardar los XML extraídos
        """
        self.input_dir = Path(input_dir)
        self.output_dir = Path(output_dir)
        self.log_file = self.output_dir / "remitentes.csv"
        
        # Patrón regex para archivos XML de facturación
        # Acepta: FE12345.xml, FE-12345.xml, FE_12345.xml (mínimo 3 dígitos)
        self.xml_pattern = re.compile(r"^FE[-_]?\d{3,}\.xml$", re.IGNORECASE)
        
        # Contadores para estadísticas
        self.total_eml_files = 0
        self.processed_eml_files = 0
        self.extracted_xml_files = 0
        self.errors = []
    
    def setup_directories(self):
        """Crear directorios necesarios."""
        self.output_dir.mkdir(parents=True, exist_ok=True)
        print(f"📁 Directorio de salida: {self.output_dir}")
    
    def validate_input_directory(self):
        """Validar que el directorio de entrada existe y contiene archivos EML."""
        if not self.input_dir.exists():
            raise FileNotFoundError(f"❌ El directorio de entrada no existe: {self.input_dir}")
        
        eml_files = list(self.input_dir.glob("*.eml"))
        if not eml_files:
            raise ValueError(f"❌ No se encontraron archivos EML en: {self.input_dir}")
        
        self.total_eml_files = len(eml_files)
        print(f"📊 Archivos EML encontrados: {self.total_eml_files:,}")
        return eml_files
    
    def initialize_log_file(self):
        """Crear archivo CSV de log con encabezados."""
        with open(self.log_file, "w", encoding="utf-8") as log:
            log.write("archivo_xml,remitente,fecha_email,fecha_procesamiento,archivo_eml_origen\n")
        print(f"📋 Log de remitentes: {self.log_file}")
    
    def process_eml_file(self, eml_path):
        """
        Procesar un archivo EML individual.
        
        Args:
            eml_path (Path): Ruta del archivo EML
            
        Returns:
            int: Número de archivos XML extraídos de este EML
        """
        extracted_count = 0
        
        try:
            # Leer y parsear el archivo EML
            with open(eml_path, "rb") as f:
                msg = email.message_from_binary_file(f, policy=policy.default)
            
            # Obtener información del correo
            sender = msg.get("From", "desconocido")
            date = msg.get("Date", "fecha_desconocida")
            subject = msg.get("Subject", "sin_asunto")
            
            self.processed_eml_files += 1
            
            # Examinar adjuntos del correo
            for part in msg.iter_attachments():
                filename = part.get_filename()
                if not filename:
                    continue
                
                # Limpiar el nombre del archivo
                clean_name = filename.strip()
                
                # Verificar si cumple con el patrón FE#####.xml
                if self.xml_pattern.match(clean_name):
                    try:
                        # Extraer el contenido del adjunto
                        xml_content = part.get_payload(decode=True)
                        
                        if xml_content:
                            # Guardar el archivo XML
                            xml_output_path = self.output_dir / clean_name
                            
                            # Si el archivo ya existe, agregar un sufijo único
                            counter = 1
                            while xml_output_path.exists():
                                name_parts = clean_name.rsplit('.', 1)
                                new_name = f"{name_parts[0]}_{counter:03d}.{name_parts[1]}"
                                xml_output_path = self.output_dir / new_name
                                counter += 1
                            
                            with open(xml_output_path, "wb") as xml_file:
                                xml_file.write(xml_content)
                            
                            # Registrar en el log CSV
                            with open(self.log_file, "a", encoding="utf-8") as log:
                                # Limpiar datos para CSV (quitar comas y saltos de línea)
                                clean_sender = sender.replace(",", ";").replace("\n", " ").strip()
                                clean_date = str(date).replace(",", ";").replace("\n", " ").strip()
                                clean_subject = subject.replace(",", ";").replace("\n", " ").strip()[:100]  # Limitar longitud
                                current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                
                                log.write(f"{xml_output_path.name},{clean_sender},{clean_date},{current_time},{eml_path.name}\n")
                            
                            extracted_count += 1
                            self.extracted_xml_files += 1
                            
                    except Exception as e:
                        error_msg = f"Error extrayendo {clean_name} de {eml_path.name}: {str(e)}"
                        self.errors.append(error_msg)
                        
        except Exception as e:
            error_msg = f"Error procesando {eml_path.name}: {str(e)}"
            self.errors.append(error_msg)
        
        return extracted_count
    
    def validate_xml_files(self):
        """Validar que todos los archivos XML extraídos sean válidos."""
        print("\n🔍 Validando archivos XML extraídos...")
        print("=" * 60)
        
        xml_files = list(self.output_dir.glob("*.xml"))
        
        if not xml_files:
            print("⚠️ No se encontraron archivos XML para validar.")
            return
        
        valid_files = []
        invalid_files = []
        
        # Validar cada archivo XML
        for xml_file in tqdm(xml_files, desc="🔍 Validando XML"):
            try:
                # Intentar parsear el XML
                etree.parse(xml_file)
                valid_files.append(xml_file.name)
            except Exception as e:
                invalid_files.append((xml_file.name, str(e)))
        
        # Mostrar resultados de la validación
        print(f"\n📊 RESULTADOS DE VALIDACIÓN:")
        print(f"✅ Archivos XML válidos:    {len(valid_files):,}")
        print(f"❌ Archivos XML inválidos:  {len(invalid_files):,}")
        print(f"📄 Total archivos XML:      {len(xml_files):,}")
        
        # Calcular porcentaje de éxito
        if xml_files:
            success_rate = (len(valid_files) / len(xml_files)) * 100
            print(f"📈 Tasa de éxito:           {success_rate:.1f}%")
        
        # Mostrar archivos inválidos si los hay
        if invalid_files:
            print(f"\n⚠️ ARCHIVOS XML INVÁLIDOS:")
            for i, (filename, error) in enumerate(invalid_files[:5]):
                print(f"   {i+1}. {filename}")
                print(f"      Error: {error[:100]}{'...' if len(error) > 100 else ''}")
            
            if len(invalid_files) > 5:
                print(f"   ... y {len(invalid_files)-5} archivos más con errores.")
        else:
            print(f"\n🎉 ¡Excelente! Todos los archivos XML son válidos y bien formados.")
    
    def print_statistics(self):
        """Mostrar estadísticas del procesamiento."""
        print("\n" + "=" * 60)
        print("📊 RESULTADOS DEL PROCESAMIENTO")
        print("=" * 60)
        print(f"📧 Archivos EML totales:     {self.total_eml_files:,}")
        print(f"✅ Archivos EML procesados:  {self.processed_eml_files:,}")
        print(f"📄 Archivos XML extraídos:   {self.extracted_xml_files:,}")
        print(f"❌ Errores encontrados:      {len(self.errors):,}")
        print()
        print(f"📁 XML guardados en:         {self.output_dir}")
        print(f"📋 Log de remitentes:        {self.log_file}")
        
        # Mostrar errores si los hay
        if self.errors:
            print("\n⚠️ ERRORES ENCONTRADOS:")
            for i, error in enumerate(self.errors[:10]):  # Mostrar solo los primeros 10
                print(f"   {i+1}. {error}")
            if len(self.errors) > 10:
                print(f"   ... y {len(self.errors)-10} errores más.")
    
    def extract_xml_files(self, validate_xml=True):
        """
        Ejecutar el proceso completo de extracción.
        
        Args:
            validate_xml (bool): Si validar los XML extraídos
        """
        print("🚀 Iniciando extracción de archivos XML desde EML...")
        print(f"📂 Directorio de entrada: {self.input_dir}")
        print(f"📂 Directorio de salida:  {self.output_dir}")
        print()
        
        # Configurar directorios
        self.setup_directories()
        
        # Validar directorio de entrada
        eml_files = self.validate_input_directory()
        
        # Inicializar archivo de log
        self.initialize_log_file()
        
        print("🚀 Iniciando procesamiento...")
        print()
        
        # Procesar cada archivo EML con barra de progreso
        for eml_file in tqdm(eml_files, desc="📧 Procesando EML", unit="archivos"):
            self.process_eml_file(eml_file)
        
        # Mostrar estadísticas
        self.print_statistics()
        
        # Validar archivos XML si se solicita
        if validate_xml:
            self.validate_xml_files()
        
        print("\n✅ Procesamiento completado con éxito.")


def main():
    """Función principal del script."""
    parser = argparse.ArgumentParser(
        description="Extractor de archivos XML de facturación desde correos EML",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos de uso:
  python extractor_xml_eml.py -i "C:/correos/eml_files" -o "C:/correos/xml_extracts"
  python extractor_xml_eml.py -i "/home/user/eml" -o "/home/user/xml" --no-validate
        """
    )
    
    parser.add_argument(
        "-i", "--input-dir",
        required=True,
        help="Directorio con archivos EML a procesar"
    )
    
    parser.add_argument(
        "-o", "--output-dir",
        required=True,
        help="Directorio donde guardar los archivos XML extraídos"
    )
    
    parser.add_argument(
        "--no-validate",
        action="store_true",
        help="No validar los archivos XML extraídos (más rápido)"
    )
    
    args = parser.parse_args()
    
    try:
        # Crear y ejecutar el extractor
        extractor = ExtractorXMLEML(args.input_dir, args.output_dir)
        extractor.extract_xml_files(validate_xml=not args.no_validate)
        
    except Exception as e:
        print(f"❌ Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()