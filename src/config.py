#!/usr/bin/env python3
"""
Configuración compartida para los extractores de XML.

Este archivo contiene las configuraciones y constantes compartidas
entre los diferentes extractores (EML y PST).

Autor: Generado automáticamente
Fecha: 2025-10-07
"""

import re
from pathlib import Path

# === CONFIGURACIÓN DE PATRONES XML ===

# Patrón regex para archivos XML de facturación
# Acepta: FE12345.xml, FE-12345.xml, FE_12345.xml (mínimo 3 dígitos)
XML_PATTERN = re.compile(r"^FE[-_]?\d{3,}\.xml$", re.IGNORECASE)

# Patrones adicionales si necesitas otros tipos de facturación
ADDITIONAL_PATTERNS = {
    "NC": re.compile(r"^NC[-_]?\d{3,}\.xml$", re.IGNORECASE),  # Notas de crédito
    "ND": re.compile(r"^ND[-_]?\d{3,}\.xml$", re.IGNORECASE),  # Notas de débito
    "DS": re.compile(r"^DS[-_]?\d{3,}\.xml$", re.IGNORECASE),  # Documentos soporte
}

# === CONFIGURACIÓN DE ARCHIVOS ===

# Nombres de archivos de log por defecto
DEFAULT_LOG_NAMES = {
    "eml": "remitentes_eml.csv",
    "pst": "remitentes_pst.csv",
    "combined": "remitentes_todos.csv"
}

# Encabezados para archivos CSV
CSV_HEADERS = {
    "eml": "archivo_xml,remitente,fecha_email,fecha_procesamiento,archivo_eml_origen",
    "pst": "archivo_xml,remitente,asunto,fecha_email,fecha_procesamiento,carpeta_origen",
    "combined": "archivo_xml,remitente,asunto,fecha_email,fecha_procesamiento,origen,ubicacion_origen"
}

# === CONFIGURACIÓN DE PROCESAMIENTO ===

# Límites de tamaño de archivo
MAX_XML_SIZE_MB = 10  # Tamaño máximo de un archivo XML individual
MAX_LOG_FIELD_LENGTH = 100  # Longitud máxima de campos en CSV

# Configuración de progreso
PROGRESS_UPDATE_INTERVAL = 100  # Actualizar progreso cada N archivos

# === CONFIGURACIÓN DE VALIDACIÓN ===

# Extensiones de archivo válidas
VALID_EMAIL_EXTENSIONS = {".eml", ".msg"}
VALID_PST_EXTENSIONS = {".pst", ".ost"}
VALID_XML_EXTENSIONS = {".xml"}

# === FUNCIONES UTILITARIAS ===

def clean_csv_field(text, max_length=MAX_LOG_FIELD_LENGTH):
    """
    Limpiar un campo para uso en CSV.
    
    Args:
        text (str): Texto a limpiar
        max_length (int): Longitud máxima del campo
        
    Returns:
        str: Texto limpio para CSV
    """
    if not text:
        return "desconocido"
    
    # Convertir a string y limpiar
    clean_text = str(text).replace(",", ";").replace("\n", " ").replace("\r", " ").strip()
    
    # Limitar longitud
    if len(clean_text) > max_length:
        clean_text = clean_text[:max_length-3] + "..."
    
    return clean_text if clean_text else "desconocido"


def is_xml_filename_valid(filename, additional_patterns=None):
    """
    Verificar si un nombre de archivo XML cumple con los patrones válidos.
    
    Args:
        filename (str): Nombre del archivo a verificar
        additional_patterns (dict): Patrones adicionales a verificar
        
    Returns:
        bool: True si el archivo es válido
    """
    if not filename:
        return False
    
    clean_name = filename.strip()
    
    # Verificar patrón principal
    if XML_PATTERN.match(clean_name):
        return True
    
    # Verificar patrones adicionales si se proporcionan
    if additional_patterns:
        for pattern in additional_patterns.values():
            if pattern.match(clean_name):
                return True
    
    return False


def create_unique_filename(output_dir, filename):
    """
    Crear un nombre de archivo único si ya existe.
    
    Args:
        output_dir (Path): Directorio de salida
        filename (str): Nombre del archivo original
        
    Returns:
        Path: Ruta completa del archivo único
    """
    output_path = Path(output_dir) / filename
    
    if not output_path.exists():
        return output_path
    
    # Generar nombre único
    name_parts = filename.rsplit('.', 1)
    base_name = name_parts[0]
    extension = name_parts[1] if len(name_parts) > 1 else ""
    
    counter = 1
    while output_path.exists():
        if extension:
            new_filename = f"{base_name}_{counter:03d}.{extension}"
        else:
            new_filename = f"{base_name}_{counter:03d}"
        
        output_path = Path(output_dir) / new_filename
        counter += 1
        
        # Evitar bucle infinito
        if counter > 9999:
            raise Exception(f"No se pudo crear nombre único para {filename}")
    
    return output_path


def validate_file_size(file_path, max_size_mb=MAX_XML_SIZE_MB):
    """
    Validar que un archivo no exceda el tamaño máximo.
    
    Args:
        file_path (Path): Ruta del archivo
        max_size_mb (int): Tamaño máximo en MB
        
    Returns:
        bool: True si el archivo es válido
    """
    try:
        file_size_mb = file_path.stat().st_size / (1024 * 1024)
        return file_size_mb <= max_size_mb
    except Exception:
        return False


# === CONFIGURACIÓN POR DEFECTO ===

DEFAULT_CONFIG = {
    "xml_pattern": XML_PATTERN,
    "additional_patterns": ADDITIONAL_PATTERNS,
    "max_xml_size_mb": MAX_XML_SIZE_MB,
    "max_log_field_length": MAX_LOG_FIELD_LENGTH,
    "progress_update_interval": PROGRESS_UPDATE_INTERVAL,
    "valid_extensions": {
        "email": VALID_EMAIL_EXTENSIONS,
        "pst": VALID_PST_EXTENSIONS,
        "xml": VALID_XML_EXTENSIONS
    },
    "csv_headers": CSV_HEADERS,
    "log_names": DEFAULT_LOG_NAMES
}


# === EJEMPLOS DE USO ===

if __name__ == "__main__":
    # Ejemplos de cómo usar las funciones de configuración
    
    print("🧪 Probando configuraciones...")
    
    # Probar patrones XML
    test_files = [
        "FE12345.xml",     # ✅ Válido
        "FE-12345.xml",    # ✅ Válido  
        "FE_12345.xml",    # ✅ Válido
        "NC123.xml",       # ✅ Válido (nota de crédito)
        "FE12.xml",        # ❌ Inválido (menos de 3 dígitos)
        "FE12345.pdf",     # ❌ Inválido (no es XML)
        "OtroArchivo.xml", # ❌ Inválido (no empieza con FE)
    ]
    
    print("\n🎯 Prueba de patrones XML:")
    for filename in test_files:
        is_valid = is_xml_filename_valid(filename, ADDITIONAL_PATTERNS)
        status = "✅ Válido" if is_valid else "❌ Inválido"
        print(f"   {filename:15} → {status}")
    
    # Probar limpieza de campos CSV
    print("\n🧹 Prueba de limpieza de campos CSV:")
    test_texts = [
        "Texto normal",
        "Texto, con comas, múltiples",
        "Texto\ncon\nsaltos\nde\nlínea",
        "Texto muy largo que excede la longitud máxima permitida para campos CSV y debería ser truncado automáticamente",
        None,
        ""
    ]
    
    for text in test_texts:
        clean = clean_csv_field(text, 50)
        print(f"   Original: {repr(text)}")
        print(f"   Limpio:   {repr(clean)}")
        print()
    
    print("✅ Configuración validada correctamente.")