#!/usr/bin/env python3
"""
Script para renombrar archivos XML según el valor del tag <Clave>

Lee cada archivo XML, extrae el contenido del tag <Clave> y renombra
el archivo con ese valor, conservando la extensión .xml

Ejemplo:
    Archivo: 01_50609042500011038073200100001010000015582100001234.XML
    Tag: <Clave>50609042500011038073200100001010000015582100001234</Clave>
    Resultado: 50609042500011038073200100001010000015582100001234.xml

Autor: Generado automáticamente
Fecha: 2025-10-13
"""

import os
from pathlib import Path
import xml.etree.ElementTree as ET
import argparse
import tkinter as tk
from tkinter import filedialog

def seleccionar_carpeta(titulo):
    """Abrir diálogo para seleccionar carpeta."""
    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes('-topmost', True)
    except Exception:
        pass
    root.update()
    carpeta = filedialog.askdirectory(title=titulo)
    root.destroy()
    return carpeta

def extraer_clave_xml(xml_path: Path) -> str:
    """
    Extraer el valor del tag <Clave> de un archivo XML.
    
    Args:
        xml_path: Ruta al archivo XML
        
    Returns:
        Contenido del tag <Clave> o cadena vacía si no se encuentra
    """
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        
        # Buscar el tag Clave con o sin namespace
        clave = root.find('.//{*}Clave')
        if clave is not None and clave.text:
            return clave.text.strip()
        
        # Intentar sin namespace
        clave = root.find('.//Clave')
        if clave is not None and clave.text:
            return clave.text.strip()
            
        return ""
    except ET.ParseError as e:
        print(f"Error parseando XML {xml_path.name}: {e}", flush=True)
        return ""
    except Exception as e:
        print(f"Error leyendo {xml_path.name}: {e}", flush=True)
        return ""

def sanitizar_nombre_archivo(nombre: str) -> str:
    """
    Sanitizar el nombre de archivo removiendo caracteres inválidos.
    
    Args:
        nombre: Nombre a sanitizar
        
    Returns:
        Nombre sanitizado
    """
    # Caracteres inválidos en Windows
    invalidos = '<>:"/\\|?*'
    for char in invalidos:
        nombre = nombre.replace(char, '_')
    return nombre.strip()

def renombrar_xml_por_clave(input_dir: str, dry_run: bool = False):
    """
    Renombrar todos los archivos XML en el directorio según su tag <Clave>.
    Los duplicados se mueven a una carpeta 'Copias' dentro de cada subdirectorio.
    
    Args:
        input_dir: Directorio con los archivos XML
        dry_run: Si es True, solo muestra lo que haría sin renombrar
    """
    base_dir = Path(input_dir)
    
    if not base_dir.exists():
        print(f"❌ El directorio {input_dir} no existe.")
        return
    
    # Buscar todos los archivos XML recursivamente
    xml_files = list(base_dir.rglob("*.xml")) + list(base_dir.rglob("*.XML"))
    
    if not xml_files:
        print("❌ No se encontraron archivos XML en el directorio.")
        return
    
    print(f"📁 Procesando {len(xml_files)} archivos XML...")
    if dry_run:
        print("⚠️  MODO PRUEBA - No se renombrará ni moverá ningún archivo")
    print()
    
    renombrados = 0
    sin_clave = 0
    errores = 0
    duplicados = 0
    movidos_a_copias = 0
    
    # Diccionario para rastrear archivos ya procesados por clave
    archivos_por_clave = {}
    
    for xml_file in xml_files:
        try:
            # Extraer la clave del XML
            clave = extraer_clave_xml(xml_file)
            
            if not clave:
                sin_clave += 1
                print(f"⚠️  Sin clave: {xml_file.name}", flush=True)
                continue            
            
            # Sanitizar el nombre
            nuevo_nombre = sanitizar_nombre_archivo(clave)
            
            # Agregar extensión .xml (en minúsculas)
            if not nuevo_nombre.lower().endswith('.xml'):
                nuevo_nombre += '.xml'
            
            # Si el nombre ya es correcto, omitir
            if xml_file.name.lower() == nuevo_nombre.lower():
                print(f"✓ Ya tiene nombre correcto: {xml_file.name}", flush=True)
                # Registrar este archivo como el original
                clave_dir = (xml_file.parent, clave)
                if clave_dir not in archivos_por_clave:
                    archivos_por_clave[clave_dir] = xml_file
                continue
            
            # Construir la nueva ruta
            nueva_ruta = xml_file.parent / nuevo_nombre
            clave_dir = (xml_file.parent, clave)
            
            # Verificar si ya existe un archivo con ese nombre en la misma carpeta
            if nueva_ruta.exists() and nueva_ruta != xml_file:
                duplicados += 1
                print(f"🔄 Duplicado detectado: {xml_file.name} -> {nuevo_nombre}", flush=True)
                
                # Crear carpeta Copias en el directorio actual
                carpeta_copias = xml_file.parent / "Copias"
                
                if not dry_run:
                    carpeta_copias.mkdir(exist_ok=True)
                
                # Mover el archivo duplicado a Copias con su clave como nombre
                ruta_copia = carpeta_copias / nuevo_nombre
                
                # Si ya existe en Copias, agregar sufijo
                contador = 1
                while ruta_copia.exists():
                    base_nombre = nuevo_nombre.rsplit('.', 1)[0]
                    ruta_copia = carpeta_copias / f"{base_nombre}_copia_{contador:03d}.xml"
                    contador += 1
                
                if dry_run:
                    print(f"   📦 Movería a: Copias/{ruta_copia.name}", flush=True)
                else:
                    import shutil
                    shutil.move(str(xml_file), str(ruta_copia))
                    movidos_a_copias += 1
                    print(f"   📦 Movido a: Copias/{ruta_copia.name}", flush=True)
                
                continue
            
            # Verificar si ya procesamos un archivo con esta clave en este directorio
            if clave_dir in archivos_por_clave:
                duplicados += 1
                print(f"🔄 Duplicado detectado: {xml_file.name} (ya existe {archivos_por_clave[clave_dir].name})", flush=True)
                
                # Crear carpeta Copias
                carpeta_copias = xml_file.parent / "Copias"
                
                if not dry_run:
                    carpeta_copias.mkdir(exist_ok=True)
                
                # Mover el duplicado a Copias con nombre basado en clave
                ruta_copia = carpeta_copias / nuevo_nombre
                
                # Si ya existe en Copias, agregar sufijo
                contador = 1
                while ruta_copia.exists():
                    base_nombre = nuevo_nombre.rsplit('.', 1)[0]
                    ruta_copia = carpeta_copias / f"{base_nombre}_copia_{contador:03d}.xml"
                    contador += 1
                
                if dry_run:
                    print(f"   📦 Movería a: Copias/{ruta_copia.name}", flush=True)
                else:
                    import shutil
                    shutil.move(str(xml_file), str(ruta_copia))
                    movidos_a_copias += 1
                    print(f"   📦 Movido a: Copias/{ruta_copia.name}", flush=True)
                
                continue
            
            # Renombrar el archivo (primer archivo con esta clave)
            if dry_run:
                print(f"🔄 {xml_file.name} -> {nuevo_nombre}", flush=True)
            else:
                xml_file.rename(nueva_ruta)
                renombrados += 1
                print(f"✅ {xml_file.name} -> {nuevo_nombre}", flush=True)
            
            # Registrar este archivo como el original para esta clave
            archivos_por_clave[clave_dir] = nueva_ruta
        
        except Exception as e:
            errores += 1
            print(f"❌ Error con {xml_file.name}: {e}", flush=True)
    
    # Resumen final
    print()
    print("=" * 60)
    print("📊 RESUMEN")
    print("=" * 60)
    print(f"Archivos procesados: {len(xml_files)}")
    print(f"Renombrados: {renombrados}")
    print(f"Duplicados movidos a Copias/: {movidos_a_copias}")
    print(f"Sin tag <Clave>: {sin_clave}")
    print(f"Errores: {errores}")
    print("=" * 60)

def main():
    """Función principal del script."""
    parser = argparse.ArgumentParser(
        description="Renombrar archivos XML según el valor del tag <Clave>",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos de uso:
  python rename_xml_por_clave.py                    # Usar GUI para seleccionar carpeta
  python rename_xml_por_clave.py --dir "C:\\XMLs"   # Especificar directorio
  python rename_xml_por_clave.py --dir "C:\\XMLs" --dry-run  # Modo prueba
  
El script busca recursivamente en todas las subcarpetas y renombra
cada archivo XML usando el contenido del tag <Clave>.
        """
    )
    
    parser.add_argument(
        '--dir', '--input-dir',
        dest='input_dir',
        default=None,
        help='Directorio con los archivos XML a renombrar'
    )
    
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Modo prueba: muestra qué haría sin renombrar archivos'
    )
    
    args = parser.parse_args()
    
    try:
        print("🏷️  RENOMBRADOR DE XML POR CLAVE")
        print("=" * 60)
        print()
        
        # Seleccionar directorio
        input_dir = args.input_dir
        if not input_dir:
            input_dir = seleccionar_carpeta("Selecciona la carpeta con archivos XML")
            if not input_dir:
                print("❌ No se seleccionó ninguna carpeta. Cancelando.")
                return
        
        print(f"📁 Directorio: {input_dir}")
        print()
        
        # Procesar archivos
        renombrar_xml_por_clave(input_dir, dry_run=args.dry_run)
        
        print()
        print("✅ Proceso completado.")
        
    except KeyboardInterrupt:
        print("\n⏹️  Operación cancelada por el usuario.")
    except Exception as e:
        print(f"\n❌ Error inesperado: {e}")

if __name__ == "__main__":
    main()
