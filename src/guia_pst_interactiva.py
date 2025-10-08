#!/usr/bin/env python3
"""
Guía Interactiva PST - Extractor XML Facturación
Herramienta paso a paso para extraer XMLs de facturación de archivos PST
"""

import os
import sys
import argparse
from pathlib import Path
import time

def mostrar_banner():
    print("""
╔═══════════════════════════════════════════════════════════════╗
║                   🧾 EXTRACTOR XML PST 🧾                    ║
║               Guía paso a paso para extraer                   ║
║              facturas electrónicas de PST                    ║
╚═══════════════════════════════════════════════════════════════╝
""")

def verificar_pst(pst_path):
    """Verifica el archivo PST y muestra información"""
    pst = Path(pst_path)
    
    if not pst.exists():
        print(f"❌ ERROR: Archivo PST no encontrado en: {pst_path}")
        return False
    
    size_mb = pst.stat().st_size / (1024 * 1024)
    print(f"""
📊 INFORMACIÓN DEL ARCHIVO PST:
   📂 Ubicación: {pst_path}
   📏 Tamaño: {size_mb:.1f} MB
   ✅ Archivo válido y accesible
""")
    
    return True

def crear_directorio_trabajo(output_dir):
    """Crea la estructura de directorios"""
    base_dir = Path(output_dir)
    
    dirs = {
        'base': base_dir,
        'xml_facturacion': base_dir / 'xml_facturacion',
        'temp': base_dir / 'temp',
        'reportes': base_dir / 'reportes'
    }
    
    print("📁 Creando estructura de directorios...")
    for name, path in dirs.items():
        path.mkdir(parents=True, exist_ok=True)
        print(f"   ✅ {name}: {path}")
    
    return dirs

def mostrar_opciones_extraccion():
    """Muestra las opciones disponibles para extraer el PST"""
    print("""
🔧 OPCIONES DE EXTRACCIÓN DISPONIBLES:

1️⃣  OPCIÓN AUTOMÁTICA - Outlook (Recomendada si tienes Outlook)
    • Usar Microsoft Outlook para importar el PST
    • Buscar automáticamente adjuntos XML
    • Más fácil y rápida

2️⃣  OPCIÓN MANUAL - Búsqueda directa
    • Abrir PST manualmente en Outlook
    • Buscar correos con adjuntos
    • Guardar XMLs manualmente

3️⃣  OPCIÓN HERRAMIENTAS - Software especializado
    • PST Walker (GUI gratuita)
    • readpst (línea de comandos)
    • Herramientas online
""")

def guia_outlook_automatica(pst_path, output_dir):
    """Guía para usar Outlook automáticamente"""
    print("""
🚀 OPCIÓN 1: EXTRACCIÓN AUTOMÁTICA CON OUTLOOK

Pasos a seguir:
""")
    
    print("1️⃣  Abriendo Microsoft Outlook...")
    print("   (Si no se abre automáticamente, ábrelo manualmente)")
    
    # Intentar abrir Outlook
    try:
        import win32com.client
        print("   ✅ Outlook COM disponible")
        
        print("\n2️⃣  Configurando acceso al PST...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        print(f"\n3️⃣  Intentando abrir PST: {pst_path}")
        print("   ⏳ Esto puede tomar unos momentos...")
        
        # Intentar añadir el PST
        try:
            namespace = outlook.GetNamespace("MAPI")
            namespace.AddStore(pst_path)
            print("   ✅ PST añadido correctamente a Outlook")
            
            # Buscar la nueva store
            stores = namespace.Stores
            pst_store = None
            
            for store in stores:
                if pst_path.lower() in store.FilePath.lower():
                    pst_store = store
                    break
            
            if pst_store:
                print(f"   ✅ PST encontrado: {pst_store.DisplayName}")
                return buscar_xmls_en_outlook(pst_store, output_dir)
            else:
                print("   ⚠️  PST añadido pero no encontrado en stores")
                return False
                
        except Exception as e:
            print(f"   ❌ Error al añadir PST: {e}")
            return False
            
    except Exception as e:
        print(f"   ❌ Error iniciando Outlook: {e}")
        return False

def buscar_xmls_en_outlook(pst_store, output_dir):
    """Busca XMLs de facturación en el PST usando Outlook"""
    import re
    
    xml_pattern = re.compile(r"^FE[-_]?\d{3,}\.xml$", re.IGNORECASE)
    xml_encontrados = []
    
    print("\n4️⃣  Buscando correos con adjuntos XML...")
    
    try:
        # Obtener carpeta raíz
        root_folder = pst_store.GetRootFolder()
        
        def procesar_carpeta(folder, level=0):
            indent = "   " * level
            print(f"{indent}📁 Procesando: {folder.Name}")
            
            try:
                # Procesar elementos en esta carpeta
                items = folder.Items
                count = 0
                
                for item in items:
                    try:
                        if hasattr(item, 'Attachments') and item.Attachments.Count > 0:
                            for attachment in item.Attachments:
                                if xml_pattern.match(attachment.FileName):
                                    count += 1
                                    xml_path = Path(output_dir) / 'xml_facturacion' / attachment.FileName
                                    
                                    # Guardar adjunto
                                    attachment.SaveAsFile(str(xml_path))
                                    xml_encontrados.append(str(xml_path))
                                    
                                    print(f"{indent}   ✅ XML guardado: {attachment.FileName}")
                    except:
                        continue
                
                if count > 0:
                    print(f"{indent}   📊 {count} XMLs encontrados en esta carpeta")
                
                # Procesar subcarpetas
                try:
                    for subfolder in folder.Folders:
                        procesar_carpeta(subfolder, level + 1)
                except:
                    pass
                    
            except Exception as e:
                print(f"{indent}   ⚠️  Error procesando carpeta: {e}")
        
        # Procesar todas las carpetas
        procesar_carpeta(root_folder)
        
        return xml_encontrados
        
    except Exception as e:
        print(f"   ❌ Error buscando XMLs: {e}")
        return []

def guia_outlook_manual(pst_path, output_dir):
    """Guía paso a paso para extracción manual"""
    print(f"""
👤 OPCIÓN 2: EXTRACCIÓN MANUAL

Sigue estos pasos:

1️⃣  Abrir Microsoft Outlook
    • Abre Outlook desde el menú inicio
    • Espera a que cargue completamente

2️⃣  Importar el archivo PST
    • Ve a: Archivo > Abrir y exportar > Abrir archivo de datos de Outlook
    • Selecciona: {pst_path}
    • Haz clic en "Aceptar"

3️⃣  Navegar por el PST importado
    • En el panel izquierdo verás una nueva carpeta con el nombre del PST
    • Expande todas las carpetas (Bandeja de entrada, Enviados, etc.)

4️⃣  Buscar correos con adjuntos
    • Busca el ícono 📎 que indica correos con adjuntos
    • O usa la búsqueda: hasattachments:yes

5️⃣  Identificar XMLs de facturación
    • Abre correos con adjuntos
    • Busca archivos con nombres como:
      - FE001.xml, FE1234.xml, FE-12345.xml, FE_67890.xml
      - Cualquier XML que empiece con "FE" seguido de números

6️⃣  Guardar los XMLs
    • Haz clic derecho en el adjunto XML
    • Selecciona "Guardar como..."
    • Guarda en: {output_dir}\\xml_facturacion\\

7️⃣  Repetir para todos los correos
    • Continúa buscando en todas las carpetas
    • Guarda todos los XMLs de facturación que encuentres

🎯 OBJETIVO: Encontrar archivos XML de facturas electrónicas
📁 DESTINO: {output_dir}\\xml_facturacion\\
""")
    
    input("\n⏸️  Presiona ENTER cuando hayas terminado de extraer los XMLs...")
    
    # Verificar qué se extrajo
    xml_dir = Path(output_dir) / 'xml_facturacion'
    xmls_encontrados = list(xml_dir.glob('*.xml'))
    
    return [str(xml) for xml in xmls_encontrados]

def guia_herramientas_externas():
    """Información sobre herramientas externas"""
    print("""
🛠️  OPCIÓN 3: HERRAMIENTAS ESPECIALIZADAS

A. PST Walker (Recomendada - GUI gratuita)
   • Descarga: https://www.pstwalker.com/
   • Instalar y abrir
   • Abrir tu archivo PST
   • Navegar por carpetas y exportar adjuntos

B. readpst (Línea de comandos)
   • Descargar libpst desde: https://www.five-ten-sg.com/libpst/
   • O instalar chocolatey: https://chocolatey.org/
   • Comando: readpst -r -o salida archivo.pst

C. Herramientas Online (para PSTs pequeños)
   • PST Viewer Online
   • SysTools PST Viewer
   • Kernel PST Viewer

D. Software Comercial
   • Stellar PST Repair
   • SysTools PST Viewer Pro
   • Recovery Toolbox for Outlook
""")

def generar_reporte(xmls_encontrados, output_dir):
    """Genera un reporte de los XMLs encontrados"""
    import csv
    from datetime import datetime
    
    if not xmls_encontrados:
        print("\n📊 RESULTADO: No se encontraron archivos XML de facturación")
        return
    
    print(f"\n🎉 ¡ÉXITO! Se encontraron {len(xmls_encontrados)} archivos XML de facturación:")
    
    # Crear reporte CSV
    reporte_path = Path(output_dir) / 'reportes' / 'xmls_extraidos.csv'
    
    with open(reporte_path, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['Archivo', 'Ruta Completa', 'Tamaño (bytes)', 'Fecha Extracción'])
        
        for xml_path in xmls_encontrados:
            xml_file = Path(xml_path)
            if xml_file.exists():
                size = xml_file.stat().st_size
                filename = xml_file.name
                writer.writerow([filename, xml_path, size, datetime.now().isoformat()])
                print(f"   ✅ {filename} ({size} bytes)")
    
    print(f"\n📋 Reporte guardado en: {reporte_path}")
    
    # Crear archivo de resumen
    resumen_path = Path(output_dir) / 'RESUMEN_EXTRACCION.txt'
    with open(resumen_path, 'w', encoding='utf-8') as f:
        f.write(f"RESUMEN DE EXTRACCIÓN XML\n")
        f.write(f"========================\n\n")
        f.write(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"XMLs encontrados: {len(xmls_encontrados)}\n\n")
        f.write("Archivos extraídos:\n")
        for xml_path in xmls_encontrados:
            f.write(f"- {Path(xml_path).name}\n")
    
    print(f"📝 Resumen guardado en: {resumen_path}")

def main():
    parser = argparse.ArgumentParser(description='Guía Interactiva PST - Extractor XML')
    parser.add_argument('-i', '--input', required=True, help='Archivo PST de entrada')
    parser.add_argument('-o', '--output', required=True, help='Directorio de salida')
    parser.add_argument('--modo', choices=['auto', 'manual', 'herramientas'], 
                       default='auto', help='Modo de extracción')
    
    args = parser.parse_args()
    
    # Mostrar banner
    mostrar_banner()
    
    # Verificar PST
    if not verificar_pst(args.input):
        return 1
    
    # Crear directorios
    dirs = crear_directorio_trabajo(args.output)
    
    # Mostrar opciones
    if args.modo == 'auto':
        mostrar_opciones_extraccion()
        print("🚀 Iniciando extracción automática...")
        xmls_encontrados = guia_outlook_automatica(args.input, args.output)
        
        if not xmls_encontrados:
            print("\n⚠️  Extracción automática falló. Cambiando a modo manual...")
            xmls_encontrados = guia_outlook_manual(args.input, args.output)
            
    elif args.modo == 'manual':
        xmls_encontrados = guia_outlook_manual(args.input, args.output)
        
    elif args.modo == 'herramientas':
        guia_herramientas_externas()
        return 0
    
    # Generar reporte
    generar_reporte(xmls_encontrados, args.output)
    
    print(f"""
🏁 EXTRACCIÓN COMPLETADA

📂 Directorio de trabajo: {args.output}
📧 XMLs de facturación: {args.output}\\xml_facturacion\\
📊 Reportes: {args.output}\\reportes\\

¡Revisa los archivos extraídos!
""")
    
    return 0

if __name__ == "__main__":
    sys.exit(main())