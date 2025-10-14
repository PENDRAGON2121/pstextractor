import os
import errno
from pathlib import Path
import shutil
import argparse
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox

def obtener_tag_raiz(xml_path: Path) -> str:
    """Obtener el nombre del tag raíz del XML (sin namespace)."""
    try:
        for _event, elem in ET.iterparse(xml_path, events=("start",)):
            tag = elem.tag
            if "}" in tag:
                tag = tag.split("}", 1)[1]
            return tag
    except ET.ParseError as e:
        raise e
    except OSError as e:
        raise e
    return ""

def tiene_tag_raiz(xml_path: Path, tag_buscado: str) -> bool:
    try:
        tag = obtener_tag_raiz(xml_path)
        return tag == tag_buscado
    except ET.ParseError as e:
        print(f"XML inválido {xml_path}: {e}", flush=True)
        return False
    except OSError as e:
        print(f"No se pudo leer {xml_path}: {e}", flush=True)
        return False

def obtener_destino_unico(destino_dir: Path, nombre_archivo: str) -> Path:
    """Generar un nombre único en el directorio destino evitando sobrescrituras."""
    destino_dir.mkdir(parents=True, exist_ok=True)
    destino = destino_dir / nombre_archivo
    if not destino.exists():
        return destino

    base, ext = os.path.splitext(nombre_archivo)
    contador = 1
    while True:
        candidato = destino_dir / f"{base}_{contador:03d}{ext}"
        if not candidato.exists():
            return candidato
        contador += 1

def mover_archivo(xml_file: Path, destino_dir: Path) -> bool:
    """Mover un archivo manejando bloqueos; retorna True si se movió."""
    destino_dir.mkdir(parents=True, exist_ok=True)

    def copiar_y_eliminar() -> bool:
        destino = obtener_destino_unico(destino_dir, xml_file.name)
        try:
            shutil.copy2(str(xml_file), str(destino))
            try:
                os.remove(xml_file)
            except PermissionError as delete_err:
                print(f"No se pudo eliminar el original {xml_file} tras copiarlo ({delete_err}).")
                return False
            return True
        except Exception as copy_err:
            print(f"No se pudo mover {xml_file} (copia fallida). Error: {copy_err}")
            return False

    destino = obtener_destino_unico(destino_dir, xml_file.name)
    try:
        shutil.move(str(xml_file), str(destino))
        return True
    except PermissionError as err:
        print(f"Archivo en uso {xml_file} ({err}). Intentando copiar y eliminar...")
        return copiar_y_eliminar()
    except OSError as err:
        win_err = getattr(err, "winerror", None)
        if err.errno in (errno.EACCES, errno.EPERM) or win_err in (32,):
            print(f"No se pudo mover {xml_file} directamente ({err}). Intentando copiar...")
            return copiar_y_eliminar()
        print(f"No se pudo mover {xml_file}. Error: {err}")
        return False

def esta_en_directorio(path: Path, directorio: Path) -> bool:
    """Verificar si path está dentro del directorio proporcionado."""
    try:
        path.resolve().relative_to(directorio.resolve())
        return True
    except ValueError:
        return False

def seleccionar_carpeta(titulo):
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

def procesar_xmls(input_dir, output_dir=None):
    base_dir = Path(input_dir)
    base_dir.mkdir(exist_ok=True)

    output_base = Path(output_dir).resolve() if output_dir else None

    xml_files = [xml for xml in base_dir.rglob("*.xml")]

    if not xml_files:
        print("No se encontraron archivos XML en la carpeta de entrada.")
        return

    movidos = 0
    procesados = 0
    for xml_file in xml_files:
        procesados += 1
        try:
            # Evitar reprocesar archivos ya movidos
            if output_base:
                try:
                    xml_file.resolve().relative_to(output_base)
                    continue
                except ValueError:
                    pass
            else:
                if any(part.lower() == "haciendaresponse" for part in xml_file.parts):
                    continue

            if tiene_tag_raiz(xml_file, "MensajeHacienda"):
                # Determinar carpeta destino manteniendo estructura
                if output_base:
                    try:
                        relative_parent = xml_file.parent.relative_to(base_dir)
                    except ValueError:
                        relative_parent = Path()
                    destino_dir = output_base / relative_parent / "HaciendaResponse"
                else:
                    destino_dir = xml_file.parent / "HaciendaResponse"

                if mover_archivo(xml_file, destino_dir):
                    movidos += 1
                    print(f"Movido: {xml_file} -> {destino_dir}", flush=True)
        except ValueError as e:
            print(f"Ruta fuera del directorio base, se omite {xml_file}: {e}", flush=True)
        except Exception as e:
            print(f"Error procesando {xml_file}: {e}", flush=True)

        if procesados % 100 == 0:
            print(f"Procesados {procesados} archivos...", flush=True)

    print(f"Procesamiento terminado. Archivos procesados: {procesados}, movidos: {movidos}", flush=True)

    # Listar facturas restantes
    print(f"Archivos en {base_dir} (que empiezan con <FacturaElectronica):")
    for xml_file in base_dir.rglob("*.xml"):
        try:
            if tiene_tag_raiz(xml_file, "FacturaElectronica"):
                print(f"  - {xml_file.relative_to(base_dir)}")
        except Exception as e:
            print(f"  - No se pudo leer {xml_file}: {e}")

    # Listar mensajes movidos
    if output_base:
        print(f"Archivos en {output_base} (que empiezan con <MensajeHacienda):")
        for xml_file in output_base.rglob("*.xml"):
            try:
                if tiene_tag_raiz(xml_file, "MensajeHacienda"):
                    print(f"  - {xml_file.relative_to(output_base)}")
            except Exception:
                pass
    else:
        print("Archivos movidos a subcarpetas HaciendaResponse (por carpeta original):")
        for hacienda_dir in base_dir.rglob("HaciendaResponse"):
            for xml_file in hacienda_dir.glob("*.xml"):
                print(f"  - {xml_file.relative_to(base_dir)}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Filtra y mueve XMLs de Hacienda")
    parser.add_argument('--input-dir', default=None, help='Carpeta de entrada de XMLs')
    parser.add_argument('--output-dir', default=None, help='Carpeta destino para MensajeHacienda')
    args = parser.parse_args()

    input_dir = args.input_dir or seleccionar_carpeta("Selecciona la carpeta de entrada de XMLs")
    if not input_dir:
        print("No se seleccionó carpeta de entrada. Cancelando.")
        exit(1)
    if args.output_dir is not None:
        output_dir = args.output_dir or seleccionar_carpeta("Selecciona la carpeta de salida para HaciendaResponse")
        if not output_dir:
            print("No se seleccionó carpeta de salida. Cancelando.")
            exit(1)
    else:
        output_dir = None

    procesar_xmls(input_dir, output_dir)
