# � Extractores de XML

Scripts Python para extraer archivos XML adjuntos (cualquier nombre con extensión `.xml`) desde correos electrónicos EML y archivos PST de Outlook.

## 📁 Archivos Principales

### 🎯 Scripts Funcionales
- **`extractor_xml_pst_gui.py`** - **★ PRINCIPAL:** Extractor PST con interfaz gráfica completa
- **`guia_pst_interactiva.py`** - Guía paso a paso para extracción manual de PST  
- **`extractor_xml_eml.py`** - Extractor para archivos EML individuales o directorios
- **`config.py`** - Configuraciones compartidas y funciones utilitarias

### 📓 Notebook Jupyter
- **`extractor_xml_facturacion.ipynb`** - Notebook interactivo con procesamiento EML

## 🚀 Instalación Rápida

1. **Instalar dependencias:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Para archivos PST (opcional):**
   ```bash
   # En Windows (puede requerir compilación)
   pip install pypff
   
   # En Linux
   sudo apt-get install libpff-dev python3-pypff
   ```

## � Uso de los Scripts

### 🎯 Extractor PST con GUI (Recomendado)

```bash
# Método más fácil - Selección automática de archivo
python src/extractor_xml_pst_gui.py

# Especificar archivo PST
python src/extractor_xml_pst_gui.py -i "archivo.pst"

```

**Características:**
- �️ Interfaz gráfica para seleccionar archivos
- 📊 Barra de progreso visual en tiempo real
- ✅ Notificaciones de éxito/error
- 📁 Creación automática de directorios
- 📋 Logs detallados y reportes

### �📧 Para Archivos EML

```bash
# Procesar directorio con archivos EML
python src/extractor_xml_eml.py -i "C:/correos/eml_files" -o "C:/xml_extracts"

# Procesar sin validación (más rápido)
python src/extractor_xml_eml.py -i "C:/correos/eml_files" -o "C:/xml_extracts" --no-validate
```

### �️ Guía Interactiva PST

```bash
# Para casos complejos o cuando la GUI no funciona
python src/guia_pst_interactiva.py -i "archivo.pst" -o "salida"
python src/extractor_xml_pst.py -i "C:/correos/archivo.pst" -o "C:/xml_extracts" --no-validate
```

### 🔄 Script Unificado (Recomendado)

```bash
# Detecta automáticamente el tipo de archivo
python src/extractor_xml_unificado.py -i "C:/correos/archivo.pst" -o "C:/xml_extracts"

# Procesar directorio con archivos EML
python src/extractor_xml_unificado.py -i "C:/correos/eml_files" -o "C:/xml_extracts"

# Búsqueda recursiva en subdirectorios
python src/extractor_xml_unificado.py -i "C:/correos" -o "C:/xml_extracts" --recursive
```

## 📊 Archivos Generados

Después de la ejecución, encontrarás:

```
xml_extracts/
├── factura1.xml           # Archivos XML extraídos (cualquier nombre)
├── FE-000124.xml
├── otro_documento.XML
├── remitentes_eml.csv     # Log de archivos EML (si aplica)
├── remitentes_pst.csv     # Log de archivos PST (si aplica)
└── remitentes_todos.csv   # Log unificado
```

### 📋 Formato del CSV de Log

```csv
archivo_xml,remitente,asunto,fecha_email,fecha_procesamiento,origen,ubicacion_origen
FE000123.xml,proveedor1@empresa.com,Factura 123,2025-10-07 10:30:00,2025-10-07 15:45:00,EML,archivo001.eml
FE000124.xml,proveedor2@empresa.com,Factura 124,2025-10-07 11:00:00,2025-10-07 15:45:01,PST,Inbox/Facturas
```

## 🎯 Características

### ✅ **Funcionalidades Principales**
- ✅ Extrae todos los archivos con extensión `.xml` (independientemente del nombre)
- ✅ Procesamiento con barras de progreso (tqdm)
- ✅ Validación automática de archivos XML extraídos
- ✅ Log detallado de remitentes y fechas
- ✅ Manejo de nombres duplicados (sufijos automáticos)
- ✅ Soporte para grandes volúmenes (30,000+ archivos)

### 🔧 **Configuraciones Avanzadas**
- ✅ Patrones XML personalizables (FE, NC, ND, DS)
- ✅ Límites de tamaño de archivo configurables
- ✅ Limpieza automática de campos CSV
- ✅ Manejo robusto de errores

### 🛡️ **Validaciones**
- ✅ Verificación de archivos XML bien formados
- ✅ Detección automática de tipos de archivo
- ✅ Validación de rutas y permisos
- ✅ Manejo de archivos corruptos

## ⚙️ Configuración Avanzada

### 📝 Personalizar Patrones XML

Edita `config.py` para agregar nuevos patrones:

```python
ADDITIONAL_PATTERNS = {
    "NC": re.compile(r"^NC[-_]?\d{3,}\.xml$", re.IGNORECASE),  # Notas de crédito
    "ND": re.compile(r"^ND[-_]?\d{3,}\.xml$", re.IGNORECASE),  # Notas de débito
    "DS": re.compile(r"^DS[-_]?\d{3,}\.xml$", re.IGNORECASE),  # Documentos soporte
}
```

### 🎛️ Ajustar Límites

```python
MAX_XML_SIZE_MB = 10          # Tamaño máximo por archivo XML
MAX_LOG_FIELD_LENGTH = 100    # Longitud máxima de campos en CSV
```

## 🐛 Solución de Problemas

### ❌ Error: "pypff could not be resolved"

Para archivos PST, necesitas instalar `pypff`:

```bash
# Windows (puede requerir Visual Studio Build Tools)
pip install pypff

# Linux
sudo apt-get install libpff-dev python3-pypff

# macOS
brew install libpff
pip install pypff
```

### ❌ Error: "No se encontraron archivos EML"

Verifica que:
- La ruta sea correcta
- Los archivos tengan extensión `.eml` o `.msg`
- Tengas permisos de lectura en el directorio

### ❌ Error: "Memory error" con archivos PST grandes

Para archivos PST muy grandes:
- Usa `--no-validate` para acelerar el proceso
- Asegúrate de tener suficiente RAM disponible
- Considera procesar por partes

## 📈 Rendimiento

### 📊 Tiempos Estimados
- **Archivos EML**: ~1,000 archivos/minuto
- **Archivos PST**: Depende del tamaño (puede tomar horas para PST de GB)
- **Validación XML**: ~5,000 archivos/minuto

### 💡 Consejos de Optimización
- Usa `--no-validate` para procesamiento inicial rápido
- Procesa archivos PST en horarios de baja actividad
- Asegúrate de tener suficiente espacio en disco (20-30% del tamaño original)

## 🆘 Soporte

Si encuentras problemas:

1. **Revisa los logs** de errores mostrados en pantalla
2. **Verifica las dependencias** con `pip list`
3. **Comprueba los permisos** de archivos y directorios
4. **Consulta la documentación** de pypff para problemas con PST

---

## 🧪 Ejemplos de Prueba

### 📝 Probar Configuración

```bash
# Probar configuraciones
python src/config.py

# Probar extractor EML con archivo de ejemplo
python src/extractor_xml_eml.py -i "ejemplos/eml" -o "prueba_xml"

# Probar script unificado
python src/extractor_xml_unificado.py -i "ejemplos" -o "prueba_xml" --recursive
```

¡Los scripts están listos para procesar tus archivos de facturación! 🚀