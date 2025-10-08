# 🧾 EXTRACTOR XML PST CON GUI - GUÍA DE USO

## ✅ COMPLETADO CON ÉXITO

Tu archivo PST ha sido procesado exitosamente usando el **Extractor XML PST con GUI**.

### 📊 RESULTADOS DE LA EXTRACCIÓN:

- **📧 Emails procesados:** 106 correos
- **📄 XMLs encontrados:** 2 archivos de facturación electrónica
- **❌ Errores:** 0 (procesamiento perfecto)
- **📁 Ubicación:** `C:\Users\mquir\Desktop\mquiros@parquetempisque.com_xml_extraidos`

### 📂 ARCHIVOS EXTRAÍDOS:

**XMLs de Facturación Electrónica:**
1. `FE-50629092500310147188400100001010000003257116489476.xml` (10,132 bytes)
2. `FE-50629092500310164800900100001010000025110147164634.xml` (10,556 bytes)

**Archivos de Soporte:**
- `remitentes_pst.csv` - Log detallado con información de remitentes y fechas
- `reportes/reporte_extraccion.txt` - Resumen completo del proceso

---

## 🚀 CÓMO USAR EL EXTRACTOR XML PST CON GUI

### Método 1: Selección Automática (Recomendado)
```bash
python src/extractor_xml_pst_gui.py
```
- Se abrirá un diálogo para seleccionar el archivo PST
- El directorio de salida se creará automáticamente
- Interfaz gráfica con barra de progreso
- Notificaciones visuales de éxito/error

### Método 2: Especificar Archivo PST
```bash
python src/extractor_xml_pst_gui.py -i "ruta/al/archivo.pst"
```
- Especifica el archivo PST desde línea de comandos
- El directorio de salida se genera automáticamente
- Mantiene la interfaz gráfica para el progreso

### Método 3: Control Completo
```bash
python src/extractor_xml_pst_gui.py -i "archivo.pst" -o "directorio_salida"
```
- Control total sobre entrada y salida
- Perfecto para automatización o scripts

---

## 🎯 CARACTERÍSTICAS PRINCIPALES

### ✨ Interfaz Gráfica Intuitiva
- **Selector de archivos:** Diálogo visual para elegir PST
- **Barra de progreso:** Muestra el avance en tiempo real
- **Estadísticas live:** Emails procesados y XMLs encontrados
- **Notificaciones:** Mensajes de éxito/error visuales

### 🔍 Detección Inteligente
- **Patrón flexible:** Encuentra FE12345.xml, FE-12345.xml, FE_12345.xml
- **Sin duplicados:** Maneja archivos con nombres similares automáticamente
- **Validación:** Verifica que los XMLs sean válidos

### 📋 Reporting Completo
- **Log CSV:** Registro detallado con remitente, asunto, fecha, carpeta
- **Reporte de estadísticas:** Resumen ejecutivo del procesamiento
- **Estructura organizada:** Carpetas separadas para XMLs y reportes

### 🛡️ Robustez y Confiabilidad
- **Múltiples métodos:** Outlook COM como método principal confiable
- **Manejo de errores:** Continúa procesando aunque encuentre problemas
- **Progreso visual:** Ventana que se puede minimizar durante el proceso

---

## 📁 ESTRUCTURA DE SALIDA

```
nombre_pst_xml_extraidos/
├── xml_facturacion/           # 📄 Archivos XML de facturación
│   ├── FE-123456789.xml
│   └── FE-987654321.xml
├── reportes/                  # 📊 Reportes y estadísticas
│   └── reporte_extraccion.txt
└── remitentes_pst.csv        # 📋 Log detallado
```

---

## 🔧 DEPENDENCIAS

### Incluidas con Python:
- `tkinter` - Interfaz gráfica
- `pathlib` - Manejo de rutas
- `re` - Expresiones regulares

### Requieren instalación:
```bash
pip install pywin32  # Para Outlook COM (crítico)
pip install tqdm     # Barras de progreso adicionales
pip install lxml     # Validación XML (opcional)
```

---

## ❓ RESOLUCIÓN DE PROBLEMAS

### Error: "win32com.client no disponible"
**Solución:**
```bash
pip install pywin32
```

### Error: "PST no se pudo cargar"
**Causas posibles:**
1. El PST está corrupto
2. El PST está en uso por Outlook
3. Permisos insuficientes

**Soluciones:**
1. Cierra Outlook completamente
2. Ejecuta como administrador
3. Usa la herramienta de reparación PST de Microsoft

### No se encuentran XMLs
**Verificar:**
1. Los emails contienen adjuntos XML
2. Los XMLs siguen el patrón FE*.xml
3. El PST tiene correos en las carpetas principales

---

## 🎉 PRÓXIMOS PASOS

1. **Revisar XMLs extraídos** en `xml_facturacion/`
2. **Verificar el log** en `remitentes_pst.csv` para detalles
3. **Consultar el reporte** en `reportes/reporte_extraccion.txt`
4. **Procesar otros PSTs** repitiendo el comando

---

## 📞 SOPORTE

Si encuentras problemas:
1. Revisa el archivo de log `remitentes_pst.csv`
2. Consulta el reporte de errores si existe
3. Verifica que todas las dependencias estén instaladas
4. Asegúrate de que Outlook no esté usando el PST

---

**¡El sistema está listo para usar con cualquier archivo PST!** 🎯