## 🎮 Uso Rápido

### Método Más Fácil (GUI)
```bash
python src/extractor_xml_pst_gui.py
```
- Se abre un diálogo para seleccionar tu archivo PST
- Barra de progreso visual durante el procesamiento  
- Notificaciones de éxito con resultados

### Especificar Archivo PST
```bash
python src/extractor_xml_pst_gui.py -i "mi_archivo.pst"
```

### Para Archivos EML
```bash
python src/extractor_xml_eml.py -i "directorio_eml" -o "salida"
```

## 📋 Requisitos del Sistema

- **Python 3.8+**
- **Windows** (requerido para acceso a PST)
- **Microsoft Outlook** instalado
- **pywin32** (`pip install pywin32`)

## 🚀 Instalación

### 1. Clonar el repositorio
```bash
git clone https://github.com/PENDRAGON2121/pstextractor.git
cd pstextractor
```

### 2. Crear entorno virtual
```bash
python -m venv venv
venv\Scripts\activate  # Windows
```

### 3. Instalar dependencias
```bash
pip install -r requirements.txt
```

### 4. Ejecutar
```bash
python src/extractor_xml_pst_gui.py
```

## 📁 Estructura del Proyecto

```
pstextractor/
├── src/
│   ├── extractor_xml_pst_gui.py    # ⭐ Extractor principal con GUI
│   ├── guia_pst_interactiva.py     # Guía paso a paso
│   ├── extractor_xml_eml.py        # Extractor para archivos EML
│   ├── config.py                   # Configuraciones compartidas
│   └── README.md                   # Documentación detallada
├── requirements.txt                # Dependencias Python
├── GUIA_USO_EXTRACTOR_GUI.md      # Guía completa de uso
└── README.md                       # Este archivo
```

## 🎯 Casos de Uso

- **Contadores/Administradores**: Extraer facturas electrónicas de correos masivos
- **Empresas**: Procesar archivos PST con miles de correos
- **Automatización**: Integrar en procesos de facturación automatizados
- **Auditorías**: Recopilar facturas para revisiones contables

## 📊 Ejemplo de Resultados

Después de ejecutar el extractor:

```
📊 RESULTADOS FINALES
============================
📧 Emails procesados: 1,247
📄 XMLs extraídos: 23
❌ Errores: 0
📁 XMLs guardados en: output/xml_facturacion/

Archivos encontrados:
✅ FE-50629092500310147188400100001010000003257116489476.xml
✅ FE-50629092500310164800900100001010000025110147164634.xml
...
```

## 🛠️ Herramientas Incluidas

| Herramienta | Descripción | Uso |
|------------|-------------|-----|
| `extractor_xml_pst_gui.py` | Extractor principal con GUI | Uso diario recomendado |
| `guia_pst_interactiva.py` | Guía interactiva paso a paso | Cuando GUI no funciona |
| `extractor_xml_eml.py` | Procesador de archivos EML | Para correos individuales |

## ❓ Preguntas Frecuentes

**¿Funciona sin Outlook instalado?**
No, requiere Microsoft Outlook para acceder a archivos PST de manera confiable.

**¿Qué tipos de XML encuentra?**
Busca y extrae cualquier archivo con extensión `.xml` adjunto en los correos del PST, sin importar el nombre.

**¿Es seguro?**
Sí, solo lee los archivos PST localmente, no envía datos a ningún servidor.

**¿Funciona con PST grandes?**
Sí, maneja archivos PST de varios GB con barra de progreso visual.

## 🤝 Contribuir

1. Fork el proyecto
2. Crea una rama para tu feature (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

## 📜 Licencia

Este proyecto está bajo la Licencia MIT - ver el archivo [LICENSE](LICENSE) para detalles.

## 👤 Autor

**PENDRAGON2121**

- GitHub: [@PENDRAGON2121](https://github.com/PENDRAGON2121)

## 🙏 Agradecimientos

- Microsoft Outlook COM API
- Comunidad Python
- Usuarios beta que probaron la herramienta