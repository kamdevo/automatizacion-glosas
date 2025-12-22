# Descargador de GLOSAS - HUV (Versión Linux)

## 📋 Requisitos del Sistema

- Python 3.8 o superior
- python3-tk
- pip3

## 🚀 Ejecución

### Método Recomendado: Lanzador Automático
El script verificará e instalará automáticamente todas las dependencias:

```bash
./lanzador_linux.sh
```

El script instalará automáticamente:
- pip3 (si no está instalado)
- python3-tk (interfaz gráfica)
- Librerías de Google API

## 📦 Distribución

Para distribuir este programa a otros usuarios de Linux, incluye:

1. **Archivo `glosas_file_downloader.py`** - Script principal
2. **Archivo `credentials.json`** - Credenciales de Google API
3. **Script `lanzador_linux.sh`** - Lanzador con instalación automática
4. **Este archivo README_LINUX.md** - Instrucciones

### Estructura de Distribución:
```
GlosasAutomatizador_Linux/
├── glosas_file_downloader.py
├── credentials.json
├── lanzador_linux.sh
└── README_LINUX.md
```

## ⚠️ Diferencias con Windows

- **Windows:** Ejecutable standalone (.exe) que no requiere Python instalado
- **Linux:** Requiere Python 3 y dependencias (instaladas automáticamente por el lanzador)

## 🔧 Solución de Problemas

### Error: "ModuleNotFoundError: No module named 'tkinter'"

**Solución:**
```bash
sudo apt install python3-tk
```

### Error: "Permission denied"

**Solución:**
```bash
chmod +x lanzador_linux.sh
```

### El programa no inicia

Ejecuta directamente con Python:
```bash
python3 glosas_file_downloader.py
```

## 📝 Notas Importantes

- Este ejecutable **SOLO funciona en Linux**
- Para Windows, usa el ejecutable en `dist/GlosasAutomatizador.exe`
- La primera ejecución abrirá el navegador para autenticación de Gmail
- Los archivos descargados se guardan en el Escritorio

## 🏥 Soporte

Hospital Universitario del Valle  
Innovación y Desarrollo  
innovacionydesarrollo@correohuv.gov.co
