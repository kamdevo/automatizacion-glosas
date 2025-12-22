#!/bin/bash
echo "========================================="
echo " Descargador de GLOSAS - HUV"
echo " Verificando dependencias..."
echo "========================================="
echo ""

# Verificar si python3 está instalado
if ! command -v python3 &> /dev/null; then
    echo "❌ ERROR: Python 3 no está instalado"
    echo "Instálalo con: sudo apt install python3"
    exit 1
fi

# Verificar si pip3 está instalado
if ! command -v pip3 &> /dev/null; then
    echo "⚠️  pip3 no está instalado"
    read -p "¿Deseas instalarlo ahora? (requiere sudo) [S/n]: " respuesta
    
    if [[ "$respuesta" =~ ^[Ss]$ ]] || [[ -z "$respuesta" ]]; then
        sudo apt update
        sudo apt install -y python3-pip
    else
        echo "❌ No se puede continuar sin pip3"
        exit 1
    fi
fi

# Verificar si python3-tk está instalado
if ! python3 -c "import tkinter" 2>/dev/null; then
    echo "⚠️  python3-tk no está instalado"
    read -p "¿Deseas instalarlo ahora? (requiere sudo) [S/n]: " respuesta
    
    if [[ "$respuesta" =~ ^[Ss]$ ]] || [[ -z "$respuesta" ]]; then
        sudo apt update
        sudo apt install -y python3-tk
    else
        echo "❌ No se puede continuar sin python3-tk"
        exit 1
    fi
fi

# Verificar e instalar dependencias de Python
echo ""
echo "📦 Verificando dependencias de Python..."
pip3 install --user --break-system-packages -q google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client 2>/dev/null

if [ $? -ne 0 ]; then
    echo "⚠️  Instalando dependencias con permisos de usuario..."
    pip3 install --user google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client
    if [ $? -ne 0 ]; then
        echo "❌ Error al instalar dependencias de Python"
        exit 1
    fi
fi

echo ""
echo "✅ Todas las dependencias están instaladas"
echo "🚀 Iniciando Descargador de GLOSAS..."
echo ""

# Obtener la ruta del script
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# Ejecutar el script Python directamente
python3 "$DIR/glosas_file_downloader.py"
