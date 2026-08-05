#!/usr/bin/env bash
# render-build.sh — Script de compilación para Render.
# Instala LibreOffice headless (conversión XLSX -> PDF) y las dependencias de Python.
set -o errexit

echo ">>> [render-build] Iniciando build..."

# Render puede ejecutar el build como root o como usuario con sudo.
if [ "$(id -u)" -eq 0 ]; then
    SUDO=""
else
    SUDO="sudo"
fi

echo ">>> [render-build] Instalando LibreOffice headless..."
$SUDO apt-get update -y
$SUDO apt-get install -y --no-install-recommends \
    libreoffice-core \
    libreoffice-calc \
    libreoffice-common \
    fonts-liberation \
    fonts-dejavu

echo ">>> [render-build] Verificando LibreOffice:"
soffice --version || libreoffice --version || { echo "ERROR: LibreOffice no quedo instalado."; exit 1; }

echo ">>> [render-build] Instalando dependencias de Python..."
pip install --upgrade pip
pip install -r requirements.txt

echo ">>> [render-build] Build finalizado correctamente."
