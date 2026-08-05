#!/usr/bin/env bash
# render-build.sh — Script de compilación para Render.
# Instala LibreOffice headless (conversión XLSX -> PDF) y las dependencias de Python.
set -o errexit

echo ">>> [render-build] Iniciando build..."

# El build de Render corre como root: no se usa sudo.
echo ">>> [render-build] Instalando LibreOffice..."
apt-get update -y
apt-get install -y libreoffice fonts-liberation

echo ">>> [render-build] Verificando LibreOffice:"
soffice --version || libreoffice --version || { echo "ERROR: LibreOffice no quedo instalado."; exit 1; }

echo ">>> [render-build] Instalando dependencias de Python..."
pip install --upgrade pip
pip install -r requirements.txt

echo ">>> [render-build] Build finalizado correctamente."
