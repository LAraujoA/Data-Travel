# 🏥 Data-Travel

![Python](https://img.shields.io/badge/Python-3.12-blue.svg)
![CustomTkinter](https://img.shields.io/badge/CustomTkinter-UI-blueviolet.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

**Data-Travel** es una herramienta moderna de automatización de escritorio diseñada para la extracción, validación y migración masiva de datos estructurados entre libros de Excel locales y hojas de cálculo en la nube (Google Sheets). Fue concebida originalmente para consolidar y trasladar reportes de diferentes tipos de manera eficiente y a prueba de errores.

---

## Características Principales

*   **Integración Híbrida (Local y Nube):**
    *   **Excel Local:** Soporte nativo y rápido usando `openpyxl`.
    *   **Google Sheets:** Integración con la API de Google a través de `gspread`, implementando estrategias de `batch_update` para minimizar el consumo de cuotas y evitar bloqueos por rate limits.
    
*   **Motor de Coincidencia (Fuzzy Matching):**
    *   Mapeo inteligente de nombres de hojas/pestañas entre origen y destino utilizando `rapidfuzz` para lidiar con errores tipográficos, diferencias de mayúsculas/minúsculas o acentos.
    *   **Confirmación Interactiva:** Diálogo modal (MappingDialog) para revisar y aprobar las rutas de migración propuestas por el sistema antes de ejecutar la escritura.

*   **Modalidades de Transferencia (Migrador Universal):**
    *   **Bloque Continuo:** Extrae y pega matrices completas de datos (ej. `A2:D10`).
    *   **Saltos de Filas/Columnas:** Permite transferir datos lineales (1D) hacia el destino siguiendo un paso o _stride_ numérico configurable.
    *   **Listas de Celdas:** Parseo robusto para procesar selecciones múltiples no contiguas (ej. `C20, C21, C22` o `C20:C22, G20:G22`).
    
*   **Interfaz Gráfica Moderna (GUI):**
    *   Desarrollada sobre `CustomTkinter`.
    *   Soporte dinámico y completo para temas **Claro y Oscuro** sin artefactos visuales, con diseño elegante basado en tarjetas elevadas.
    *   Ejecución responsiva multihilo (`threading`), garantizando barras de progreso fluidas sin congelar la ventana principal.

---

## 🛠️ Stack Tecnológico

*   **Lenguaje:** [Python](https://www.python.org/)
*   **Interfaz Gráfica:** [CustomTkinter](https://customtkinter.tomschimansky.com/)
*   **Procesamiento de Excel:** [OpenPyXL](https://openpyxl.readthedocs.io/)
*   **Procesamiento de Google Sheets:** [GSpread](https://docs.gspread.org/), Google Auth, Google API Python Client.
*   **Fuzzy Matching:** [RapidFuzz](https://maxbachmann.github.io/RapidFuzz/)
*   **Compilación/Empaquetado:** [PyInstaller](https://pyinstaller.org/)

---

## 🚀 Guía de Instalación y Entorno Local

Sigue estos pasos para clonar el proyecto, instalar sus dependencias y ejecutarlo en modo de desarrollo.

### 1. Clonar el repositorio
```bash
git clone https://github.com/LAraujoA/Data-Travel.git
cd Data-Travel
```

### 2. Crear y activar el entorno virtual
```bash
# Crear entorno virtual
python -m venv venv

# Activar en Windows
venv\Scripts\activate

```

### 3. Instalar dependencias
```bash
pip install -r requirements.txt
```

### 4. Configurar Credenciales de Google Sheets
Para permitir la escritura en Google Sheets, necesitas una cuenta de servicio de Google Cloud:
1. Crea un proyecto en [Google Cloud Console](https://console.cloud.google.com/).
2. Habilita la **Google Sheets API** y **Google Drive API**.
3. Genera una Service Account y descarga su clave privada en formato JSON.
4. Nombra el archivo como `credentials.json` y colócalo en la raíz de este proyecto (o en el directorio base si ya usas el ejecutable).
5. **Importante:** Recuerda compartir (dar permisos de Editor) los Spreadsheets de destino al correo electrónico de tu Service Account.

### 5. Ejecutar la aplicación
```bash
python src/main.py
```

---

## 📦 Compilación (Ejecutable de Windows)

Data-Travel incluye un script preparado para autocompilar la aplicación en un archivo `.exe` portable para Windows, ideal para distribución a usuarios finales sin necesidad de instalar Python.

Ejecuta el script de empaquetado:
```bash
python build.py
```

El proceso:
1. Limpiará los directorios temporales `build/` y `dist/`.
2. Ocultará la consola por defecto y empaquetará todas las dependencias estáticas (RapidFuzz, CustomTkinter, Google Auth).
3. Generará el binario en: **`dist/Data-Travel.exe`**

---

## 📄 Licencia

Este proyecto está bajo la licencia **MIT**. Para más detalles, consulta el archivo [LICENSE](LICENSE) incluido en la raíz de este repositorio.

Copyright (c) 2026 Luis Araujo (LAraujoA)
