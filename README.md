# Script para Filtrado de Correos Gmail por Fechas y Lotes

Este script permite buscar y filtrar correos electrónicos en una cuenta de Gmail dentro de un rango de fechas determinado, segmentando la búsqueda en lotes mensuales.

---

## Requisitos Previos

- Tener una cuenta de Google.
- Acceso a [Google Cloud Console](https://console.cloud.google.com/).
- Google Chrome instalado para autenticación.

---

## Configuración Inicial

1. **Obtener credenciales de Google API**

   - Ingresa a Google Cloud Console y crea un nuevo proyecto.
   - Navega a **APIs y Servicios > Biblioteca**, busca y habilita la **API de Gmail**.
   - Luego, ve a **APIs y Servicios > Credenciales**.
   - Haz clic en **Crear credenciales > ID de cliente OAuth**.
   - Selecciona **Aplicación de escritorio** como tipo de aplicación.
   - Descarga el archivo JSON generado y renómbralo (si lo deseas) como `credentials.json`.

2. **Ubicación del archivo `credentials.json`**

   - Guarda el archivo `credentials.json` en una ruta accesible para el script.
   - Al ejecutar el script, deberás indicar la ruta donde se encuentra este archivo.

---

## Uso del Script

1. Al iniciar, el script abrirá Google Chrome para que inicies sesión y autorices el acceso a tu cuenta de Gmail.

2. Ingresa la fecha inicial y la fecha final para filtrar los correos que deseas consultar.

3. Define el tamaño de los lotes en meses.  
   Por ejemplo, si seleccionas un rango del 01/01/2025 al 01/01/2026 y estableces lotes de 3 meses, el script realizará búsquedas segmentadas de 3 en 3 meses dentro de ese rango.

---

## Instalación de Dependencias

Antes de ejecutar el script, asegúrate de instalar las dependencias necesarias con:

```bash
pip install -r requirements.txt

Notas
En caso de realizar modificaciones al script, recuerda actualizar las dependencias en requirements.txt.

El proceso de autenticación es requerido una única vez, a menos que cambies de cuenta o modifiques las credenciales.
