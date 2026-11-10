# 📦 Organizador Inteligente de Archivos (Prototipo Portable)

Organiza automáticamente cualquier carpeta de tu equipo (**Descargas, Escritorio, Documentos o cualquier ruta personalizada**) en carpetas según categorías inteligentes, mediante una interfaz gráfica moderna que corre 100% en tu propio computador.

---

## 🚀 Cómo ejecutar en 1 paso

### En Windows:
1. Haz doble clic en **`iniciar_organizador.bat`**.
2. ¡Listo! Se verificará Python y se abrirá la aplicación en tu navegador automáticamente en `http://localhost:8501`.

### En Mac / Linux:
1. Abre la terminal en esta carpeta.
2. Ejecuta:
   ```bash
   chmod +x iniciar_mac_linux.sh
   ./iniciar_mac_linux.sh
   ```

---

## ✨ Novedades y Características

- 📂 **Origen y Destino 100% dinámicos:** Puedes elegir ordenar Descargas, tu Escritorio, Documentos o escribir cualquier ruta personalizada, y elegir dónde guardar las carpetas creadas.
- 🧠 **Clasificación por Nombre y Contenido:** Inspecciona el nombre del archivo y el texto interno de documentos **.pdf, .docx y .txt**.
- 🗂️ **Categorías personalizadas:** Añade o elimina carpetas (Facturas, Universidad, Fotos, Trabajo...) a tu gusto.
- 👁️ **Vista previa con Casillas y Selector:**
  - Desmarca los archivos que desees conservar en su lugar original.
  - Modifica la carpeta destino asignada directamente en la tabla si prefieres moverlo a otro lado.
- ↩️ **Deshacer (Ctrl + Z):** Un botón dedicado revierte al instante la última organización si te arrepientes.
- 💾 **Memoria:** Guarda tu configuración en `config.json`.
- 🔒 **100% Privado y Seguro:** Ningún archivo sale de tu equipo ni se sube a servidores externos.

---

## 🛠️ Requisitos manuales (opcional)

Si prefieres ejecutarlo manualmente con comandos:

```bash
pip install -r requirements.txt
streamlit run app.py
```
