"""
================================================================================
  ORGANIZADOR INTELIGENTE DE ARCHIVOS (Desktop/Web UI)
  Ejecutar: streamlit run app.py
  O doble clic en: iniciar_organizador.bat
================================================================================
"""

import os
import json
import shutil
import socket
from pathlib import Path
import streamlit as st

# ─── Librerías opcionales para lectura profunda de contenido ──────────────────
try:
    import pdfplumber
    PDF_OK = True
except ImportError:
    PDF_OK = False

try:
    from docx import Document
    DOCX_OK = True
except ImportError:
    DOCX_OK = False

# ══════════════════════════════════════════════════════════════════════════════
#  CONFIGURACIÓN DE STREAMLIT
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Organizador Inteligente de Archivos",
    page_icon="📦",
    layout="centered",
    initial_sidebar_state="collapsed",
)

RUTA_BASE     = Path(__file__).parent.resolve()
RUTA_CONFIG   = RUTA_BASE / "config.json"
RUTA_UNDO_LOG = RUTA_BASE / "undo_log.json"

# ─── ESTILOS MODERNOS (GLASSMORPHISM + GRADIENTES) ───────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"] {
    background: radial-gradient(circle at 50% 0%, #171b2d 0%, #0c0e16 75%, #08090f 100%);
    color: #e6ebf5;
    font-family: system-ui, -apple-system, sans-serif;
}
[data-testid="stHeader"] { background: transparent; }

/* Contenedores con efecto cristal */
div[data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlock"] {
    background: rgba(255, 255, 255, 0.035);
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: 16px;
    padding: 1.25rem;
    box-shadow: 0 8px 32px rgba(0, 0, 0, 0.25);
    backdrop-filter: blur(8px);
}

/* Botones principales */
div[data-testid="stButton"] > button[kind="primary"] {
    background: linear-gradient(135deg, #7c6af7 0%, #5846e0 100%) !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    padding: 0.65rem 1.4rem !important;
    box-shadow: 0 4px 18px rgba(124, 106, 247, 0.38) !important;
    transition: all 0.2s ease !important;
}
div[data-testid="stButton"] > button[kind="primary"]:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 6px 24px rgba(124, 106, 247, 0.55) !important;
}

/* Botones secundarios */
div[data-testid="stButton"] > button[kind="secondary"] {
    background: rgba(255, 255, 255, 0.06) !important;
    color: #e2e8f0 !important;
    border: 1px solid rgba(255, 255, 255, 0.12) !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    transition: all 0.2s ease !important;
}
div[data-testid="stButton"] > button[kind="secondary"]:hover {
    background: rgba(255, 255, 255, 0.12) !important;
    border-color: rgba(124, 106, 247, 0.4) !important;
}

/* Inputs de texto y selectboxes */
div[data-testid="stTextInput"] input {
    background: rgba(255, 255, 255, 0.05) !important;
    border: 1px solid rgba(255, 255, 255, 0.12) !important;
    border-radius: 10px !important;
    color: #f1f5f9 !important;
}
div[data-testid="stTextInput"] input:focus {
    border-color: #7c6af7 !important;
    box-shadow: 0 0 0 2px rgba(124, 106, 247, 0.25) !important;
}
div[data-testid="stSelectbox"] > div > div {
    background: rgba(255, 255, 255, 0.05) !important;
    border: 1px solid rgba(255, 255, 255, 0.12) !important;
    border-radius: 10px !important;
    color: #f1f5f9 !important;
}

/* Tarjetas de métricas */
div[data-testid="stMetric"] {
    background: rgba(124, 106, 247, 0.08) !important;
    border: 1px solid rgba(124, 106, 247, 0.25) !important;
    border-radius: 12px !important;
    padding: 0.8rem 1rem !important;
}

/* Tablas interactivas */
div[data-testid="stDataFrame"] {
    border-radius: 12px !important;
    border: 1px solid rgba(255, 255, 255, 0.08) !important;
    overflow: hidden !important;
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  DETECCIÓN INTELIGENTE DE RUTAS (WINDOWS / ONEDRIVE / MAC / LINUX)
# ══════════════════════════════════════════════════════════════════════════════

def _detectar_ruta(tipo: str) -> Path:
    """
    Detecta la ruta real de carpetas de usuario, resolviendo automáticamente
    OneDrive o rutas personalizadas en Windows, y funcionando en Mac/Linux.
    """
    home = Path.home()
    
    # 1. Intentar por registro de Windows
    claves_win = {
        "downloads": ("{374DE290-123F-4565-9164-39C4925E467B}", ["Downloads", "Descargas"]),
        "desktop":   ("Desktop", ["Desktop", "Escritorio"]),
        "documents": ("Personal", ["Documents", "Documentos"]),
        "pictures":  ("My Pictures", ["Pictures", "Imágenes", "Imagenes", "Fotos"]),
        "music":     ("My Music", ["Music", "Música", "Musica"]),
        "videos":    ("My Video", ["Videos", "Vídeos"]),
    }
    
    reg_key, nombres = claves_win.get(tipo, ("", []))
    
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
        )
        ruta_str, _ = winreg.QueryValueEx(key, reg_key)
        winreg.CloseKey(key)
        cand = Path(ruta_str)
        if cand.exists():
            return cand
    except Exception:
        pass
    
    # 2. Búsqueda en OneDrive (común en Windows 10/11)
    onedrive_folders = [
        home / "OneDrive",
        home / "OneDrive - Personal",
    ]
    for od in onedrive_folders:
        for nombre in nombres:
            cand = od / nombre
            if cand.exists():
                return cand

    # 3. Búsqueda directa en Home del usuario
    for nombre in nombres:
        cand = home / nombre
        if cand.exists():
            return cand

    # Fallback predeterminado
    return home / (nombres[0] if nombres else "Downloads")

# Diccionario accesible de carpetas predefinidas
CARPETAS_PREDEFINIDAS = {
    "📥 Descargas":    _detectar_ruta("downloads"),
    "🖥️ Escritorio":   _detectar_ruta("desktop"),
    "📄 Documentos":   _detectar_ruta("documents"),
    "🖼️ Imágenes":     _detectar_ruta("pictures"),
    "🎵 Música":       _detectar_ruta("music"),
    "🎬 Videos":       _detectar_ruta("videos"),
    "📂 Otra ubicación personalizada...": None,
}

# ══════════════════════════════════════════════════════════════════════════════
#  GESTIÓN DE CONFIGURACIÓN Y DESHACER (JSON)
# ══════════════════════════════════════════════════════════════════════════════

def cargar_config() -> dict:
    if RUTA_CONFIG.exists():
        try:
            with open(RUTA_CONFIG, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def guardar_config(datos: dict):
    with open(RUTA_CONFIG, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=4)

def guardar_undo(movidos: list):
    with open(RUTA_UNDO_LOG, "w", encoding="utf-8") as f:
        json.dump(movidos, f, ensure_ascii=False, indent=4)

def cargar_undo() -> list:
    if RUTA_UNDO_LOG.exists():
        try:
            with open(RUTA_UNDO_LOG, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return []

def deshacer() -> dict:
    log = cargar_undo()
    if not log:
        return {"restaurados": [], "errores": [], "vacio": True}

    restaurados, errores = [], []
    for entrada in log:
        src = Path(entrada["desde"])
        dst = Path(entrada["hasta"])

        if not src.exists():
            errores.append({"nombre": src.name, "error": "El archivo ya no se encuentra en la carpeta destino."})
            continue

        try:
            dst_final = dst
            if dst_final.exists():
                base, ext, i = dst.stem, dst.suffix, 1
                while dst_final.exists():
                    dst_final = dst.parent / f"{base}_restaurado_{i}{ext}"
                    i += 1

            shutil.move(str(src), str(dst_final))
            restaurados.append({
                "📄 Archivo": src.name,
                "📁 De carpeta": src.parent.name,
                "📍 Restaurado en": dst_final.name,
            })
        except Exception as ex:
            errores.append({"nombre": src.name, "error": str(ex)})

    if RUTA_UNDO_LOG.exists():
        try:
            RUTA_UNDO_LOG.unlink()
        except Exception:
            pass

    return {"restaurados": restaurados, "errores": errores, "vacio": False}

# ══════════════════════════════════════════════════════════════════════════════
#  LECTURA DE CONTENIDO INTERNO (.TXT, .PDF, .DOCX)
# ══════════════════════════════════════════════════════════════════════════════

def leer_contenido(ruta: Path) -> str:
    ext = ruta.suffix.lower()
    try:
        if ext == ".txt":
            return ruta.read_text(encoding="utf-8", errors="ignore").lower()
        if ext == ".pdf" and PDF_OK:
            texto = ""
            with pdfplumber.open(ruta) as pdf:
                for pag in pdf.pages[:5]:  # Lee las primeras 5 páginas para velocidad
                    txt = pag.extract_text()
                    if txt:
                        texto += txt + " "
            return texto.lower()
        if ext == ".docx" and DOCX_OK:
            doc = Document(ruta)
            return " ".join(p.text for p in doc.paragraphs[:40]).lower()
    except Exception:
        pass
    return ""

# ══════════════════════════════════════════════════════════════════════════════
#  CLASIFICACIÓN INTELIGENTE (NOMBRE + CONTENIDO)
# ══════════════════════════════════════════════════════════════════════════════

def clasificar(archivo: Path, carpetas: list) -> str:
    nombre_limpio = archivo.stem.lower()
    contenido = leer_contenido(archivo)
    todo_texto = f"{nombre_limpio} {contenido}"

    mejor_carpeta, mejor_puntaje = "Otros", 0

    for c in carpetas:
        palabras = [p.strip() for p in c.lower().split() if len(p.strip()) > 2]
        if not palabras:
            palabras = [c.lower().strip()]

        puntaje = 0
        for palabra in palabras:
            # Coincidencia en el nombre tiene más peso (+3)
            if palabra in nombre_limpio:
                puntaje += 3
            # Coincidencia en el contenido (+1)
            elif palabra in todo_texto:
                puntaje += 1

        if puntaje > mejor_puntaje:
            mejor_puntaje = puntaje
            mejor_carpeta = c

    return mejor_carpeta

def listar_preview(origen: Path, carpetas: list) -> list[dict]:
    items = []
    if not origen or not origen.exists():
        return items

    for elem in sorted(origen.iterdir(), key=lambda x: x.name.lower()):
        if elem.is_dir() or elem.name.startswith("."):
            continue
        try:
            tam_kb = round(elem.stat().st_size / 1024, 1)
            items.append({
                "✅ Incluir": True,
                "📄 Archivo": elem.name,
                "📁 Destino": clasificar(elem, carpetas),
                "📏 Tamaño": f"{tam_kb} KB" if tam_kb < 1024 else f"{tam_kb/1024:.1f} MB",
                "🔖 Ext.": elem.suffix.lower() or "(sin ext.)",
            })
        except Exception:
            continue
    return items

def ejecutar_organizacion(df_seleccionados, origen: Path, destino: Path, carpetas: list) -> dict:
    carpetas_completas = carpetas + ["Otros"]
    for c in carpetas_completas:
        (destino / c).mkdir(parents=True, exist_ok=True)

    movidos, errores, omitidos, undo_entries = [], [], [], []

    for _, fila in df_seleccionados.iterrows():
        nombre = fila["📄 Archivo"]
        incluir = fila["✅ Incluir"]
        dest_carpeta = fila["📁 Destino"]

        src = origen / nombre

        if not incluir:
            omitidos.append(nombre)
            continue

        if not src.exists():
            errores.append({"nombre": nombre, "error": "El archivo ya no existe en el origen."})
            continue

        try:
            dest_dir = destino / dest_carpeta
            dest_file = dest_dir / nombre

            # Manejo de archivos duplicados con sufijo numérico
            if dest_file.exists():
                base, ext, i = src.stem, src.suffix, 1
                while dest_file.exists():
                    dest_file = dest_dir / f"{base}_{i}{ext}"
                    i += 1

            shutil.move(str(src), str(dest_file))
            movidos.append({
                "📄 Archivo original": nombre,
                "📁 Carpeta destino": dest_carpeta,
                "📍 Guardado como": dest_file.name
            })
            undo_entries.append({
                "desde": str(dest_file),
                "hasta": str(src)
            })
        except Exception as ex:
            errores.append({"nombre": nombre, "error": str(ex)})

    if undo_entries:
        guardar_undo(undo_entries)

    return {"movidos": movidos, "errores": errores, "omitidos": omitidos}

def get_ip() -> str:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "localhost"

# ══════════════════════════════════════════════════════════════════════════════
#  APLICACIÓN PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════

def main():
    import pandas as pd

    # ── CABECERA / HERO ──────────────────────────────────────────────────────
    st.markdown("""
    <div style="text-align:center; padding: 2.2rem 0 1.2rem;">
      <div style="font-size: 3.8rem; margin-bottom: 0.4rem;">📦</div>
      <h1 style="font-size: 2.4rem; font-weight: 900; margin: 0;
                 background: linear-gradient(90deg, #7c6af7, #5de0d8);
                 -webkit-background-clip: text; -webkit-text-fill-color: transparent;">
        Organizador Inteligente
      </h1>
      <p style="color: #94a3b8; font-size: 1.05rem; margin-top: 0.6rem; max-width: 580px; margin-left: auto; margin-right: auto;">
        Elige cualquier carpeta de tu computador, define tus categorías y deja que la app clasifique tus archivos con seguridad total.
      </p>
    </div>
    """, unsafe_allow_html=True)

    ip_actual = get_ip()
    st.info(f"🌐 **Acceso local:** `http://localhost:8501`  ·  **En tu misma red WiFi:** `http://{ip_actual}:8501`", icon=None)

    # ── CARGA DE ESTADO ──────────────────────────────────────────────────────
    cfg = cargar_config()
    carpetas_guardadas = cfg.get("carpetas", ["Facturas", "Documentos", "Universidad", "Fotos", "Trabajo"])

    if "carpetas" not in st.session_state:
        st.session_state.carpetas = carpetas_guardadas
    if "modo_editor" not in st.session_state:
        st.session_state.modo_editor = False
    if "preview_df" not in st.session_state:
        st.session_state.preview_df = None
    if "resultado" not in st.session_state:
        st.session_state.resultado = None

    carpetas = st.session_state.carpetas

    st.divider()

    # ══════════════════════════════════════════════════════════════════════════
    #  PASO 0: ELEGIR UBICACIONES (ORIGEN Y DESTINO DINÁMICOS)
    # ══════════════════════════════════════════════════════════════════════════
    with st.container(border=True):
        st.markdown("### 00 · ¿Qué carpeta deseas organizar?")
        st.caption("Puedes elegir Descargas, el Escritorio, Documentos o cualquier carpeta personalizada de tu equipo.")

        col_orig, col_dest = st.columns(2)

        # ── 1. Origen
        with col_orig:
            st.markdown("**📂 Carpeta a Organizar (Origen)**")
            opciones_origen = list(CARPETAS_PREDEFINIDAS.keys())
            def_orig_idx = 0
            if "origen_sel" in cfg and cfg["origen_sel"] in opciones_origen:
                def_orig_idx = opciones_origen.index(cfg["origen_sel"])

            origen_sel = st.selectbox("Origen", opciones_origen, index=def_orig_idx,
                                      label_visibility="collapsed", key="select_origen")

            if origen_sel.startswith("📂"):
                custom_orig = st.text_input(
                    "Escribe la ruta completa del origen:",
                    value=cfg.get("origen_custom", ""),
                    placeholder=r"Ej: D:\Proyectos o C:\Users\Nombre\Carpeta",
                    key="input_origen_custom"
                )
                ruta_origen = Path(custom_orig) if custom_orig.strip() else None
            else:
                ruta_origen = CARPETAS_PREDEFINIDAS[origen_sel]

            if ruta_origen and ruta_origen.exists():
                try:
                    num_archivos = sum(1 for e in ruta_origen.iterdir() if e.is_file() and not e.name.startswith("."))
                    st.caption(f"✅ `{ruta_origen}` ({num_archivos} archivo(s))")
                except Exception:
                    st.caption(f"✅ `{ruta_origen}`")
            elif ruta_origen:
                st.error("❌ Esta ruta no existe en tu equipo.", icon="❌")

        # ── 2. Destino
        with col_dest:
            st.markdown("**📁 Dónde guardar las carpetas (Destino)**")
            opciones_destino = [k for k in CARPETAS_PREDEFINIDAS.keys() if not k.startswith("📥")] + ["📥 Descargas"]
            def_dest_idx = 0
            if "destino_sel" in cfg and cfg["destino_sel"] in opciones_destino:
                def_dest_idx = opciones_destino.index(cfg["destino_sel"])

            destino_sel = st.selectbox("Destino", opciones_destino, index=def_dest_idx,
                                       label_visibility="collapsed", key="select_destino")

            if destino_sel.startswith("📂"):
                custom_dest = st.text_input(
                    "Escribe la ruta completa del destino:",
                    value=cfg.get("destino_custom", ""),
                    placeholder=r"Ej: C:\Users\Nombre\Organizado",
                    key="input_destino_custom"
                )
                ruta_destino = Path(custom_dest) if custom_dest.strip() else None
            else:
                ruta_destino = CARPETAS_PREDEFINIDAS.get(destino_sel)

            if ruta_destino and ruta_destino.exists():
                st.caption(f"📁 `{ruta_destino}`")
            elif ruta_destino:
                st.warning("⚠️ La carpeta aún no existe, pero se creará al organizar.")

        # Guardar selección en config
        if ruta_origen and ruta_destino:
            cfg_actual = cargar_config()
            cfg_actual["origen_sel"] = origen_sel
            cfg_actual["destino_sel"] = destino_sel
            if origen_sel.startswith("📂"):
                cfg_actual["origen_custom"] = str(ruta_origen)
            if destino_sel.startswith("📂"):
                cfg_actual["destino_custom"] = str(ruta_destino)
            guardar_config(cfg_actual)

    # ══════════════════════════════════════════════════════════════════════════
    #  PASO 1: CATEGORÍAS PERSONALIZADAS
    # ══════════════════════════════════════════════════════════════════════════
    with st.container(border=True):
        c_title, c_btn = st.columns([4, 1])
        with c_title:
            st.markdown("### 01 · Mis Carpetas de Clasificación")
            st.caption("Crea o ajusta los nombres de las carpetas donde se agruparán tus archivos.")
        with c_btn:
            lbl = "✏️ Editar" if not st.session_state.modo_editor else "✖ Cerrar"
            if st.button(lbl, use_container_width=True):
                st.session_state.modo_editor = not st.session_state.modo_editor
                st.rerun()

        # Vista de solo lectura (chips)
        if not st.session_state.modo_editor:
            if carpetas:
                chips_html = " ".join([
                    f'<span style="background:rgba(124,106,247,0.15); border:1px solid rgba(124,106,247,0.38); '
                    f'border-radius:8px; padding:5px 13px; font-size:0.88rem; font-weight:600; color:#c4baff; display:inline-block; margin:3px 2px;">'
                    f'📁 {c}</span>'
                    for c in carpetas
                ])
                chips_html += (' <span style="background:rgba(255,255,255,0.06); border:1px solid rgba(255,255,255,0.12); '
                               'border-radius:8px; padding:5px 13px; font-size:0.88rem; color:#94a3b8; display:inline-block; margin:3px 2px;">'
                               '📁 Otros (sin coincidencia)</span>')
                st.markdown(chips_html, unsafe_allow_html=True)
                st.caption(f"💾 Configuración guardada en: `{RUTA_CONFIG}`")
            else:
                st.warning("No tienes carpetas definidas. Haz clic en **✏️ Editar** para agregar tus carpetas.", icon="⚠️")

        # Modo editor interactivo
        else:
            if "carpetas_tmp" not in st.session_state:
                st.session_state.carpetas_tmp = list(carpetas)

            col_in, col_add = st.columns([3, 1])
            with col_in:
                nueva = st.text_input("Nueva carpeta", placeholder="Ej: Facturas, Tareas, Música...",
                                      label_visibility="collapsed", key="input_nueva_cat")
            with col_add:
                if st.button("➕ Agregar", type="primary", use_container_width=True):
                    nueva_limpia = nueva.strip()
                    if not nueva_limpia:
                        st.toast("⚠️ Ingresa un nombre para la carpeta.")
                    elif nueva_limpia.lower() in [c.lower() for c in st.session_state.carpetas_tmp]:
                        st.toast(f"'{nueva_limpia}' ya está en tu lista.")
                    else:
                        st.session_state.carpetas_tmp.append(nueva_limpia)
                        st.rerun()

            if st.session_state.carpetas_tmp:
                st.markdown("**Carpetas actuales (clic en ✕ para eliminar):**")
                cols = st.columns(min(len(st.session_state.carpetas_tmp), 4) or 1)
                eliminar_idx = None
                for i, n in enumerate(st.session_state.carpetas_tmp):
                    with cols[i % 4]:
                        if st.button(f"📁 {n} ✕", key=f"del_cat_{i}", use_container_width=True):
                            eliminar_idx = i
                if eliminar_idx is not None:
                    st.session_state.carpetas_tmp.pop(eliminar_idx)
                    st.rerun()

            st.markdown("")
            c_canc, c_save = st.columns(2)
            with c_canc:
                if st.button("Cancelar", use_container_width=True):
                    del st.session_state.carpetas_tmp
                    st.session_state.modo_editor = False
                    st.rerun()
            with c_save:
                if st.button("💾 Guardar Cambios", type="primary", use_container_width=True):
                    if not st.session_state.carpetas_tmp:
                        st.toast("⚠️ Agrega al menos una carpeta.")
                    else:
                        st.session_state.carpetas = list(st.session_state.carpetas_tmp)
                        cfg_save = cargar_config()
                        cfg_save["carpetas"] = st.session_state.carpetas
                        guardar_config(cfg_save)
                        del st.session_state.carpetas_tmp
                        st.session_state.modo_editor = False
                        st.session_state.preview_df = None  # Reiniciar preview para recalcular
                        st.toast("✅ Configuración guardada.")
                        st.rerun()

    # ══════════════════════════════════════════════════════════════════════════
    #  PASO 2: VISTA PREVIA INTERACTIVA CON CHECKBOXES Y SELECTBOX
    # ══════════════════════════════════════════════════════════════════════════
    with st.container(border=True):
        c_title, c_btn = st.columns([4, 1])
        with c_title:
            st.markdown("### 02 · Vista Previa Inteligente")
            st.caption("Revisa cómo se clasificarán tus archivos. Puedes desmarcar casillas o cambiar la carpeta destino.")
        with c_btn:
            if st.button("🔍 Analizar", use_container_width=True):
                if not carpetas:
                    st.error("Define tus carpetas en el Paso 01 primero.", icon="❌")
                elif not ruta_origen or not ruta_origen.exists():
                    st.error("La carpeta de origen seleccionada no existe.", icon="❌")
                else:
                    with st.spinner("Analizando nombres y contenido de archivos..."):
                        archivos = listar_preview(ruta_origen, carpetas)
                    st.session_state.preview_df = pd.DataFrame(archivos)
                    st.session_state.resultado = None

        if st.session_state.preview_df is not None:
            df = st.session_state.preview_df

            if df.empty:
                st.info("✨ No se encontraron archivos para organizar en la carpeta origen seleccionada.", icon="✨")
            else:
                total = len(df)
                n_sel = int(df["✅ Incluir"].sum())
                n_otros = int((df["📁 Destino"] == "Otros").sum())

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("📄 Total encontrados", total)
                c2.metric("☑️ Seleccionados", n_sel)
                c3.metric("⬜ Excluidos", total - n_sel)
                c4.metric("📂 Asignados a Otros", n_otros)

                st.caption("💡 **Tip:** Desmarca la casilla si no deseas mover un archivo, o cámbiale la carpeta destino en la columna **📁 Destino**.")

                opciones_destinos_tabla = carpetas + ["Otros"]

                # Editor interactivo donde el usuario puede activar/desactivar y cambiar carpeta
                edited = st.data_editor(
                    df,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "✅ Incluir": st.column_config.CheckboxColumn(
                            "✅ Incluir",
                            help="Marca para mover este archivo; desmarca para dejarlo donde está.",
                            default=True
                        ),
                        "📄 Archivo": st.column_config.TextColumn(
                            "📄 Archivo",
                            width="large"
                        ),
                        "📁 Destino": st.column_config.SelectboxColumn(
                            "📁 Destino",
                            help="Puedes cambiar la carpeta destino sugerida si lo prefieres.",
                            options=opciones_destinos_tabla,
                            required=True,
                            width="medium"
                        ),
                        "📏 Tamaño": st.column_config.TextColumn("📏 Tamaño", width="small"),
                        "🔖 Ext.": st.column_config.TextColumn("🔖 Ext.", width="small"),
                    },
                    disabled=["📄 Archivo", "📏 Tamaño", "🔖 Ext."],
                    key="editor_vista_previa"
                )

                st.session_state.preview_df = edited
                n_sel_actual = int(edited["✅ Incluir"].sum())

                if n_sel_actual == 0:
                    st.warning("Todos los archivos están desmarcados. Marca al menos uno para organizar.", icon="⚠️")
                else:
                    st.success(f"Se moverán **{n_sel_actual}** de {total} archivo(s) hacia `{ruta_destino}`", icon="✅")

        else:
            st.caption("Haz clic en **🔍 Analizar** para examinar la carpeta de origen.")

    # ══════════════════════════════════════════════════════════════════════════
    #  PASO 3: ORGANIZAR + BOTÓN DESHACER (CTRL + Z)
    # ══════════════════════════════════════════════════════════════════════════
    with st.container(border=True):
        st.markdown("### 03 · Aplicar Cambios")
        st.caption("Moverá únicamente los archivos marcados con ✅ hacia sus carpetas correspondientes.")

        hay_seleccion = (
            st.session_state.preview_df is not None
            and not st.session_state.preview_df.empty
            and st.session_state.preview_df["✅ Incluir"].any()
        )
        hay_undo = RUTA_UNDO_LOG.exists()

        c_org, c_und = st.columns(2)

        with c_org:
            btn_organizar = st.button(
                "✨ Organizar Ahora",
                type="primary",
                use_container_width=True,
                disabled=not (carpetas and hay_seleccion and ruta_origen and ruta_destino),
                help="Procesa y mueve los archivos seleccionados."
            )

        with c_und:
            btn_deshacer = st.button(
                "↩️ Deshacer Última Organización (Ctrl+Z)",
                use_container_width=True,
                disabled=not hay_undo,
                help="Devuelve todos los archivos de la última organización a su carpeta original." if hay_undo else "No hay operaciones recientes para deshacer."
            )

        if not hay_undo:
            st.caption("El botón **↩️ Deshacer** se activa automáticamente tras completar una organización.")

        # ── Ejecutar Organización
        if btn_organizar:
            if not ruta_origen or not ruta_origen.exists():
                st.error("La carpeta origen no es válida.", icon="❌")
            elif not ruta_destino:
                st.error("Debes seleccionar una carpeta destino válida.", icon="❌")
            else:
                with st.spinner("Moviendo archivos a sus carpetas destino..."):
                    res = ejecutar_organizacion(
                        st.session_state.preview_df,
                        ruta_origen,
                        ruta_destino,
                        carpetas
                    )
                st.session_state.resultado = res
                st.session_state.preview_df = None  # Limpiar preview ya que los archivos ya fueron movidos
                st.rerun()

        # ── Ejecutar Deshacer
        if btn_deshacer:
            with st.spinner("Revirtiendo archivos a su ubicación original..."):
                res_u = deshacer()

            if res_u["vacio"]:
                st.info("No hay nada que revertir.", icon="ℹ️")
            else:
                restaurados = res_u["restaurados"]
                if restaurados:
                    st.success(f"↩️ ¡{len(restaurados)} archivo(s) devueltos exitosamente!", icon="✅")
                    with st.expander(f"📋 Ver detalle de los {len(restaurados)} archivos restaurados"):
                        st.dataframe(pd.DataFrame(restaurados), use_container_width=True, hide_index=True)
                if res_u["errores"]:
                    with st.expander(f"⚠️ {len(res_u['errores'])} advertencia(s) al deshacer"):
                        for e in res_u["errores"]:
                            st.error(f"**{e['nombre']}** — {e['error']}")
                st.rerun()

        # ── Reporte de resultados
        if st.session_state.resultado:
            res = st.session_state.resultado
            movidos = res["movidos"]
            omitidos = res["omitidos"]
            errores = res["errores"]

            st.divider()

            if movidos:
                st.success(f"🎉 ¡{len(movidos)} archivo(s) organizados con éxito en `{ruta_destino}`!", icon="✅")
                if omitidos:
                    st.info(f"⬜ {len(omitidos)} archivo(s) fueron omitidos según tu selección.", icon="ℹ️")

                with st.expander(f"📋 Ver detalle ({len(movidos)} archivos movidos)"):
                    st.dataframe(pd.DataFrame(movidos), use_container_width=True, hide_index=True)

            if errores:
                with st.expander(f"⚠️ {len(errores)} error(es)"):
                    for e in errores:
                        st.error(f"**{e['nombre']}** — {e['error']}")

    # ── PIE DE PÁGINA ────────────────────────────────────────────────────────
    st.divider()
    st.markdown(
        f"<p style='text-align:center; color:rgba(255,255,255,0.25); font-size:0.82rem;'>"
        f"Organizador Inteligente de Archivos · Corre de forma local y 100% privada · Red WiFi: http://{ip_actual}:8501"
        f"</p>",
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
