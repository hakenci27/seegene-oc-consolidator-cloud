"""
app.py
Version para Streamlit Cloud del consolidador de Nuevas OC.

A diferencia de la version anterior (que hacia una extraccion simplificada
sin cruzar contra archivos maestros), esta version reutiliza EL MISMO motor
que la version local (oc_consolidator_clean/: extractor.py, matcher.py,
builder.py, main.py, config.py -- copiados tal cual en esta carpeta para que
el repo de este deploy sea autocontenido). La diferencia real con la version
local es solo de dónde salen los archivos:

- Local: lee la "Carpeta de trabajo" del disco del usuario.
- Cloud: cada archivo maestro, cada OC, y opcionalmente el Excel de control
  ya existente, se suben desde el navegador. Se escriben a una carpeta
  temporal (una por corrida, aislada por sesion) y se le pasa al mismo
  main.run() de siempre -- asi que el cruce contra Product master list,
  customer master list (con alias), precios, credito, etc. funciona IGUAL
  que en la version local.

IMPORTANTE: si se corrigen bugs en matcher.py/builder.py/main.py de la
version local, hay que copiar esos mismos archivos aqui tambien (no se
comparten automaticamente entre las dos carpetas).

Uso local:
    streamlit run app.py

Requiere ANTHROPIC_API_KEY en st.secrets (ver README.md).
"""

import json
import shutil
import tempfile
from datetime import datetime
from pathlib import Path

import streamlit as st

import main as oc_main
from config import (
    OCS_SUBFOLDER,
    FILE_SALES_PROGRESS, FILE_CUSTOMER_MASTER, FILE_DAILY_REPORT,
    FILE_PRICE_LIST, FILE_CREDIT_REPORT, FILE_PRODUCT_MASTER,
)

DEFAULT_OUTPUT_FILENAME = "Nuevas OC - Control de Ordenes de Compra ver2.xlsx"


def _find_logo():
    assets_dir = Path(__file__).parent / "assets"
    for name in ("seegene_logo.png", "seegene_logo.jpg", "seegene_logo.jpeg", "seegene_logo.svg"):
        p = assets_dir / name
        if p.exists():
            return p
    return None


LOGO_PATH = _find_logo()

MASTER_FILES = [
    (FILE_CUSTOMER_MASTER, "Codigo de cliente, PIC, clasificacion, RFC"),
    (FILE_SALES_PROGRESS, "OC ya registradas + categoria de producto"),
    (FILE_DAILY_REPORT, "Precio de referencia + inventario/caducidad"),
    (FILE_PRICE_LIST, "Lista de precios por clasificacion"),
    (FILE_CREDIT_REPORT, "Limite de credito (CREDITO)"),
    (FILE_PRODUCT_MASTER, "Codigo/Descripcion/Categoria de producto"),
]

WARNING_LABELS = {
    "master_file_missing": "Archivo maestro no encontrado",
    "customer_not_found": "Cliente no encontrado en customer master list",
    "code_uncertain": "Codigo de producto incierto (revisar el documento original)",
    "extraction_error": "Error al leer el documento",
}


def get_api_key():
    try:
        return st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        return None


st.set_page_config(page_title="Consolidador de Nuevas OC (Cloud)", page_icon="☁️", layout="wide")

header_col, logo_col = st.columns([5, 1])
with header_col:
    st.title("☁️ Consolidador de Nuevas OC — Cloud")
    st.caption(
        "Version en la nube disponible para todo el equipo desde este enlace. Usa el "
        "mismo motor que la version local -- cruce automatico de codigo de cliente/"
        "producto, verificacion de limite de credito, etc. -- con la unica diferencia "
        "de que este servidor no tiene acceso a los archivos de tu computadora, asi "
        "que hay que subir los archivos maestros cada vez."
    )
with logo_col:
    if LOGO_PATH is not None:
        st.image(str(LOGO_PATH), use_container_width=True)

with st.sidebar:
    st.markdown("## ⚙️ Configuracion")
    fake_mode = st.checkbox(
        "🧪 Modo de prueba (sin API, datos ficticios)",
        value=False,
        help=(
            "No llama a Claude ni cuesta nada -- genera una OC ficticia por cada "
            "archivo subido, marcada claramente como datos de prueba, para probar "
            "el flujo completo (subida, cruce con maestros, descarga del Excel) "
            "sin gastar API."
        ),
    )
    if fake_mode:
        st.caption("🧪 Modo de prueba activo -- los datos de las OC seran ficticios.")

api_key = get_api_key()
if not api_key and not fake_mode:
    st.error(
        "No se ha configurado la API key de Claude. Agrega `ANTHROPIC_API_KEY` en "
        "`.streamlit/secrets.toml` (local) o en Settings > Secrets de Streamlit "
        "Cloud (ver README.md). O activa el modo de prueba en la barra lateral "
        "para probar el flujo sin API key."
    )
    st.stop()

# ---------------------------------------------------------------------------
# Barra lateral: subir archivos maestros + configuracion
# ---------------------------------------------------------------------------
with st.sidebar:
    with st.container(border=True):
        st.markdown("**🗂️ Archivos maestros**")
        st.caption(
            "No son obligatorios (si falta alguno, solo se pierde esa funcion "
            "especifica), pero si los subes se cruzan automaticamente los codigos "
            "de cliente/producto. Si te resulta pesado subir los mismos archivos "
            "cada dia, guardalos en una carpeta en tu computadora para arrastrarlos "
            "rapido cada vez."
        )
        master_uploads = {}
        for filename, desc in MASTER_FILES:
            f = st.file_uploader(
                filename, type=["xlsx", "xlsm"], key=f"master_{filename}",
                help=desc,
            )
            master_uploads[filename] = f
        found = sum(1 for f in master_uploads.values() if f is not None)
        if found == len(MASTER_FILES):
            st.success(f"{found} / {len(MASTER_FILES)} subidos")
        else:
            st.warning(f"{found} / {len(MASTER_FILES)} subidos")

    with st.container(border=True):
        st.markdown("**📄 Archivo de control existente (opcional)**")
        st.caption(
            "Si ya tienes el Excel acumulado hasta ayer, subelo aqui -- las OC "
            "nuevas se agregan al final de ese archivo (sin duplicar, sin borrar "
            "las columnas que llenaste a mano). Si no subes nada, se crea uno nuevo."
        )
        previous_output = st.file_uploader(
            "Archivo existente Nuevas OC - Control...xlsx", type=["xlsx"], key="previous_output",
            label_visibility="collapsed",
        )
        if st.session_state.get("cloud_accumulated_bytes"):
            if previous_output is not None:
                st.caption("ℹ️ Se usara el archivo que acabas de subir arriba (tiene prioridad).")
            else:
                st.caption(
                    f"✅ Se seguira acumulando automaticamente sobre el resultado de "
                    f"esta sesion ({st.session_state.get('cloud_session_total_added', 0)} "
                    f"linea(s) hasta ahora)."
                )
            if st.button("🔄 Empezar de cero en esta sesion", use_container_width=True):
                st.session_state["cloud_accumulated_bytes"] = None
                st.session_state["cloud_session_total_added"] = 0
                st.rerun()

    with st.container(border=True):
        st.markdown("**💾 Nombre del archivo de salida**")
        output_filename = st.text_input(
            "Nombre del archivo de salida", value=DEFAULT_OUTPUT_FILENAME, label_visibility="collapsed",
        )

# ---------------------------------------------------------------------------
# Area principal
# ---------------------------------------------------------------------------
st.divider()

if "cloud_is_running" not in st.session_state:
    st.session_state["cloud_is_running"] = False
if "cloud_last_run" not in st.session_state:
    st.session_state["cloud_last_run"] = None
if "cloud_accumulated_bytes" not in st.session_state:
    st.session_state["cloud_accumulated_bytes"] = None
if "cloud_session_total_added" not in st.session_state:
    st.session_state["cloud_session_total_added"] = 0
if "cloud_uploader_key" not in st.session_state:
    st.session_state["cloud_uploader_key"] = 0

with st.container(border=True):
    st.subheader("1️⃣ Subir Ordenes de Compra (OC) a procesar")
    uploaded_files = st.file_uploader(
        "Archivos de OC (PDF, Excel)",
        type=["pdf", "xlsx", "xlsm"],
        accept_multiple_files=True,
        label_visibility="collapsed",
        key=f"oc_uploader_{st.session_state['cloud_uploader_key']}",
    )
    if uploaded_files:
        st.caption(f"{len(uploaded_files)} archivo(s) seleccionado(s)")
    if st.session_state["cloud_last_run"] is not None:
        st.caption(
            "Puedes subir mas OC aqui y volver a darle a Ejecutar cuantas veces "
            "quieras -- se van acumulando en el mismo archivo dentro de esta sesion."
        )
    cached_count = st.session_state.get("cloud_extraction_cache_json")
    if cached_count:
        try:
            n_cached = len(json.loads(cached_count).get("entries", {}))
        except Exception:
            n_cached = 0
        st.caption(
            f"💾 {n_cached} archivo(s) ya leidos en esta sesion -- si los vuelves a "
            f"subir NO se le vuelve a cobrar a la API, se usa lo ya extraido."
        )

with st.container(border=True):
    st.subheader("2️⃣ Ejecutar")
    run_clicked = st.button(
        "🚀 Ejecutar consolidacion",
        type="primary",
        use_container_width=True,
        disabled=not uploaded_files or st.session_state["cloud_is_running"],
    )

    if run_clicked:
        st.session_state["cloud_is_running"] = True
        work_dir = None
        try:
            work_dir = Path(tempfile.mkdtemp(prefix="oc_cloud_"))
            ocs_folder = work_dir / OCS_SUBFOLDER
            ocs_folder.mkdir(parents=True, exist_ok=True)

            # El cache de extraccion (por contenido/MD5, ver main.py) evita
            # volver a pagarle a Claude por un archivo ya procesado -- pero
            # main.run() lo guarda dentro de esta carpeta temporal, que se
            # borra al terminar la corrida. Se restaura aqui desde la sesion
            # (si ya se corrio algo antes en esta misma sesion) para que
            # subir el mismo archivo dos veces en un rato no cueste doble.
            cached_json = st.session_state.get("cloud_extraction_cache_json")
            if cached_json:
                (ocs_folder / oc_main.CACHE_FILENAME).write_text(cached_json, encoding="utf-8")

            for filename, f in master_uploads.items():
                if f is not None:
                    (work_dir / filename).write_bytes(f.getvalue())

            if previous_output is not None:
                seed_bytes = previous_output.getvalue()
            else:
                seed_bytes = st.session_state.get("cloud_accumulated_bytes")
            if seed_bytes:
                (work_dir / output_filename).write_bytes(seed_bytes)

            oc_file_paths = []
            for f in uploaded_files:
                dest = ocs_folder / f.name
                dest.write_bytes(f.getvalue())
                oc_file_paths.append(dest)

            log_lines = []
            warnings = []
            progress_bar = st.progress(0.0)
            status = st.status("Procesando...", expanded=True)

            def _log(message):
                text = str(message)
                log_lines.append(text)
                status.write(text)

            def _progress(i, total):
                if total:
                    progress_bar.progress(i / total, text=f"{i}/{total}")

            try:
                with status:
                    added = oc_main.run(
                        str(work_dir), api_key,
                        log=_log,
                        progress=_progress,
                        oc_file_paths=oc_file_paths,
                        output_filename=output_filename,
                        warnings=warnings,
                        fake_mode=fake_mode,
                    )
                output_path = work_dir / output_filename
                output_bytes = output_path.read_bytes() if output_path.exists() else None
                cache_path = ocs_folder / oc_main.CACHE_FILENAME
                if cache_path.exists():
                    st.session_state["cloud_extraction_cache_json"] = cache_path.read_text(encoding="utf-8")
                status.update(label=f"Completado — {added} linea(s) nueva(s) agregada(s)", state="complete")
                st.session_state["cloud_session_total_added"] += added
                st.session_state["cloud_last_run"] = {
                    "timestamp": datetime.now().isoformat(timespec="seconds"),
                    "added": added,
                    "files_processed": len(oc_file_paths),
                    "warnings": warnings,
                    "log_lines": log_lines,
                    "output_bytes": output_bytes,
                    "output_filename": output_filename,
                }
                if output_bytes:
                    # Se guarda para que la proxima corrida en esta misma
                    # sesion siga acumulando sobre este resultado sin que el
                    # usuario tenga que descargar y volver a subir el
                    # archivo manualmente en la barra lateral.
                    st.session_state["cloud_accumulated_bytes"] = output_bytes
                st.session_state["cloud_uploader_key"] += 1
                st.session_state["cloud_is_running"] = False
                if work_dir is not None:
                    shutil.rmtree(work_dir, ignore_errors=True)
                    work_dir = None
                st.rerun()
            except Exception as e:
                status.update(label="Ocurrio un error", state="error")
                st.exception(e)
        finally:
            if work_dir is not None:
                shutil.rmtree(work_dir, ignore_errors=True)
            st.session_state["cloud_is_running"] = False

# ---------------------------------------------------------------------------
# Resultados
# ---------------------------------------------------------------------------
last_run = st.session_state.get("cloud_last_run")
if last_run:
    with st.container(border=True):
        st.subheader("3️⃣ Resultados")
        st.caption(f"Ultima ejecucion: {last_run['timestamp']}")

        col1, col2, col3 = st.columns(3)
        col1.metric("Lineas nuevas", last_run["added"])
        col2.metric("Archivos procesados", last_run["files_processed"])
        col3.metric("Items a revisar", len(last_run["warnings"]))

        grouped = {}
        for w in last_run["warnings"]:
            grouped.setdefault(w["type"], []).append(w)

        if grouped:
            with st.expander(f"⚠️ Revisar ({len(last_run['warnings'])})", expanded=True):
                for wtype, items in grouped.items():
                    st.markdown(f"**{WARNING_LABELS.get(wtype, wtype)}** ({len(items)})")
                    for item in items:
                        if wtype == "master_file_missing":
                            st.write(f"- {item['file']}")
                        elif wtype == "customer_not_found":
                            st.write(f"- OC {item['po_number']}: cliente '{item['customer_name']}' (RFC: {item['customer_rfc']})")
                        elif wtype == "code_uncertain":
                            st.write(f"- OC {item['po_number']}: codigo '{item['code']}' ({item['description']})")
                        elif wtype == "extraction_error":
                            st.write(f"- {item['file']}: {item['error']}")
                        else:
                            st.write(f"- {item}")
        else:
            st.success("No hay items que revisar.")

        with st.expander("Ver log completo"):
            st.code("\n".join(last_run["log_lines"]))

        if last_run["output_bytes"]:
            st.download_button(
                "⬇️ Descargar Excel",
                data=last_run["output_bytes"],
                file_name=last_run["output_filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
