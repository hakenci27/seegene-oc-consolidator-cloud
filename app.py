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
    "code_uncertain": "Codigo de producto incierto (revisar documento original)",
    "extraction_error": "Error al leer el documento",
}


def get_api_key():
    try:
        return st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        return None


st.set_page_config(page_title="Consolidador de Nuevas OC (Cloud)", page_icon="☁️", layout="wide")

st.title("☁️ Consolidador de Nuevas OC — Cloud")
st.caption(
    "팀원 누구나 이 링크에서 사용할 수 있는 클라우드 버전입니다. 로컬 버전과 동일한 "
    "엔진을 사용해서 고객사/제품 코드 자동 매칭, 신용한도 체크까지 그대로 동작합니다 — "
    "다만 이 서버는 회사 PC의 파일에 직접 접근할 수 없기 때문에, 마스터 파일들을 매번 "
    "직접 업로드해주셔야 합니다."
)

api_key = get_api_key()
if not api_key:
    st.error(
        "Claude API 키가 설정되지 않았습니다. `.streamlit/secrets.toml`에 "
        "`ANTHROPIC_API_KEY`를 넣거나(로컬), Streamlit Cloud의 앱 설정 > Secrets에 "
        "추가해주세요. (README.md 참고)"
    )
    st.stop()

# ---------------------------------------------------------------------------
# 사이드바: 마스터 파일 업로드 + 설정
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("## ⚙️ 설정")

    with st.container(border=True):
        st.markdown("**🗂️ 마스터 파일 업로드**")
        st.caption(
            "없어도 실행은 되지만(해당 기능만 빠짐), 있으면 고객사/제품 코드가 "
            "자동으로 매칭됩니다. 매일 같은 파일을 올리는 게 번거로우면, 본인 PC에 "
            "이 파일들을 따로 모아두고 매번 드래그해서 올리시는 걸 추천드립니다."
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
            st.success(f"{found} / {len(MASTER_FILES)} 업로드됨")
        else:
            st.warning(f"{found} / {len(MASTER_FILES)} 업로드됨")

    with st.container(border=True):
        st.markdown("**📄 기존 결과 파일 (선택)**")
        st.caption(
            "어제까지 누적된 결과 엑셀이 있으면 여기에 올려주세요 — 그 파일 뒤에 "
            "새 OC만 이어붙입니다 (중복 방지, 손으로 채운 컬럼 보존). 없으면 새로 "
            "만듭니다."
        )
        previous_output = st.file_uploader(
            "기존 Nuevas OC - Control...xlsx", type=["xlsx"], key="previous_output",
            label_visibility="collapsed",
        )

    with st.container(border=True):
        st.markdown("**💾 출력 파일명**")
        output_filename = st.text_input(
            "출력 파일명", value=DEFAULT_OUTPUT_FILENAME, label_visibility="collapsed",
        )

# ---------------------------------------------------------------------------
# 메인 영역
# ---------------------------------------------------------------------------
st.divider()

with st.container(border=True):
    st.subheader("1️⃣ 처리할 구매주문서(OC) 업로드")
    uploaded_files = st.file_uploader(
        "OC 파일 (PDF, Excel)",
        type=["pdf", "xlsx", "xlsm"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )
    if uploaded_files:
        st.caption(f"{len(uploaded_files)}개 파일 선택됨")

if "cloud_is_running" not in st.session_state:
    st.session_state["cloud_is_running"] = False
if "cloud_last_run" not in st.session_state:
    st.session_state["cloud_last_run"] = None

with st.container(border=True):
    st.subheader("2️⃣ 실행")
    run_clicked = st.button(
        "🚀 취합 실행",
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

            for filename, f in master_uploads.items():
                if f is not None:
                    (work_dir / filename).write_bytes(f.getvalue())

            if previous_output is not None:
                (work_dir / output_filename).write_bytes(previous_output.getvalue())

            oc_file_paths = []
            for f in uploaded_files:
                dest = ocs_folder / f.name
                dest.write_bytes(f.getvalue())
                oc_file_paths.append(dest)

            log_lines = []
            warnings = []
            progress_bar = st.progress(0.0)
            status = st.status("처리 중...", expanded=True)

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
                    )
                output_path = work_dir / output_filename
                output_bytes = output_path.read_bytes() if output_path.exists() else None
                status.update(label=f"완료 — {added}건 신규 추가", state="complete")
                st.session_state["cloud_last_run"] = {
                    "timestamp": datetime.now().isoformat(timespec="seconds"),
                    "added": added,
                    "files_processed": len(oc_file_paths),
                    "warnings": warnings,
                    "log_lines": log_lines,
                    "output_bytes": output_bytes,
                    "output_filename": output_filename,
                }
            except Exception as e:
                status.update(label="오류 발생", state="error")
                st.exception(e)
        finally:
            if work_dir is not None:
                shutil.rmtree(work_dir, ignore_errors=True)
            st.session_state["cloud_is_running"] = False

# ---------------------------------------------------------------------------
# 결과
# ---------------------------------------------------------------------------
last_run = st.session_state.get("cloud_last_run")
if last_run:
    with st.container(border=True):
        st.subheader("3️⃣ 결과")
        st.caption(f"마지막 실행: {last_run['timestamp']}")

        col1, col2, col3 = st.columns(3)
        col1.metric("신규 라인", last_run["added"])
        col2.metric("처리한 파일", last_run["files_processed"])
        col3.metric("확인 필요 항목", len(last_run["warnings"]))

        grouped = {}
        for w in last_run["warnings"]:
            grouped.setdefault(w["type"], []).append(w)

        if grouped:
            with st.expander(f"⚠️ 확인 필요 ({len(last_run['warnings'])})", expanded=True):
                for wtype, items in grouped.items():
                    st.markdown(f"**{WARNING_LABELS.get(wtype, wtype)}** ({len(items)})")
                    for item in items:
                        if wtype == "master_file_missing":
                            st.write(f"- {item['file']}")
                        elif wtype == "customer_not_found":
                            st.write(f"- OC {item['po_number']}: 고객 '{item['customer_name']}' (RFC: {item['customer_rfc']})")
                        elif wtype == "code_uncertain":
                            st.write(f"- OC {item['po_number']}: 코드 '{item['code']}' ({item['description']})")
                        elif wtype == "extraction_error":
                            st.write(f"- {item['file']}: {item['error']}")
                        else:
                            st.write(f"- {item}")
        else:
            st.success("확인이 필요한 항목이 없습니다.")

        with st.expander("전체 로그 보기"):
            st.code("\n".join(last_run["log_lines"]))

        if last_run["output_bytes"]:
            st.download_button(
                "⬇️ 엑셀 다운로드",
                data=last_run["output_bytes"],
                file_name=last_run["output_filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
