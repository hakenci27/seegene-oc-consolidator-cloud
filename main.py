"""
main.py
Orquesta el flujo completo:
  1. Carga los archivos maestros (customer master list, daily report, price list, CREDITO).
  2. Extrae cada OC subida con la API de Claude (con cache por contenido, para
     no repetir llamadas si el archivo ya se proceso antes). Toda OC que el
     usuario suba se procesa -- no se filtra contra Sales progress.xlsx.
  3. Cruza cada partida contra los archivos maestros.
  4. Actualiza/crea el archivo de salida (builder.py evita duplicar una OC
     que ya este en el archivo de salida por su PO No.).

Se puede correr por linea de comandos (para pruebas) o ser llamado desde
gui.py (Tkinter) o streamlit_app.py (Streamlit).
"""

import hashlib
import json
from dataclasses import asdict
from datetime import date
from pathlib import Path

from config import OCS_SUBFOLDER, OUTPUT_FILENAME, CUSTOMER_OVERRIDE_BY_PO_NUMBER
from extractor import extract_po, ExtractedPO, LineItem
from matcher import MasterData
from builder import build_or_update

CACHE_FILENAME = ".oc_extraction_cache.json"
CACHE_VERSION = 2


def _load_cache(ocs_folder: Path):
    """Cache con clave = hash MD5 del contenido del archivo (no depende de
    filename/mtime, para que funcione igual con archivos subidos por
    Streamlit que con archivos de una carpeta local)."""
    cache_path = ocs_folder / CACHE_FILENAME
    if cache_path.exists():
        try:
            data = json.loads(cache_path.read_text(encoding="utf-8"))
            if isinstance(data, dict) and data.get("version") == CACHE_VERSION:
                return data.get("entries", {})
        except Exception:
            pass
    return {}


def _save_cache(ocs_folder: Path, entries: dict):
    cache_path = ocs_folder / CACHE_FILENAME
    payload = {"version": CACHE_VERSION, "entries": entries}
    cache_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8")


def _extracted_from_dict(d):
    items = [LineItem(**li) for li in d["line_items"]]
    d2 = dict(d)
    d2["line_items"] = items
    return ExtractedPO(**d2)


def _file_hash(path: Path) -> str:
    return hashlib.md5(path.read_bytes()).hexdigest()


def _fake_extract_po(file_path):
    """Genera una OC ficticia sin llamar a la API de Claude -- para probar
    el flujo completo (subida, cruce contra maestros, generacion del Excel)
    sin necesitar una API key todavia. Los datos son inventados a proposito
    y quedan marcados como tales en 'notes' para que no se confundan con
    una extraccion real."""
    name = Path(file_path).name
    stem = Path(file_path).stem
    return ExtractedPO(
        source_file=file_path,
        po_number=f"TEST-{stem}"[:40],
        po_date=date.today().isoformat(),
        customer_name="Cliente de prueba (modo sin API)",
        customer_rfc=None,
        notes="DATOS FICTICIOS -- modo de prueba sin API activado, no se leyo el documento real.",
        line_items=[
            LineItem(
                code="TEST-001",
                description=f"Producto de prueba (basado en {name})",
                qty=1,
                unit_price=100.0,
                code_uncertain=False,
            ),
        ],
    )


def run(folder_str, api_key, log=print, progress=None,
        oc_file_paths=None, output_filename=None, warnings=None, fake_mode=False):
    """
    folder_str: carpeta de trabajo que contiene OCs/, Sales progress.xlsx,
                customer master list.xlsx, daily report.xlsx,
                190626 Official sales price list.xlsx, sales collection report.xlsx.
                Los archivos maestros que falten se degradan de forma segura
                (ver matcher.py) en vez de interrumpir el proceso.
    api_key: API key de Anthropic (para leer PDFs/Excels con Claude).
    log: funcion para mostrar mensajes de progreso (por defecto print; la GUI
         le pasa una funcion que escribe en el cuadro de texto).
    progress: funcion opcional progress(i, total) para una barra de progreso.
    oc_file_paths: lista opcional de Path a procesar en vez de escanear toda
                   la carpeta OCs/ -- lo usa streamlit_app.py para procesar
                   solo los archivos que el usuario acaba de subir/seleccionar
                   en esta corrida.
    output_filename: nombre de archivo de salida opcional (por defecto
                      config.OUTPUT_FILENAME).
    warnings: lista opcional donde se van agregando dicts estructurados
              (ademas de los mensajes de `log`) para que la interfaz pueda
              armar un panel de "revisar manualmente" sin tener que parsear
              texto: {"type": "master_file_missing"|"customer_not_found"|
              "code_uncertain"|"extraction_error", ...}.
    fake_mode: si es True, no se llama a la API de Claude ni se necesita
               una api_key valida -- cada archivo genera una OC ficticia
               (ver _fake_extract_po) para poder probar el resto del flujo
               (cruce contra maestros, generacion/actualizacion del Excel)
               sin costo. El cache de extraccion se ignora por completo en
               este modo para no mezclar datos ficticios con extracciones
               reales de una corrida posterior.
    """
    folder = Path(folder_str)
    ocs_folder = folder / OCS_SUBFOLDER
    output_path = folder / (output_filename or OUTPUT_FILENAME)
    if warnings is None:
        warnings = []

    log("Cargando archivos maestros...")
    master = MasterData(folder, log=log)
    for missing in master.missing_files:
        warnings.append({"type": "master_file_missing", "file": missing})

    if oc_file_paths is not None:
        files = sorted(Path(p) for p in oc_file_paths)
    else:
        files = sorted(
            p for p in ocs_folder.iterdir()
            if p.is_file() and p.suffix.lower() in (".pdf", ".xlsx", ".xlsm")
            and not p.name.startswith("~$")
        )
    log(f"Encontrados {len(files)} archivos para procesar")

    if fake_mode:
        log("MODO DE PRUEBA SIN API activado: se generaran datos ficticios en vez de leer los documentos.")

    cache = {} if fake_mode else _load_cache(ocs_folder)
    extracted_list = []
    for i, f in enumerate(files, start=1):
        if progress:
            progress(i, len(files))
        if fake_mode:
            log(f"[{i}/{len(files)}] {f.name}: generando datos ficticios (sin API)")
            extracted_list.append(_fake_extract_po(str(f)))
            continue
        content_hash = _file_hash(f)
        cached_entry = cache.get(content_hash)
        if cached_entry:
            log(f"[{i}/{len(files)}] {f.name}: usando extraccion en cache")
            extracted = _extracted_from_dict(cached_entry["data"])
        else:
            log(f"[{i}/{len(files)}] {f.name}: extrayendo con Claude...")
            try:
                extracted = extract_po(str(f), api_key, log=log)
            except Exception as e:
                log(f"  ERROR extrayendo {f.name}: {e}")
                warnings.append({"type": "extraction_error", "file": f.name, "error": str(e)})
                continue
            cache[content_hash] = {"source_name": f.name, "data": asdict(extracted)}
        extracted_list.append(extracted)

    if not fake_mode:
        _save_cache(ocs_folder, cache)

    new_pos = []
    for extracted in extracted_list:
        # No se filtra contra Sales progress.xlsx: toda OC que el usuario
        # suba explicitamente se procesa. El unico control de duplicados es
        # el de builder.py contra el PO No. ya presente en el archivo de
        # salida (para no duplicar si la misma OC se vuelve a subir).
        resolved_items = []
        # A veces el documento menciona a "Seegene Mexico" (el vendedor) mas
        # prominente que al cliente real (el que de verdad mando la OC) -- en
        # esos casos el nombre extraido queda mal sin importar el RFC, y la
        # unica forma confiable de identificarlo es por PO No. (ver
        # config.CUSTOMER_OVERRIDE_BY_PO_NUMBER, casos ya confirmados a mano).
        override_name = CUSTOMER_OVERRIDE_BY_PO_NUMBER.get(extracted.po_number.strip())
        if override_name:
            extracted.customer_name = override_name
        customer_record = master.find_customer(rfc=extracted.customer_rfc, name=extracted.customer_name)
        if customer_record:
            customer_code = customer_record["customer_code"]
            pic = customer_record["pic"] or "No disponible"
            customer_name_final = customer_record["customer_name"]
        else:
            customer_code = "No encontrado"
            pic = "No disponible"
            customer_name_final = extracted.customer_name
            log(f"  ATENCION: cliente '{extracted.customer_name}' (RFC: {extracted.customer_rfc}) "
                f"no encontrado en customer master list -- revisar manualmente.")
            warnings.append({
                "type": "customer_not_found",
                "po_number": extracted.po_number,
                "customer_name": extracted.customer_name,
                "customer_rfc": extracted.customer_rfc,
            })

        if extracted.notes:
            log(f"  NOTA (OC {extracted.po_number}): {extracted.notes}")

        for item in extracted.line_items:
            # Cruce contra Product master list: si el codigo extraido (incluso
            # combinado o con guiones, ej. "AD-BM-CN-SD9802X") hace match con
            # "Code" del maestro, se reemplazan codigo/descripcion/categoria
            # por los del maestro. Si el codigo no da match, se busca tambien
            # dentro de la descripcion (ej. "...25RX SD10245Z").
            product_match = master.find_product(item.code, item.description)
            if product_match:
                item.code = product_match["code"]
                item.description = product_match["description"] or item.description
                category = product_match["category"] or master.category_for_code(item.code)
                if not product_match.get("from_master", True):
                    # El codigo se saco de la descripcion (ej. "...*SD7700X")
                    # pero no esta en Product master list -- no se puede
                    # confirmar categoria/descripcion, se marca para revision.
                    item.code_uncertain = True
            else:
                category = master.category_for_code(item.code)
            if item.code_uncertain:
                log(f"  ATENCION: OC {extracted.po_number}, codigo '{item.code}' incierto "
                    f"-- confirmar con el documento original.")
                warnings.append({
                    "type": "code_uncertain",
                    "po_number": extracted.po_number,
                    "code": item.code,
                    "description": item.description,
                })
            # Q (Comments (SCM)) se deja siempre en blanco, sin texto y sin
            # color -- toda esta informacion (cliente no encontrado, codigo
            # incierto, notas de extraccion) ya queda en el log/warnings de
            # arriba.
            fields = {
                "pic": pic,
                "po_date": extracted.po_date,
                "customer_code": customer_code,
                "customer_name": customer_name_final,
                "category": category,
            }
            resolved_items.append((item, fields))
        new_pos.append((extracted, resolved_items))

    if not new_pos:
        log("No hay OC nuevas por agregar.")
        if not output_path.exists():
            log("Todo esta al dia (no hay archivo de salida existente para refrescar).")
            return 0
        # Aunque no haya OC nuevas, si el archivo de salida ya existe hay que
        # refrescar igual las columnas calculadas (V..AB) y el cruce contra
        # Product master list (G/H/I) de las filas ya existentes -- por
        # ejemplo, si el usuario agrego o actualizo Product master list
        # despues de que esas filas ya se habian creado.
        log("Refrescando columnas calculadas y cruce con Product master list de las filas existentes...")
        return build_or_update(output_path, [], master, log=log)

    log("Actualizando archivo de salida...")
    added = build_or_update(output_path, new_pos, master, log=log)
    log(f"Listo. {added} lineas nuevas agregadas en: {output_path}")
    return added


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        print("Uso: python main.py <carpeta_de_trabajo> <anthropic_api_key>")
        sys.exit(1)
    run(sys.argv[1], sys.argv[2])
