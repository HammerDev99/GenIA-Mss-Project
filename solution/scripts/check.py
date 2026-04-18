"""Validador de coherencia para los CSV y el XML generado.

Comprueba:
  1. Todas las predecesoras referencian IDs existentes.
  2. La tarea 1.2.4 (Desarrollo de código) es de tipo Duración fija.
  3. Existe la dependencia SS de 1.2.3 (Animaciones) -> 1.2.2 (Texturizado).
  4. Programador VR tiene capacidad máxima 300%.
  5. Las asignaciones referencian tareas y recursos existentes.
  6. Las fases suman las duraciones esperadas del BORRADOR.

Uso:
    python check.py
"""

import csv
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
XML_FILE = Path(__file__).resolve().parent.parent / "xml" / "proyecto-vr.xml"
NS = "{http://schemas.microsoft.com/project}"

EXPECTED_PHASE_DURATIONS = {
    "1.1": 29,  # 7 + 15 + 7
    "1.2": 80,  # 15 + 10 + 25 + 30
    "1.3": 34,  # 14 + 5 + 15
    "1.4": 10,  # 5 + 5
}


def load_csv(path: Path) -> list[dict]:
    with path.open(newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def check_predecessors(tareas: list[dict]) -> list[str]:
    errors = []
    ids = {int(t["id"]) for t in tareas}
    for t in tareas:
        if t["predecesora_id"]:
            pid = int(t["predecesora_id"])
            if pid not in ids:
                errors.append(f"  - Tarea {t['wbs']} {t['nombre']!r}: predecesora {pid} no existe")
    return errors


def check_fixed_duration_1_2_4(tareas: list[dict]) -> list[str]:
    for t in tareas:
        if t["wbs"] == "1.2.4":
            if t["tipo_tarea"] != "duracion_fija":
                return [f"  - 1.2.4 debe ser 'duracion_fija', es '{t['tipo_tarea']}'"]
            return []
    return ["  - Tarea 1.2.4 no encontrada"]


def check_ss_animaciones(tareas: list[dict]) -> list[str]:
    for t in tareas:
        if t["wbs"] == "1.2.3":
            if t["tipo_dep"] != "SS":
                return [f"  - 1.2.3 Animaciones debe tener dependencia SS, tiene '{t['tipo_dep']}'"]
            # Comprobar que la predecesora es 1.2.2
            pid = int(t["predecesora_id"])
            for tt in tareas:
                if int(tt["id"]) == pid:
                    if tt["wbs"] != "1.2.2":
                        return [f"  - 1.2.3 debería depender de 1.2.2, depende de {tt['wbs']}"]
                    return []
            return [f"  - Predecesora {pid} de 1.2.3 no encontrada"]
    return ["  - 1.2.3 Animaciones no encontrada"]


def check_programador_vr_300(recursos: list[dict]) -> list[str]:
    for r in recursos:
        if r["nombre"] == "Programador VR":
            if r["cap_max_pct"] != "300":
                return [f"  - Programador VR cap_max debe ser 300, es {r['cap_max_pct']}"]
            return []
    return ["  - Recurso 'Programador VR' no encontrado"]


def check_assignments_refs(
    asignaciones: list[dict],
    tareas: list[dict],
    recursos: list[dict],
) -> list[str]:
    errors = []
    tids = {int(t["id"]) for t in tareas}
    rids = {int(r["id"]) for r in recursos}
    for a in asignaciones:
        if int(a["tarea_id"]) not in tids:
            errors.append(f"  - Asignación tarea_id={a['tarea_id']} no existe")
        if int(a["recurso_id"]) not in rids:
            errors.append(f"  - Asignación recurso_id={a['recurso_id']} no existe")
    return errors


def check_phase_durations(tareas: list[dict]) -> list[str]:
    errors = []
    totals = {k: 0.0 for k in EXPECTED_PHASE_DURATIONS}
    for t in tareas:
        if t["es_resumen"].strip().lower() == "true":
            continue
        for phase in totals:
            if t["wbs"].startswith(phase + "."):
                totals[phase] += float(t["duracion_dias"])
    for phase, expected in EXPECTED_PHASE_DURATIONS.items():
        if abs(totals[phase] - expected) > 0.001:
            errors.append(
                f"  - Fase {phase}: suma {totals[phase]}d, esperado {expected}d"
            )
    return errors


def check_xml_festivos(festivos: list[dict]) -> list[str]:
    if not XML_FILE.exists():
        return [f"  - {XML_FILE} no existe (ejecuta build_xml.py primero)"]
    tree = ET.parse(XML_FILE)
    root = tree.getroot()
    exceptions = root.findall(f".//{NS}Calendar/{NS}Exceptions/{NS}Exception")
    fechas_xml = set()
    for exc in exceptions:
        fd = exc.find(f"{NS}TimePeriod/{NS}FromDate")
        if fd is not None and fd.text:
            fechas_xml.add(fd.text[:10])
    errors = []
    for f in festivos:
        if f["fecha"] not in fechas_xml:
            errors.append(f"  - Festivo {f['fecha']} {f['nombre']} no aparece en el XML")
    return errors


def main() -> int:
    tareas = load_csv(DATA_DIR / "tareas.csv")
    recursos = load_csv(DATA_DIR / "recursos.csv")
    asignaciones = load_csv(DATA_DIR / "asignaciones.csv")
    festivos = load_csv(DATA_DIR / "festivos.csv")

    checks = [
        ("Predecesoras referencian IDs existentes", check_predecessors(tareas)),
        ("1.2.4 Desarrollo código es Duración fija", check_fixed_duration_1_2_4(tareas)),
        ("1.2.3 Animaciones depende SS de 1.2.2", check_ss_animaciones(tareas)),
        ("Programador VR tiene MaxUnits=300%", check_programador_vr_300(recursos)),
        ("Asignaciones referencian IDs válidos", check_assignments_refs(asignaciones, tareas, recursos)),
        ("Duraciones por fase coinciden con BORRADOR", check_phase_durations(tareas)),
        ("Festivos colombianos presentes en XML", check_xml_festivos(festivos)),
    ]

    total_errors = 0
    for name, errs in checks:
        if errs:
            total_errors += len(errs)
            print(f"[FAIL] {name}")
            for e in errs:
                print(e)
        else:
            print(f"[OK]   {name}")

    if total_errors:
        print(f"\n{total_errors} error(es) encontrado(s).")
        return 1
    print("\nTodo coherente.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
