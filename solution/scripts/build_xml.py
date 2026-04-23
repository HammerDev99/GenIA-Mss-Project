"""Generador de Microsoft Project XML desde los CSV de solution/data.

Sigue el esquema oficial de Microsoft Project (Project.xsd, namespace
http://schemas.microsoft.com/project). El XML resultante puede abrirse
en MS Project Professional 2021 y guardarse como .mpp sin pérdida de datos.

Uso:
    python build_xml.py

Lee:
    solution/data/tareas.csv
    solution/data/recursos.csv
    solution/data/asignaciones.csv
    solution/data/festivos.csv

Escribe:
    solution/xml/proyecto-vr.xml
"""

import csv
import os
import xml.etree.ElementTree as ET
from datetime import date, datetime
from pathlib import Path

# ----- Constantes del proyecto (alineadas con BORRADOR.docx) ---------------

NS = "http://schemas.microsoft.com/project"
PROJECT_NAME = "Proyecto VR - Aplicación de Realidad Virtual"
PROJECT_TITLE = "Actividad 2 - Laboratorio MS Project"
PROJECT_AUTHOR = "Daniel Arbeláez Álvarez"
PROJECT_MANAGER = "Project Manager"
START_DATE = "2020-09-02T08:00:00"
CALENDAR_UID = 1
CALENDAR_NAME = "Calendario Colombia"

# Jornada: lun-vie 08:00-12:00 + 13:00-17:00 (8 h efectivas, 40 h/sem)
WORKING_TIMES = [("08:00:00", "12:00:00"), ("13:00:00", "17:00:00")]
MINUTES_PER_DAY = 480
MINUTES_PER_WEEK = 2400
DAYS_PER_MONTH = 20
HOURS_PER_DAY = 8

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
OUT_FILE = Path(__file__).resolve().parent.parent / "xml" / "proyecto-vr.xml"


# ----- Helpers de esquema --------------------------------------------------

TASK_TYPE = {"unidades_fijas": 0, "duracion_fija": 1, "trabajo_fijo": 2}
# Link types (esquema MS Project XML):
#   0 = FF, 1 = FS, 2 = SF, 3 = SS
LINK_TYPE = {"FF": 0, "FS": 1, "SF": 2, "SS": 3}
# DayType en <WeekDay>: 1=Dom, 2=Lun, 3=Mar, 4=Mié, 5=Jue, 6=Vie, 7=Sáb
WEEKDAYS_WORKING = [2, 3, 4, 5, 6]  # Lun a Vie

RESOURCE_TYPE = {"trabajo": 1, "material": 0, "costo": 2}


def iso_duration_from_days(dias: float) -> str:
    """Convierte días a duración ISO 8601 con jornada de 8h."""
    horas = int(round(dias * HOURS_PER_DAY))
    return f"PT{horas}H0M0S"


def sub(parent: ET.Element, tag: str, text=None) -> ET.Element:
    el = ET.SubElement(parent, tag)
    if text is not None:
        el.text = str(text)
    return el


# ----- Lectura de CSV ------------------------------------------------------

def read_csv(path: Path) -> list[dict]:
    with path.open(newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


# ----- Construcción del XML ------------------------------------------------

def build_calendars(root: ET.Element, festivos: list[dict]) -> None:
    calendars = sub(root, "Calendars")
    cal = sub(calendars, "Calendar")
    sub(cal, "UID", CALENDAR_UID)
    sub(cal, "Name", CALENDAR_NAME)
    sub(cal, "IsBaseCalendar", 1)
    sub(cal, "BaseCalendarUID", -1)

    weekdays = sub(cal, "WeekDays")
    for day in range(1, 8):
        wd = sub(weekdays, "WeekDay")
        sub(wd, "DayType", day)
        working = 1 if day in WEEKDAYS_WORKING else 0
        sub(wd, "DayWorking", working)
        if working:
            wt_parent = sub(wd, "WorkingTimes")
            for ft, tt in WORKING_TIMES:
                wt = sub(wt_parent, "WorkingTime")
                sub(wt, "FromTime", ft)
                sub(wt, "ToTime", tt)

    # Excepciones: festivos como día no laborable
    if festivos:
        exc_parent = sub(cal, "Exceptions")
        for idx, fest in enumerate(festivos, start=1):
            fecha = fest["fecha"]
            exc = sub(exc_parent, "Exception")
            sub(exc, "Enabled", 1)
            sub(exc, "Name", fest["nombre"])
            tr = sub(exc, "TimePeriod")
            sub(tr, "FromDate", f"{fecha}T00:00:00")
            sub(tr, "ToDate", f"{fecha}T23:59:00")
            sub(exc, "Type", 0)  # 0 = Once
            sub(exc, "Occurrences", 1)
            sub(exc, "DaysOfWeek", 0)
            sub(exc, "DayWorking", 0)  # día no laborable
            sub(exc, "UID", idx)


def build_tasks(root: ET.Element, tareas: list[dict]) -> None:
    """Emite <Task> en el orden CANÓNICO del esquema MS Project XML.

    El esquema usa xs:sequence: si los elementos no están en orden, MS
    Project descarta silenciosamente los que llegan fuera de turno (en
    particular PredecessorLink, lo que rompía el encadenamiento).

    Incluimos también los elementos "housekeeping" que Project espera
    ver (Stop, Resume, EffortDriven, Milestone, PercentComplete, etc.)
    sin los cuales su motor de scheduling no recalcula bien.
    """
    tasks_el = sub(root, "Tasks")
    for t in tareas:
        task_uid = int(t["id"])
        is_summary = t["es_resumen"].strip().lower() == "true"
        task_el = sub(tasks_el, "Task")
        sub(task_el, "UID", task_uid)
        sub(task_el, "ID", task_uid)
        sub(task_el, "Name", t["nombre"])
        sub(task_el, "Type", TASK_TYPE[t["tipo_tarea"]])
        sub(task_el, "IsNull", 0)
        sub(task_el, "CreateDate", START_DATE)
        sub(task_el, "WBS", t["wbs"])
        sub(task_el, "OutlineNumber", t["wbs"])
        sub(task_el, "OutlineLevel", t["nivel"])
        sub(task_el, "Priority", 500)
        # Start/Finish: se emiten como hints. Con ConstraintType=0 (ASAP)
        # Project los trata como valores iniciales y los recalcula según
        # predecesoras + calendario. Sin Start/Finish en el XML, Project
        # no arranca el motor de autoschedule y deja todo en parallel.
        sub(task_el, "Start", START_DATE)
        sub(task_el, "Finish", START_DATE)
        if not is_summary and t["duracion_dias"]:
            dur = iso_duration_from_days(float(t["duracion_dias"]))
            sub(task_el, "Duration", dur)
            sub(task_el, "DurationFormat", 7)  # 7 = días
            # NOTA: no emitimos <Work> a nivel <Task>. Project lo calcula
            # como suma del trabajo de las asignaciones; emitirlo aquí
            # creaba conflicto con assignment.Work y rompía Duración fija.
        sub(task_el, "Stop", "NA")
        sub(task_el, "Resume", "NA")
        sub(task_el, "ResumeValid", 0)
        # EffortDriven=0: no reducir duración al agregar recursos.
        # Clave para Prototipo (3 recursos al 100%) y para que 1.2.4
        # Duración fija mantenga sus 30 días con Programador VR[300%].
        sub(task_el, "EffortDriven", 0)
        sub(task_el, "Recurring", 0)
        sub(task_el, "OverAllocated", 0)
        sub(task_el, "Estimated", 0)
        sub(task_el, "Milestone", 0)
        sub(task_el, "Summary", 1 if is_summary else 0)
        sub(task_el, "Critical", 0)
        sub(task_el, "IsSubproject", 0)
        sub(task_el, "IsSubprojectReadOnly", 0)
        sub(task_el, "ExternalTask", 0)
        sub(task_el, "FixedCost", 0)
        sub(task_el, "FixedCostAccrual", 3)
        sub(task_el, "PercentComplete", 0)
        sub(task_el, "PercentWorkComplete", 0)
        sub(task_el, "Cost", 0)
        sub(task_el, "OvertimeCost", 0)
        sub(task_el, "OvertimeWork", "PT0H0M0S")
        # ConstraintType=0 -> "Lo antes posible" (ASAP).
        sub(task_el, "ConstraintType", 0)
        sub(task_el, "CalendarUID", -1)  # -1 = usa calendario del proyecto
        sub(task_el, "LevelAssignments", 1)
        sub(task_el, "LevelingCanSplit", 1)
        sub(task_el, "LevelingDelay", 0)
        sub(task_el, "LevelingDelayFormat", 7)
        sub(task_el, "HideBar", 0)
        sub(task_el, "Rollup", 0)
        if t["predecesora_id"]:
            link = sub(task_el, "PredecessorLink")
            sub(link, "PredecessorUID", int(t["predecesora_id"]))
            sub(link, "Type", LINK_TYPE[t["tipo_dep"]])
            lag_dias = float(t["lag_dias"] or 0)
            lag_min = int(round(lag_dias * MINUTES_PER_DAY))
            sub(link, "LinkLag", lag_min * 10)  # décimas de minuto
            sub(link, "LagFormat", 7)  # días
            sub(link, "CrossProject", 0)
        sub(task_el, "Active", 1)
        sub(task_el, "Manual", 0)  # 0 = autoprogramado


def build_resources(root: ET.Element, recursos: list[dict]) -> None:
    res_parent = sub(root, "Resources")
    for r in recursos:
        uid = int(r["id"])
        tipo = r["tipo"].strip().lower()
        res = sub(res_parent, "Resource")
        sub(res, "UID", uid)
        sub(res, "ID", uid)
        sub(res, "Name", r["nombre"])
        sub(res, "Type", RESOURCE_TYPE[tipo])
        sub(res, "IsNull", 0)
        sub(res, "Initials", r["nombre"][:3].upper())
        if tipo == "trabajo":
            cap = float(r["cap_max_pct"]) / 100.0
            sub(res, "MaxUnits", f"{cap:.6f}")
            sub(res, "StandardRate", float(r["tasa_estandar_eur_h"]))
            sub(res, "StandardRateFormat", 2)  # 2 = por hora
            sub(res, "OvertimeRate", float(r["tasa_extra_eur_h"]))
            sub(res, "OvertimeRateFormat", 2)
            sub(res, "CalendarUID", CALENDAR_UID)
        elif tipo == "material":
            sub(res, "MaterialLabel", "uso")
            if r["costo_uso_eur"]:
                sub(res, "CostPerUse", float(r["costo_uso_eur"]))


def build_assignments(
    root: ET.Element,
    asignaciones: list[dict],
    tareas_by_id: dict[int, dict],
    recursos_by_id: dict[int, dict],
) -> None:
    asg_parent = sub(root, "Assignments")
    for idx, a in enumerate(asignaciones, start=1):
        tarea_id = int(a["tarea_id"])
        recurso_id = int(a["recurso_id"])
        tarea = tareas_by_id[tarea_id]
        recurso = recursos_by_id[recurso_id]
        es_material = recurso["tipo"].strip().lower() == "material"
        unidades_raw = a["unidades_pct"].strip() if a["unidades_pct"] else ""

        if es_material:
            # Para material, Units = cantidad consumida (1 uso por defecto).
            # No se emite Work en horas — Project no lo espera para material.
            units = float(unidades_raw) if unidades_raw else 1.0
            horas = 0.0
        else:
            units = float(unidades_raw) / 100.0 if unidades_raw else 1.0
            duracion = float(tarea["duracion_dias"]) if tarea["duracion_dias"] else 0
            horas = duracion * HOURS_PER_DAY * units

        asg = sub(asg_parent, "Assignment")
        sub(asg, "UID", idx)
        sub(asg, "TaskUID", tarea_id)
        sub(asg, "ResourceUID", recurso_id)
        sub(asg, "Units", f"{units:.6f}")
        if horas > 0:
            sub(asg, "Work", f"PT{int(round(horas))}H0M0S")


def build_project(tareas, recursos, asignaciones, festivos) -> ET.ElementTree:
    root = ET.Element("Project", xmlns=NS)
    sub(root, "Name", PROJECT_NAME)
    sub(root, "Title", PROJECT_TITLE)
    sub(root, "Author", PROJECT_AUTHOR)
    sub(root, "Manager", PROJECT_MANAGER)
    sub(root, "ScheduleFromStart", 1)
    sub(root, "StartDate", START_DATE)
    # NO emitir <FinishDate>. Si se emite igual a StartDate, Project lo
    # trata como fecha fin obligatoria del proyecto y comprime/solapa
    # todas las tareas (Desarrollo cae a 1 día, Programador VR[9000%]).
    sub(root, "FYStartDate", 1)
    sub(root, "CriticalSlackLimit", 0)
    sub(root, "CurrencyDigits", 2)
    sub(root, "CurrencySymbol", "€")
    sub(root, "CurrencySymbolPosition", 3)
    sub(root, "CalendarUID", CALENDAR_UID)
    sub(root, "DefaultStartTime", "08:00:00")
    sub(root, "DefaultFinishTime", "17:00:00")
    sub(root, "MinutesPerDay", MINUTES_PER_DAY)
    sub(root, "MinutesPerWeek", MINUTES_PER_WEEK)
    sub(root, "DaysPerMonth", DAYS_PER_MONTH)
    sub(root, "DefaultTaskType", 0)
    sub(root, "DefaultFixedCostAccrual", 3)
    sub(root, "DefaultStandardRate", 0)
    sub(root, "DefaultOvertimeRate", 0)
    sub(root, "DurationFormat", 7)
    sub(root, "WorkFormat", 2)
    sub(root, "SpreadActualCost", 0)
    sub(root, "SpreadPercentComplete", 0)
    sub(root, "TaskUpdatesResource", 1)
    sub(root, "FiscalYearStart", 0)
    sub(root, "WeekStartDay", 1)
    sub(root, "NewTasksAreManual", 0)  # clave: autoprogramado
    sub(root, "ShowProjectSummaryTask", 1)

    build_calendars(root, festivos)
    build_tasks(root, tareas)
    build_resources(root, recursos)
    tareas_by_id = {int(t["id"]): t for t in tareas}
    recursos_by_id = {int(r["id"]): r for r in recursos}
    build_assignments(root, asignaciones, tareas_by_id, recursos_by_id)

    return ET.ElementTree(root)


def main() -> None:
    tareas = read_csv(DATA_DIR / "tareas.csv")
    recursos = read_csv(DATA_DIR / "recursos.csv")
    asignaciones = read_csv(DATA_DIR / "asignaciones.csv")
    festivos = read_csv(DATA_DIR / "festivos.csv")

    tree = build_project(tareas, recursos, asignaciones, festivos)
    ET.indent(tree, space="  ")
    OUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    tree.write(OUT_FILE, encoding="UTF-8", xml_declaration=True)
    print(f"OK - {OUT_FILE} ({OUT_FILE.stat().st_size} bytes)")
    print(f"  Tareas: {len(tareas)}")
    print(f"  Recursos: {len(recursos)}")
    print(f"  Asignaciones: {len(asignaciones)}")
    print(f"  Festivos: {len(festivos)}")


if __name__ == "__main__":
    main()
