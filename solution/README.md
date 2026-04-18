# Solución Actividad 2 — Laboratorio MS Project

Pipeline IA → XML nativo de Microsoft Project → `.mpp`.

## Estructura

```
solution/
├── data/              # Fuente única de verdad (CSV legibles)
│   ├── tareas.csv
│   ├── recursos.csv
│   ├── asignaciones.csv
│   └── festivos.csv
├── scripts/
│   ├── build_xml.py   # Genera el XML nativo de MS Project
│   └── check.py       # Valida coherencia de CSV + XML
├── xml/
│   └── proyecto-vr.xml  # Salida para abrir en MS Project
└── capturas/          # Figuras 1-18 del informe (se llenan tras operar en Project)
```

## Rutas soportadas

Quedan solo dos rutas del pipeline investigado:

- **Ruta A (principal) — XML nativo.** `proyecto-vr.xml` transporta EDT + calendario Colombia + recursos + asignaciones. Se abre en MS Project Pro 2021 y se guarda como `.mpp`.
- **Ruta C (respaldo) — CSV/Excel.** Si el XML fallara, los CSV de `data/` sirven como fuente para el asistente oficial *Importar datos de Excel a Project* (ref: https://support.microsoft.com/es-es/topic/importar-datos-de-excel-a-project-cb3fb91a-ad05-4506-b0af-8aa8b2247119).

## Cómo compilar

```bash
python3 solution/scripts/build_xml.py
python3 solution/scripts/check.py
```

`check.py` debe imprimir 7 checks en `[OK]`.

## Flujo en MS Project (Windows 11)

### Paso 0 — Abrir

1. `Archivo → Abrir` → seleccionar `solution/xml/proyecto-vr.xml`.
2. Project pregunta el formato; aceptar "XML de Microsoft Project".
3. `Archivo → Guardar como` → `proyecto-vr.mpp`.

Al abrir, Project debe mostrar:

- 12 tareas + 4 resúmenes de fase + resumen raíz del proyecto.
- Fecha de inicio `02/09/2020`.
- Calendario Colombia activo en `Proyecto → Información del proyecto`.
- Jornada partida 08:00‑12:00 y 13:00‑17:00 en `Proyecto → Cambiar tiempo de trabajo`.
- Dependencia SS entre 1.2.3 Animaciones y 1.2.2 Texturizado.

Los **Pasos 1‑4** del laboratorio quedan cubiertos por el XML. Los pasos siguientes se ejecutan en la UI.

### Paso 5 — Línea base

`Proyecto → Establecer línea base → Línea base → Proyecto completo → Aceptar`.

**Captura Figura 7:** Gantt con barras grises de línea base bajo las azules.

### Paso 6 — Fechas sobre las barras

`Formato → Estilos de barra` (o doble clic sobre una barra) → pestaña `Texto` → `Izquierda = Comienzo`, `Derecha = Fin`.

**Captura Figura 8.**

### Paso 7 — Tablas personalizadas

`Vista → Tablas → Más tablas → Nueva`. Crear las cuatro:

| Tabla | Columnas |
|---|---|
| `Tabla-General-Proyecto-VR` | Nombre, Duración, Comienzo, Fin, Predecesoras, Nombres de los recursos, Indicadores |
| `Tabla-Costes-Proyecto-VR` | Nombre, Costo previsto, Costo real, Variación de costo |
| `Tabla-Trabajo-Proyecto-VR` | Nombre, Trabajo previsto, Trabajo real, Variación de trabajo, % trabajo completado |
| `Tabla-Cronograma-Proyecto-VR` | Nombre, Duración prevista, Duración real, Variación de duración, % completado, Comienzo, Fin |

**Capturas Figura 9, 10, 11, 12.**

### Paso 8 — Gestión del cambio (sobreasignación)

El XML ya incluye en el CSV `asignaciones.csv` la asignación base de `Modelador y texturizador 3D` a 1.2.2 Texturizado. Para simular el cambio:

1. En `Vista → Diagrama de Gantt`, columna `Nombres de recursos` de la tarea `1.2.3 Animaciones`, añadir `Modelador y texturizador 3D[50%]`.
2. Verificar en `Tabla-General-Proyecto-VR` el indicador rojo de sobreasignación sobre el Modelador.
   **Captura Figura 13.**
3. `Vista → Uso de recursos` → localizar Modelador → en la fila de la tarea 1.2.3, redistribuir las 100 h entre las fechas posteriores al fin de 1.2.2. Con Calendario Madrid eran 18‑nov a 08‑dic; con Calendario Colombia las fechas pueden variar por los festivos 02‑nov, 16‑nov, 08‑dic — anotar las fechas reales que Project calcule y corregir el texto del informe en el paso 8 del BORRADOR.
   **Captura Figura 14.**
4. Actualizar columna `Trabajo previsto` a 300 h en la tarea 1.2.3.
   **Captura Figura 15.**
5. Volver a establecer línea base.

### Paso 9 — Avance y horas extra

1. Marcar `% completado = 100` en todas las tareas anteriores a `1.2.4 Desarrollo de código`.
2. `Proyecto → Fecha de estado` → `25/01/2021`.
   **Captura Figura 16.**
3. En la tarea `1.2.4`: doble clic → formulario de tarea → `Trabajo real = 724h`, `Trabajo horas extra = 4h`.
4. Comprobar la variación de costo en `Tabla-Costes-Proyecto-VR` — valor esperado ≈ `4h × (17 − 15) €/h = 8 €`.
   **Captura Figura 17.**
5. `Vista → Uso de recursos` → localizar Programador VR en 1.2.4 → colocar 2 h extra el 21/01 y 2 h el 22/01, eliminar el reparto automático.
   **Captura Figura 18.**

## Placeholders del informe a resolver

Tras completar los pasos 5‑9, abrir `docs/Actividad2_Laboratorio_MS_Project_BORRADOR.docx` y sustituir:

| Placeholder | Valor |
|---|---|
| `[INSERTAR FECHA CALCULADA POR MS PROJECT]` (paso 2) | Fecha fin que Project muestra en `Proyecto → Información del proyecto`. |
| `[INSERTAR VALOR DE LA TABLA-COSTES-PROYECTO-VR]` (paso 9) | Variación de costo real que aparece en la tabla. |

Además, reemplazar en el texto del informe **Calendario Madrid → Calendario Colombia** y la lista de festivos por la de `data/festivos.csv`, ya que se decidió usar festivos colombianos.

## Notas sobre asignaciones

`data/asignaciones.csv` contiene asignaciones razonables por el nombre de cada tarea; el BORRADOR menciona una tabla de asignaciones del enunciado que no está incluida en el documento. Si el enunciado original difiere, basta con editar el CSV y volver a ejecutar `build_xml.py`.
