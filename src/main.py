"""main.py
---------
Orquestador: itera sobre trimestres y materias, selecciona los dropdowns
en la web y sincroniza las notas desde el Excel.

Uso
---
    # Probar una materia en un trimestre (recomendado para el primer test):
    python src/main.py "MATEMÁTICA" --trimestre "1er Trimestre"

    # Todas las materias en un trimestre:
    python src/main.py --trimestre "1er Trimestre"

    # Una materia en todos los trimestres:
    python src/main.py "MATEMÁTICA"

    # Todo (todas las materias, todos los trimestres):
    python src/main.py

Requisito previo
----------------
    brave-browser --remote-debugging-port=9222
Luego loguearse en SINIDE y navegar hasta Calificaciones. Después correr este script.
"""

from __future__ import annotations

import argparse
import logging
import sys
from typing import Optional, Sequence

from excel_reader import ExcelReader
from web_handler import WebHandler

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
)
logger = logging.getLogger(__name__)

# Solo estos trimestres tienen notas para cargar
TRIMESTRES_CARGABLES = [
    "1er Trimestre",
    "2do Trimestre",
    "3er Trimestre",
    "Evaluación Final",
    "Recup Diciembre",
    "Recup Febrero",
    "Evaluación Definitiva",
]


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Sincroniza calificaciones desde Excel hacia SINIDE.",
    )
    parser.add_argument(
        "subject",
        nargs="?",
        default=None,
        help="Nombre de la materia (ej: 'MATEMÁTICA'). Si se omite, se procesan todas.",
    )
    parser.add_argument(
        "--trimestre", "-t",
        dest="trimestre",
        default=None,
        help="Nombre del trimestre (ej: '1er Trimestre'). Si se omite, se procesan todos.",
    )
    return parser


# ---------------------------------------------------------------------------
# Lógica principal
# ---------------------------------------------------------------------------

def run(subject: Optional[str], trimestre: Optional[str]) -> int:
    reader = ExcelReader()

    subjects   = [subject]   if subject   else reader.list_subjects()
    trimestres = [trimestre] if trimestre else TRIMESTRES_CARGABLES

    # Validaciones previas
    for s in subjects:
        if s not in reader.list_subjects():
            logger.error("Materia '%s' no está en config.json. Disponibles: %s", s, reader.list_subjects())
            return 1

    for t in trimestres:
        if t not in reader.list_trimesters():
            logger.error("Trimestre '%s' no está en config.json. Disponibles: %s", t, reader.list_trimesters())
            return 1

    total_updated = 0

    with WebHandler() as wh:
        for trimestre_actual in trimestres:
            logger.info("=" * 50)
            logger.info("TRIMESTRE: %s", trimestre_actual)
            logger.info("=" * 50)

            # Cargar la pestaña del Excel para este trimestre
            try:
                reader.load_sheet(trimestre_actual)
            except (FileNotFoundError, ValueError) as exc:
                logger.error("No se pudo cargar el Excel para '%s': %s", trimestre_actual, exc)
                continue

            for materia_actual in subjects:
                logger.info("--- Materia: %s ---", materia_actual)

                # 1. Primero seleccionar la materia
                try:
                    wh.select_subject(materia_actual)
                except Exception as exc:
                    logger.error("No se pudo seleccionar materia '%s': %s", materia_actual, exc)
                    continue

                # 2. Luego seleccionar el trimestre
                try:
                    ok = wh.select_trimester_with_retry(trimestre_actual)
                    if not ok:
                        continue
                except Exception as exc:
                    logger.error("No se pudo seleccionar trimestre '%s': %s", trimestre_actual, exc)
                    continue

                # Sincronizar notas
                grade_resolver = lambda name, s=materia_actual: reader.get_grade(name, s)
                updated = wh.sync_grades(grade_resolver)
                total_updated += updated

                logger.info(
                    "✓ %s | %s: %d notas cargadas.",
                    trimestre_actual, materia_actual, updated,
                )

    logger.info("=" * 50)
    logger.info("FIN. Total de notas cargadas: %d", total_updated)
    logger.info("=" * 50)
    return 0


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        return run(subject=args.subject, trimestre=args.trimestre)
    except KeyboardInterrupt:
        logger.info("Interrumpido por el usuario.")
        return 0
    except Exception as exc:
        logger.exception("Error inesperado: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
