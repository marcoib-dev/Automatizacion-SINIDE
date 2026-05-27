"""main.py
---------
Punto de entrada para sincronizar calificaciones entre el Excel y SINIDE.

El script carga una pestaña del Excel, se conecta a una sesión ya abierta de
Brave mediante remote debugging y recorre los alumnos visibles en la tabla
web para escribir la nota correspondiente en cada fila.
"""

from __future__ import annotations

import argparse
import logging
from pathlib import Path
from typing import Sequence

import pandas as pd

from excel_reader import ExcelReader
from web_handler import WebHandler


logger = logging.getLogger(__name__)


def build_parser() -> argparse.ArgumentParser:
	"""Construye el parser de línea de comandos del script."""
	parser = argparse.ArgumentParser(
		description="Sincroniza calificaciones desde Excel hacia SINIDE.",
	)
	parser.add_argument(
		"subject",
		help="Nombre de la materia según config.json, por ejemplo 'CIENCIAS NATURALES'.",
	)
	parser.add_argument(
		"--sheet",
		dest="sheet_name",
		help="Nombre de la pestaña del Excel a cargar. Si se omite, se usa la primera hoja disponible.",
	)
	return parser


def get_available_sheets(excel_path: Path) -> list[str]:
	"""Devuelve las pestañas disponibles en el archivo Excel."""
	return pd.ExcelFile(excel_path).sheet_names


def resolve_sheet_name(requested_sheet: str | None, available_sheets: Sequence[str]) -> str:
	"""Elige la pestaña a cargar a partir de un valor pedido o el primer nombre disponible."""
	if not available_sheets:
		raise ValueError("El archivo Excel no contiene pestañas disponibles.")

	if requested_sheet:
		if requested_sheet not in available_sheets:
			raise ValueError(
				f"La pestaña '{requested_sheet}' no existe. Disponibles: {list(available_sheets)}"
			)
		return requested_sheet

	return available_sheets[0]


def run(subject: str, sheet_name: str | None = None) -> int:
	"""Ejecuta la sincronización completa de notas."""
	excel_reader = ExcelReader()

	available_sheets = get_available_sheets(excel_reader.excel_path)
	resolved_sheet = resolve_sheet_name(sheet_name, available_sheets)

	logger.info("Cargando pestaña '%s' del Excel.", resolved_sheet)
	excel_reader.load_sheet(resolved_sheet)

	if subject not in excel_reader.list_subjects():
		raise KeyError(
			f"La materia '{subject}' no está definida en config.json. "
			f"Materias disponibles: {excel_reader.list_subjects()}"
		)

	with WebHandler() as web_handler:
		web_students = web_handler.list_web_students()
		logger.info("Se detectaron %d alumnos visibles en la web.", len(web_students))

		updated_rows = 0

		for index, student_name in enumerate(web_students, start=1):
			logger.info("Procesando alumno %d/%d: %s", index, len(web_students), student_name)

			try:
				grade = excel_reader.get_grade(student_name, subject)
			except Exception:
				logger.exception("No se pudo resolver la nota para '%s'.", student_name)
				continue

			if grade is None or str(grade).strip() == "":
				logger.info("El alumno '%s' no tiene nota cargada; se omite.", student_name)
				continue

			if web_handler.fill_note_for_student(student_name, grade):
				updated_rows += 1
			else:
				logger.warning("No se pudo escribir la nota para '%s'.", student_name)

		logger.info(
			"Sincronización completada. Filas actualizadas: %d de %d.",
			updated_rows,
			len(web_students),
		)

	return 0


def main(argv: Sequence[str] | None = None) -> int:
	"""Punto de entrada CLI."""
	parser = build_parser()
	args = parser.parse_args(argv)

	try:
		return run(subject=args.subject, sheet_name=args.sheet_name)
	except Exception as exc:
		logger.exception("La sincronización terminó con error: %s", exc)
		return 1


if __name__ == "__main__":
	raise SystemExit(main())
