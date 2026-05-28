"""
excel_reader.py
---------------
Lee calificaciones desde un Excel con una pestaña por trimestre.

Estructura esperada del Excel
------------------------------
- Una pestaña por período: "1er Trimestre", "2do Trimestre", "3er Trimestre",
  "Evaluacion Final" (los nombres deben coincidir con las claves de
  ``trimestres`` en config.json).
- Cada pestaña tiene una columna de alumnos y una columna por materia,
  cuyos nombres se mapean desde config.json.

Uso típico
----------
    reader = ExcelReader()
    reader.load_sheet("1er Trimestre")
    nota = reader.get_grade("GARCIA, LUCAS", "MATEMÁTICA")
"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Optional

import pandas as pd

logger = logging.getLogger(__name__)

CONFIG_PATH = Path(__file__).resolve().parents[1] / "config.json"


# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

def load_config(config_path: Path = CONFIG_PATH) -> dict:
    if not config_path.exists():
        raise FileNotFoundError(f"No se encontró config.json en: {config_path}")

    with config_path.open(encoding="utf-8") as f:
        config = json.load(f)

    required = {"excel_path", "columna_alumnos", "mapeo_materias", "trimestres"}
    missing = required - config.keys()
    if missing:
        raise KeyError(f"Faltan claves en config.json: {missing}")

    return config


# ---------------------------------------------------------------------------
# ExcelReader
# ---------------------------------------------------------------------------

class ExcelReader:
    """Lee notas desde un Excel con pestañas por trimestre."""

    def __init__(self, config_path: Path = CONFIG_PATH) -> None:
        self.config         = load_config(config_path)
        self.excel_path     = Path(self.config["excel_path"])
        self.student_col    = self.config["columna_alumnos"]
        self.subject_map    = self.config["mapeo_materias"]   # web_name → col_excel
        self.trimester_map  = self.config["trimestres"]        # web_name → sheet_name

        self.df: Optional[pd.DataFrame] = None
        self.active_sheet: Optional[str] = None

    # ------------------------------------------------------------------
    # Carga de pestaña
    # ------------------------------------------------------------------

    def load_sheet(self, trimester_web_name: str) -> None:
        """Carga la pestaña correspondiente al trimestre indicado.

        Args:
            trimester_web_name: Nombre del trimestre tal como aparece en la web
                                (p. ej. "1er Trimestre").

        Raises:
            ValueError: Si el trimestre no está mapeado en config.json o la
                        pestaña no existe en el Excel.
            FileNotFoundError: Si el archivo Excel no existe.
        """
        if not self.excel_path.exists():
            raise FileNotFoundError(f"Archivo Excel no encontrado: {self.excel_path}")

        if trimester_web_name not in self.trimester_map:
            raise ValueError(
                f"El trimestre '{trimester_web_name}' no está en config.json. "
                f"Disponibles: {list(self.trimester_map.keys())}"
            )

        sheet_name = self.trimester_map[trimester_web_name]
        available  = pd.ExcelFile(self.excel_path).sheet_names

        if sheet_name not in available:
            raise ValueError(
                f"La pestaña '{sheet_name}' no existe en el Excel. "
                f"Pestañas disponibles: {available}"
            )

        self.df = pd.read_excel(
            self.excel_path,
            sheet_name=sheet_name,
            dtype=str,
            keep_default_na=False,
        )
        self.df.columns        = [c.strip() for c in self.df.columns]
        self.df[self.student_col] = self.df[self.student_col].str.strip().str.upper()
        self.active_sheet = sheet_name

        logger.info("Pestaña '%s' cargada: %d alumnos.", sheet_name, len(self.df))

    # ------------------------------------------------------------------
    # Consulta de notas
    # ------------------------------------------------------------------

    def get_grade(self, student_name: str, web_subject_name: str) -> Optional[str]:
        """Devuelve la nota de un alumno para una materia.

        Args:
            student_name:     Nombre como aparece en la web (se normaliza internamente).
            web_subject_name: Nombre de la materia según la web (clave de mapeo_materias).

        Returns:
            La nota como string, o None si el alumno no tiene nota cargada.
        """
        self._require_sheet()

        excel_col = self._resolve_col(web_subject_name)
        name_norm = student_name.strip().upper()

        matches = self.df[self.df[self.student_col] == name_norm]

        if matches.empty:
            logger.warning("Alumno '%s' no encontrado en '%s'.", student_name, self.active_sheet)
            return None

        raw = matches.iloc[0][excel_col]
        if raw == "" or pd.isna(raw):
            return None

        return raw.strip()

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def list_students(self) -> list[str]:
        self._require_sheet()
        return self.df[self.student_col].tolist()

    def list_subjects(self) -> list[str]:
        return list(self.subject_map.keys())

    def list_trimesters(self) -> list[str]:
        return list(self.trimester_map.keys())

    def _require_sheet(self) -> None:
        if self.df is None:
            raise RuntimeError("Llamá a load_sheet() antes de consultar notas.")

    def _resolve_col(self, web_subject_name: str) -> str:
        if web_subject_name not in self.subject_map:
            raise KeyError(
                f"Materia '{web_subject_name}' sin mapeo. "
                f"Disponibles: {list(self.subject_map.keys())}"
            )
        col = self.subject_map[web_subject_name]
        if col not in self.df.columns:
            raise KeyError(
                f"Columna '{col}' no encontrada en '{self.active_sheet}'. "
                f"Columnas: {list(self.df.columns)}"
            )
        return col
