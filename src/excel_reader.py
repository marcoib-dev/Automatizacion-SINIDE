"""
excel_reader.py
---------------
Módulo para la lectura y consulta de calificaciones desde un archivo Excel.

Carga la configuración desde `config.json` (rutas, columna de alumnos y
mapeo de materias) y expone la clase `ExcelReader`, que permite:
  - Cargar una pestaña específica del Excel.
  - Consultar la nota de un alumno para una materia dada (por nombre de web).

Dependencias externas: pandas, openpyxl
"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Optional

import pandas as pd

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logger = logging.getLogger(__name__)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
)

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------
CONFIG_PATH = Path(__file__).resolve().parents[1] / "config.json"


# ---------------------------------------------------------------------------
# Carga de configuración
# ---------------------------------------------------------------------------

def load_config(config_path: Path = CONFIG_PATH) -> dict:
    """Carga y valida el archivo de configuración JSON.

    El JSON debe tener, como mínimo, la siguiente estructura::

        {
            "excel_path": "data/notas_colegio.xlsx",
            "columna_alumnos": "Nombre del Alumno",
            "mapeo_materias": {
                "Matemática": "MATEMATICA",
                "Lengua":     "LENGUA"
            }
        }

    Args:
        config_path: Ruta al archivo ``config.json``.
                     Por defecto apunta a la raíz del proyecto.

    Returns:
        Diccionario con la configuración cargada.

    Raises:
        FileNotFoundError: Si ``config_path`` no existe.
        KeyError: Si faltan claves obligatorias en el JSON.
        json.JSONDecodeError: Si el archivo no es JSON válido.
    """
    if not config_path.exists():
        raise FileNotFoundError(f"No se encontró el archivo de configuración: {config_path}")

    with config_path.open(encoding="utf-8") as f:
        config = json.load(f)

    # Validación de claves obligatorias
    required_keys = {"excel_path", "columna_alumnos", "mapeo_materias"}
    missing = required_keys - config.keys()
    if missing:
        raise KeyError(f"Faltan claves obligatorias en config.json: {missing}")

    logger.info("Configuración cargada correctamente desde '%s'.", config_path)
    return config


# ---------------------------------------------------------------------------
# Clase principal
# ---------------------------------------------------------------------------

class ExcelReader:
    """Lee y consulta calificaciones desde un archivo Excel multi-pestaña.

    Example:
        >>> reader = ExcelReader()
        >>> reader.load_sheet("1er Trimestre")
        >>> nota = reader.get_grade("GARCIA, LUCAS", "Matemática")
        >>> print(nota)  # → 8.5

    Attributes:
        config (dict): Configuración cargada desde ``config.json``.
        excel_path (Path): Ruta al archivo Excel.
        student_col (str): Nombre de la columna que identifica a los alumnos.
        subject_mapping (dict[str, str]): Mapeo ``nombre_web → columna_excel``.
        df (pd.DataFrame | None): DataFrame de la pestaña actualmente cargada.
        active_sheet (str | None): Nombre de la pestaña en memoria.
    """

    def __init__(self, config_path: Path = CONFIG_PATH) -> None:
        """Inicializa el lector cargando la configuración.

        Args:
            config_path: Ruta al archivo ``config.json``.
        """
        self.config = load_config(config_path)

        self.excel_path: Path = Path(self.config["excel_path"])
        self.student_col: str = self.config["columna_alumnos"]
        self.subject_mapping: dict[str, str] = self.config["mapeo_materias"]

        # Estado interno — se completan al llamar a `load_sheet`
        self.df: Optional[pd.DataFrame] = None
        self.active_sheet: Optional[str] = None

    # ------------------------------------------------------------------
    # Métodos públicos
    # ------------------------------------------------------------------

    def load_sheet(self, sheet_name: str) -> None:
        """Carga una pestaña del Excel en memoria como DataFrame.

        Normaliza los nombres de columnas (elimina espacios sobrantes) y
        limpia los valores de la columna de alumnos para facilitar la búsqueda.

        Args:
            sheet_name: Nombre exacto de la pestaña a cargar
                        (p. ej. ``"1er Trimestre"``).

        Raises:
            FileNotFoundError: Si el archivo Excel no existe en la ruta configurada.
            ValueError: Si la pestaña ``sheet_name`` no existe en el Excel.
            KeyError: Si la columna de alumnos no se encuentra en la pestaña.
        """
        if not self.excel_path.exists():
            raise FileNotFoundError(
                f"No se encontró el archivo Excel: '{self.excel_path}'. "
                "Verificá la clave 'excel_path' en config.json."
            )

        # Verificar que la pestaña exista antes de cargar todo el archivo
        available_sheets = pd.ExcelFile(self.excel_path).sheet_names
        if sheet_name not in available_sheets:
            raise ValueError(
                f"La pestaña '{sheet_name}' no existe en '{self.excel_path}'. "
                f"Pestañas disponibles: {available_sheets}"
            )

        self.df = pd.read_excel(
            self.excel_path,
            sheet_name=sheet_name,
            dtype=str,          # Leer todo como texto para preservar formatos
            keep_default_na=False,
        )

        # Normalización: eliminar espacios en nombres de columnas
        self.df.columns = [col.strip() for col in self.df.columns]

        # Verificar que la columna de alumnos esté presente
        if self.student_col not in self.df.columns:
            raise KeyError(
                f"La columna de alumnos '{self.student_col}' no se encontró "
                f"en la pestaña '{sheet_name}'. "
                f"Columnas disponibles: {list(self.df.columns)}"
            )

        # Normalizar la columna de alumnos: mayúsculas y sin espacios extra
        self.df[self.student_col] = (
            self.df[self.student_col]
            .str.strip()
            .str.upper()
        )

        self.active_sheet = sheet_name
        logger.info(
            "Pestaña '%s' cargada: %d filas, %d columnas.",
            sheet_name, len(self.df), len(self.df.columns),
        )

    def get_grade(self, student_name: str, web_subject_name: str) -> Optional[str]:
        """Devuelve la calificación de un alumno para una materia dada.

        Usa el ``mapeo_materias`` del JSON para traducir el nombre de la materia
        tal como aparece en la web al nombre de la columna real en el Excel.

        Args:
            student_name: Nombre del alumno en formato ``'APELLIDO, NOMBRE'``.
                          La búsqueda no distingue mayúsculas/minúsculas ni
                          espacios sobrantes.
            web_subject_name: Nombre de la materia según el sistema web
                              (p. ej. ``"Matemática"``).

        Returns:
            La calificación como string si se encuentra y no está vacía.
            ``None`` si el alumno no existe o la celda está vacía.

        Raises:
            RuntimeError: Si se llama sin haber cargado una pestaña primero
                          con :meth:`load_sheet`.
            KeyError: Si ``web_subject_name`` no tiene mapeo en ``config.json``.
            KeyError: Si la columna mapeada no existe en el DataFrame cargado.
        """
        self._require_loaded_sheet()

        # 1. Resolver el nombre de la columna en el Excel
        excel_col = self._resolve_subject_column(web_subject_name)

        # 2. Buscar al alumno (normalización idéntica a la del load)
        normalized_name = student_name.strip().upper()
        mask = self.df[self.student_col] == normalized_name
        matching_rows = self.df[mask]

        if matching_rows.empty:
            logger.warning(
                "Alumno '%s' no encontrado en la pestaña '%s'.",
                student_name, self.active_sheet,
            )
            return None

        if len(matching_rows) > 1:
            logger.warning(
                "Se encontraron %d filas para el alumno '%s'. Se usa la primera.",
                len(matching_rows), student_name,
            )

        # 3. Obtener la calificación
        raw_value: str = matching_rows.iloc[0][excel_col]

        if raw_value == "" or pd.isna(raw_value):
            logger.info(
                "La nota de '%s' en '%s' está vacía.",
                student_name, web_subject_name,
            )
            return None

        logger.info(
            "Nota de '%s' en '%s' (%s): %s",
            student_name, web_subject_name, excel_col, raw_value,
        )
        return raw_value.strip()

    def list_students(self) -> list[str]:
        """Devuelve la lista de alumnos de la pestaña actualmente cargada.

        Returns:
            Lista de nombres de alumnos (ya normalizados a mayúsculas).

        Raises:
            RuntimeError: Si no hay ninguna pestaña cargada.
        """
        self._require_loaded_sheet()
        return self.df[self.student_col].tolist()

    def list_subjects(self) -> list[str]:
        """Devuelve los nombres de materias disponibles según el mapeo del JSON.

        Returns:
            Lista de nombres de materias tal como aparecen en el sistema web.
        """
        return list(self.subject_mapping.keys())

    # ------------------------------------------------------------------
    # Métodos privados / helpers
    # ------------------------------------------------------------------

    def _require_loaded_sheet(self) -> None:
        """Verifica que haya una pestaña cargada en memoria.

        Raises:
            RuntimeError: Si ``self.df`` es ``None``.
        """
        if self.df is None:
            raise RuntimeError(
                "No hay ninguna pestaña cargada. "
                "Llamá a `load_sheet(sheet_name)` antes de hacer consultas."
            )

    def _resolve_subject_column(self, web_subject_name: str) -> str:
        """Traduce el nombre web de una materia a la columna del Excel.

        Args:
            web_subject_name: Nombre de la materia en el sistema web.

        Returns:
            Nombre de la columna correspondiente en el Excel.

        Raises:
            KeyError: Si la materia no está en el mapeo o la columna
                      no existe en el DataFrame cargado.
        """
        if web_subject_name not in self.subject_mapping:
            raise KeyError(
                f"La materia '{web_subject_name}' no tiene mapeo en config.json. "
                f"Materias disponibles: {list(self.subject_mapping.keys())}"
            )

        excel_col = self.subject_mapping[web_subject_name]

        if excel_col not in self.df.columns:
            raise KeyError(
                f"La columna '{excel_col}' (mapeada desde '{web_subject_name}') "
                f"no existe en la pestaña '{self.active_sheet}'. "
                f"Columnas disponibles: {list(self.df.columns)}"
            )

        return excel_col
    

if __name__ == "__main__":
    try:
        # 1. Inicializar el lector usando la configuración de la raíz
        lector = ExcelReader()

        # 2. Cargar la primera pestaña disponible para probar la consulta
        primera_pestana = pd.ExcelFile(lector.excel_path).sheet_names[0]
        lector.load_sheet(primera_pestana)

        # 3. Datos de prueba basados en tu captura real
        alumno_test = "BULACIO PONCE, BASTIAN FABRICIO"
        materia_test = "CIENCIAS NATURALES"

        # Intentamos obtener la nota con la API pública real
        nota = lector.get_grade(alumno_test, materia_test)
        
        print("\n=== ¡TEST EXITOSO! ===")
        print(f"Pestaña cargada: {primera_pestana}")
        print(f"Alumno: {alumno_test}")
        print(f"Materia (Web): {materia_test}")
        print(f"Nota encontrada en Excel: {nota}")
        print("======================\n")
        
    except Exception as e:
        print(f"\n❌ Error en el test revisado: {e}\n")