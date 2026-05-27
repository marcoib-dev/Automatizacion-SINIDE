"""web_handler.py
-----------------
Módulo para interactuar con la interfaz web de SINIDE usando Selenium.

La estrategia del handler es conectarse a una sesión ya abierta de Brave
mediante remote debugging en el puerto 9222, y operar sobre la tabla de
alumnos usando selectores declarados en ``config.json``.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Callable, Optional

from selenium import webdriver
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from excel_reader import CONFIG_PATH, load_config


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logger = logging.getLogger(__name__)
logging.basicConfig(
	level=logging.INFO,
	format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
)


GradeResolver = Callable[[str], Optional[object]]


class WebHandler:
	"""Controla la carga de notas en la tabla web de SINIDE.

	Example:
		>>> from excel_reader import ExcelReader
		>>> excel = ExcelReader()
		>>> excel.load_sheet("Hoja1")
		>>> handler = WebHandler()
		>>> handler.sync_grades(lambda student: excel.get_grade(student, "CIENCIAS NATURALES"))
	"""

	def __init__(self, config_path: Path = CONFIG_PATH, driver: Optional[webdriver.Chrome] = None) -> None:
		"""Inicializa el handler y se conecta al Brave ya abierto.

		Args:
			config_path: Ruta al archivo ``config.json`` de la raíz del proyecto.
			driver: Instancia de WebDriver opcional para inyectar desde tests.

		Raises:
			KeyError: Si faltan claves de Selenium en el JSON.
			TimeoutException: Si la tabla no está disponible al inicializar.
		"""
		self.config = load_config(config_path)

		selenium_config = self.config.get("selenium")
		if not selenium_config:
			raise KeyError("Falta la sección 'selenium' en config.json.")

		self.timeout: int = int(selenium_config.get("timeout_segundos", 10))
		self.selectors: dict[str, str] = selenium_config.get("selectores", {})

		required_selectors = {"tabla_alumnos", "fila_alumno", "celda_nombre", "celda_nota_1t"}
		missing_selectors = required_selectors - self.selectors.keys()
		if missing_selectors:
			raise KeyError(f"Faltan selectores obligatorios en config.json: {missing_selectors}")

		self.driver: webdriver.Chrome = driver or self._create_driver()
		self.wait = WebDriverWait(self.driver, self.timeout)

		logger.info("WebHandler inicializado y conectado a Brave en 127.0.0.1:9222.")

	# ------------------------------------------------------------------
	# Ciclo de vida
	# ------------------------------------------------------------------

	def close(self) -> None:
		"""Cierra el WebDriver sin afectar la ventana de Brave."""
		if self.driver is not None:
			self.driver.quit()

	def __enter__(self) -> "WebHandler":
		return self

	def __exit__(self, exc_type, exc, tb) -> None:
		self.close()

	# ------------------------------------------------------------------
	# Conexión / inicialización
	# ------------------------------------------------------------------

	def _create_driver(self) -> webdriver.Chrome:
		"""Crea un ChromeDriver conectado a la sesión remota de Brave."""
		options = webdriver.ChromeOptions()
		options.debugger_address = "127.0.0.1:9222"

		try:
			driver = webdriver.Chrome(options=options)
		except Exception:
			logger.exception("No se pudo conectar a Brave usando remote debugging en 9222.")
			raise

		logger.info("ChromeDriver conectado correctamente a la sesión remota.")
		return driver

	# ------------------------------------------------------------------
	# Selectores y esperas
	# ------------------------------------------------------------------

	def _wait_for_table(self) -> WebElement:
		"""Espera hasta que la tabla principal esté presente en la página."""
		locator = (By.XPATH, self.selectors["tabla_alumnos"])
		return self.wait.until(EC.presence_of_element_located(locator))

	def _get_table_rows(self) -> list[WebElement]:
		"""Devuelve las filas visibles de la tabla de alumnos."""
		table = self._wait_for_table()
		rows = table.find_elements(By.XPATH, self.selectors["fila_alumno"])

		if not rows:
			logger.warning("La tabla fue encontrada, pero no contiene filas visibles.")

		return rows

	def _wait_for_row_input(self, row: WebElement) -> WebElement:
		"""Espera a que el input de nota dentro de una fila esté disponible."""
		locator = (By.XPATH, self.selectors["celda_nota_1t"])

		def find_input(_driver: webdriver.Chrome) -> WebElement:
			return row.find_element(*locator)

		return self.wait.until(find_input)

	# ------------------------------------------------------------------
	# Extracción de datos
	# ------------------------------------------------------------------

	def get_student_name_from_row(self, row: WebElement) -> str:
		"""Extrae el nombre del alumno desde una fila de la tabla."""
		locator = (By.XPATH, self.selectors["celda_nombre"])
		cell = row.find_element(*locator)
		student_name = cell.text.strip()

		if not student_name:
			raise ValueError("Se encontró una fila sin nombre de alumno visible.")

		return student_name

	def list_web_students(self) -> list[str]:
		"""Devuelve los nombres de alumnos visibles en la tabla web."""
		students: list[str] = []
		for row in self._get_table_rows():
			try:
				students.append(self.get_student_name_from_row(row))
			except (StaleElementReferenceException, ValueError):
				logger.warning("No se pudo leer una fila de la tabla; se omite.")
		return students

	# ------------------------------------------------------------------
	# Escritura de notas
	# ------------------------------------------------------------------

	def write_grade_in_row(self, row: WebElement, grade: object) -> None:
		"""Escribe una nota en el input asociado a una fila."""
		input_element = self._wait_for_row_input(row)
		grade_text = "" if grade is None else str(grade).strip()

		if not grade_text:
			logger.info("Se omitió una fila porque la nota está vacía.")
			return

		input_element.click()
		input_element.send_keys(Keys.CONTROL, "a")
		input_element.send_keys(Keys.BACKSPACE)
		input_element.send_keys(grade_text)

	def sync_grades(self, grade_resolver: GradeResolver) -> int:
		"""Recorre la tabla y escribe la nota que devuelva ``grade_resolver``.

		Args:
			grade_resolver: Función que recibe el nombre del alumno y devuelve
				la nota a escribir. Si devuelve ``None`` o una cadena vacía, la
				fila se omite.

		Returns:
			Cantidad de filas actualizadas.
		"""
		updated_rows = 0
		rows = self._get_table_rows()

		for index, row in enumerate(rows, start=1):
			try:
				student_name = self.get_student_name_from_row(row)
				grade = grade_resolver(student_name)

				if grade is None or str(grade).strip() == "":
					logger.info("Sin nota para '%s'; fila %d omitida.", student_name, index)
					continue

				self.write_grade_in_row(row, grade)
				updated_rows += 1
				logger.info("Nota escrita para '%s' en la fila %d: %s", student_name, index, grade)

			except StaleElementReferenceException:
				logger.warning("La fila %d cambió mientras se procesaba; se omite.", index)
			except TimeoutException:
				logger.warning("No se pudo localizar el input de nota en la fila %d; se omite.", index)

		logger.info("Sincronización finalizada. Filas actualizadas: %d.", updated_rows)
		return updated_rows

	def fill_note_for_student(self, student_name: str, grade: object) -> bool:
		"""Busca un alumno por nombre y escribe su nota en la fila correspondiente."""
		rows = self._get_table_rows()

		for row in rows:
			try:
				current_student = self.get_student_name_from_row(row)
				if current_student.strip().upper() != student_name.strip().upper():
					continue

				self.write_grade_in_row(row, grade)
				logger.info("Nota escrita para '%s': %s", student_name, grade)
				return True
			except StaleElementReferenceException:
				logger.warning("La fila del alumno '%s' cambió durante la búsqueda.", student_name)
				break

		logger.warning("No se encontró la fila del alumno '%s' en la tabla web.", student_name)
		return False


if __name__ == "__main__":
	try:
		with WebHandler() as handler:
			print("WebHandler inicializado correctamente. Conectado a Brave.")
			print("Alumnos visibles:")
			for student in handler.list_web_students():
				print(f"- {student}")
	except Exception as exc:
		print(f"\n❌ Error al inicializar WebHandler: {exc}\n")
