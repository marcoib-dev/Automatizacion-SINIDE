"""
web_handler.py
--------------
Controla la interfaz web de SINIDE desde dentro del iframe AngularJS.

Puntos clave del HTML real:
  - Todo el formulario vive en un <iframe src="/ui/index.html">
  - Los dropdowns son Bootstrap: button.dropdown-toggle + ul.dropdown-menu li a span
  - La celda de nota muestra un <span>S/C</span> hasta que se hace clic,
    momento en que AngularJS lo reemplaza por un <input>
  - IDs de los dropdowns: #calificaciones-materias-selector, #periodo-notas-selector
"""

from __future__ import annotations

import json
import logging
import time
from pathlib import Path
from typing import Callable, Optional

from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from excel_reader import CONFIG_PATH, load_config

logger = logging.getLogger(__name__)

GradeResolver = Callable[[str], Optional[str]]

_DROPDOWN_SETTLE = 1.2   # segundos tras seleccionar en un dropdown
_TABLE_SETTLE    = 1.5   # segundos tras cambiar materia/trimestre


class WebHandler:
    """Interactúa con la tabla de calificaciones de SINIDE (dentro del iframe)."""

    def __init__(
        self,
        config_path: Path = CONFIG_PATH,
        driver: Optional[webdriver.Chrome] = None,
    ) -> None:
        self.config    = load_config(config_path)
        sel_cfg        = self.config.get("selenium", {})
        self.timeout   = int(sel_cfg.get("timeout_segundos", 15))
        self.sel       = sel_cfg.get("selectores", {})

        self.driver = driver or self._create_driver()
        self.wait   = WebDriverWait(self.driver, self.timeout)

        logger.info("WebHandler conectado. Entrando al iframe...")
        self._switch_to_iframe()
        logger.info("Contexto dentro del iframe establecido.")

    # ------------------------------------------------------------------
    # Ciclo de vida
    # ------------------------------------------------------------------

    def close(self) -> None:
        if self.driver:
            self.driver.quit()

    def __enter__(self) -> "WebHandler":
        return self

    def __exit__(self, *_) -> None:
        self.close()

    # ------------------------------------------------------------------
    # Conexión al navegador
    # ------------------------------------------------------------------

    def _create_driver(self) -> webdriver.Chrome:
        from webdriver_manager.chrome import ChromeDriverManager
        from webdriver_manager.core.os_manager import ChromeType
        from selenium.webdriver.chrome.service import Service

        options = webdriver.ChromeOptions()
        options.debugger_address = "127.0.0.1:9222"
        try:
            service = Service(
                ChromeDriverManager(chrome_type=ChromeType.BRAVE).install()
            )
            return webdriver.Chrome(service=service, options=options)
        except Exception:
            logger.exception("No se pudo conectar a Brave en el puerto 9222.")
            raise

    def _switch_to_sinide_tab(self) -> None:
        """Cambia el foco de Selenium a la pestaña de SINIDE."""
        for handle in self.driver.window_handles:
            self.driver.switch_to.window(handle)
            if "sge.meducacionsantiago" in self.driver.current_url or \
               "notas" in self.driver.current_url:
                logger.info("Pestaña SINIDE encontrada: %s", self.driver.current_url)
                return
        # Si no encontramos por URL, usar la última pestaña disponible
        logger.warning("No se encontró pestaña de SINIDE por URL. Usando la pestaña activa.")

    def _switch_to_iframe(self) -> None:
        """Cambia el contexto de Selenium al interior del iframe de notas."""
        # Primero asegurarse de estar en la pestaña correcta
        self._switch_to_sinide_tab()

        # Volver al contexto principal del documento
        self.driver.switch_to.default_content()

        # Esperar a que el iframe esté presente
        iframe = self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, self.sel["iframe"]))
        )
        self.driver.switch_to.frame(iframe)

        # Esperar a que el contenido del iframe cargue (tabla de notas)
        self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table.sge-notas"))
        )
        logger.info("Contexto cambiado al iframe. Tabla de notas visible.")

    def _ensure_iframe_context(self) -> None:
        """Verifica que seguimos dentro del iframe; si no, vuelve a entrar."""
        try:
            # Si esto funciona, ya estamos dentro del iframe
            self.driver.find_element(By.CSS_SELECTOR, "table.sge-notas")
        except NoSuchElementException:
            logger.warning("Perdimos el contexto del iframe. Volviendo a entrar...")
            self._switch_to_iframe()

    # ------------------------------------------------------------------
    # Selección de dropdowns Bootstrap
    # ------------------------------------------------------------------

    def select_trimester(self, trimester_name: str) -> None:
        """Selecciona el trimestre en el dropdown #periodo-notas-selector."""
        logger.info("Seleccionando trimestre: '%s'.", trimester_name)
        self._select_bootstrap_dropdown(
            self.sel["dropdown_trimestre"],
            self.sel["items_trimestre"],
            trimester_name,
        )
        time.sleep(_TABLE_SETTLE)

    def select_subject(self, subject_name: str) -> None:
        """Selecciona la materia en el dropdown #calificaciones-materias-selector."""
        logger.info("Seleccionando materia: '%s'.", subject_name)
        self._select_bootstrap_dropdown(
            self.sel["dropdown_materia"],
            self.sel["items_materia"],
            subject_name,
        )
        time.sleep(_TABLE_SETTLE)

    def _select_bootstrap_dropdown(
        self, toggle_css: str, items_css: str, target_text: str
    ) -> None:
        """
        Abre un dropdown Bootstrap y hace clic en el item cuyo texto
        coincida exactamente con target_text.

        Los dropdowns de SINIDE tienen esta estructura:
          <div class="btn-group dropdown">
            <button class="btn btn-primary dropdown-toggle">...</button>
            <ul class="dropdown-menu">
              <li class="clickable"><a><span>TEXTO</span></a></li>
            </ul>
          </div>
        """
        self._ensure_iframe_context()

        # 1. Abrir el dropdown
        toggle = self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, toggle_css))
        )
        toggle.click()
        time.sleep(0.4)  # esperar animación Bootstrap

        # 2. Buscar la opción por texto exacto
        items = self.wait.until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, items_css))
        )

        for item in items:
            if item.text.strip() == target_text:
                item.click()
                logger.debug("Opción '%s' seleccionada.", target_text)
                return

        raise ValueError(
            f"No se encontró la opción '{target_text}' en el dropdown. "
            f"Opciones disponibles: {[i.text.strip() for i in items]}"
        )

    # ------------------------------------------------------------------
    # Lectura de la tabla
    # ------------------------------------------------------------------

    def list_web_students(self) -> list[str]:
        """Devuelve los nombres de alumnos visibles en la tabla."""
        self._ensure_iframe_context()
        students = []
        for row in self._get_rows():
            try:
                students.append(self._get_name(row))
            except (StaleElementReferenceException, ValueError, NoSuchElementException):
                logger.warning("No se pudo leer una fila; se omite.")
        return students

    def _get_rows(self) -> list[WebElement]:
        """Devuelve las filas <tr> del tbody de la tabla de notas."""
        self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, self.sel["filas_alumnos"]))
        )
        rows = self.driver.find_elements(By.CSS_SELECTOR, self.sel["filas_alumnos"])
        if not rows:
            logger.warning("La tabla no tiene filas visibles.")
        return rows

    def _get_name(self, row: WebElement) -> str:
        """Extrae el nombre del alumno de la primera celda de la fila."""
        span = row.find_element(By.CSS_SELECTOR, self.sel["celda_nombre"])
        name = span.text.strip()
        if not name:
            raise ValueError("Fila sin nombre.")
        return name

    # ------------------------------------------------------------------
    # Escritura de notas
    # ------------------------------------------------------------------

    def sync_grades(self, grade_resolver: GradeResolver) -> int:
        """
        Recorre la tabla fila por fila y escribe las notas.

        Para cada alumno:
          1. Obtiene el nombre del alumno.
          2. Llama a grade_resolver(nombre) para obtener la nota.
          3. Hace clic en el span S/C para activar el input de AngularJS.
          4. Escribe la nota y confirma con TAB.

        Returns:
            Cantidad de filas actualizadas.
        """
        self._ensure_iframe_context()
        updated = 0
        rows    = self._get_rows()

        for idx, row in enumerate(rows, 1):
            try:
                name  = self._get_name(row)
                grade = grade_resolver(name)

                if not grade or str(grade).strip() == "":
                    logger.info("[%d/%d] Sin nota para '%s'; omitido.", idx, len(rows), name)
                    continue

                self._write_grade(row, grade)
                updated += 1
                logger.info("[%d/%d] Nota '%s' escrita para '%s'.", idx, len(rows), grade, name)

            except StaleElementReferenceException:
                logger.warning("Fila %d cambió durante el procesamiento; omitida.", idx)
            except TimeoutException:
                logger.warning("Timeout esperando input en fila %d; omitida.", idx)
            except Exception:
                logger.exception("Error inesperado en fila %d.", idx)

        logger.info("Sync completo: %d/%d filas actualizadas.", updated, len(rows))
        return updated

    def _write_grade(self, row: WebElement, grade: object) -> None:
        """
        Escribe una nota en la celda de la fila dada.

        Estrategia para el editor AngularJS de SINIDE:
          1. La celda muestra <span>S/C</span> o el valor anterior.
          2. Hacer clic en el span activa el modo edición: el span
             desaparece y aparece un <input>.
          3. Limpiar el input y escribir la nota.
          4. Enviar TAB para confirmar y pasar a la siguiente celda.
        """
        grade_str = str(grade).strip()

        # Paso 1: hacer clic en el span S/C (o valor existente) para abrir el input
        try:
            span = row.find_element(By.CSS_SELECTOR, self.sel["celda_nota_sc"])
            span.click()
        except NoSuchElementException:
            # Si ya está en modo edición (input visible), continuar
            pass

        # Paso 2: esperar que aparezca el input
        input_el = WebDriverWait(row, self.timeout).until(
            lambda r: r.find_element(By.CSS_SELECTOR, self.sel["input_nota"])
        )

        # Paso 3: limpiar y escribir
        input_el.click()
        input_el.send_keys(Keys.CONTROL, "a")
        input_el.send_keys(Keys.DELETE)
        input_el.send_keys(grade_str)

        # Paso 4: confirmar con TAB (AngularJS guarda al perder el foco)
        input_el.send_keys(Keys.TAB)
        time.sleep(0.3)
