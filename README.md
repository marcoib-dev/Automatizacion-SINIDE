# SINIDE - Automatización de Carga de Calificaciones 

Este proyecto surge como una solución a una problemática real en la gestión educativa: la carga manual, tediosa y propensa a errores de las calificaciones trimestrales en la plataforma estatal SINIDE.

El script automatiza el proceso utilizando **Python** y **Selenium**, permitiendo a los usuarios (preceptores o docentes) iniciar sesión de forma segura y delegar la carga repetitiva de datos al bot.

## Características Principales

- **Enfoque UX Real:** El script se adapta a las planillas tradicionales utilizadas por las escuelas (matrices divididas por pestañas de trimestres), evitando que el usuario deba reestructurar sus datos.
- **Seguridad por Diseño:** No almacena credenciales ni automatiza el login. El usuario se autentica manualmente y activa el script directamente en la pantalla de carga, evitando bloqueos de seguridad del sistema.
- **Robustez contra desfasajes:** Mapea las notas buscando de forma exacta el nombre del alumno en la interfaz web, impidiendo que un orden diferente entre el Excel y el sistema provoque errores en las calificaciones.
- **Configuración desacoplada:** Selectores web y mapeos de nombres gestionados a través de un archivo `config.json` para facilitar el mantenimiento.

## Tecnologías utilizadas

- **Python 3**
- **Selenium WebDriver** (Automatización del navegador)
- **Pandas** & **OpenPyXL** (Procesamiento eficiente de datos y lectura de matrices Excel)

## Estructura del Proyecto

```text
├── config.json          # Configuración de selectores web y mapeo de materias
├── requirements.txt     # Dependencias del proyecto
├── data/
│   └── notas_ejemplo.xlsx  # Plantilla de ejemplo con datos ficticios
└── src/
    ├── main.py          # Orquestador principal del flujo
    ├── excel_reader.py  # Lógica de lectura y filtrado del Excel con Pandas
    └── web_handler.py   # Interacción con el navegador mediante Selenium


