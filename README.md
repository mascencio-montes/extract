# Extractores Santander — Automation Anywhere A360

Extractores Python locales para PDFs de Santander, organizados en dos procesos independientes. **Estado: versión inicial, no validada para producción financiera.**

## Contenido

- `bbva/extractor_bbva.py`: extracción de estados de cuenta BBVA Fideicomisos.
- `bbva/LEEME_A360.txt`: instalación, ejecución y salidas BBVA.

- `reportos/extractor_santander.py`: confirmaciones de operaciones de reporto.
- `reportos/LEEME_Reportos.txt`: instrucciones y salidas.
- `posicion_diaria/extractor_posicion_diaria.py`: estados de cuenta / posición diaria.
- `posicion_diaria/LEEME_Posicion_Diaria.txt`: instrucciones y salidas.

No se publican PDFs bancarios, salidas financieras, cuentas ni credenciales.

## Instalación

En Windows:

```powershell
py -m pip install pymupdf pdfplumber artifact-tool openpyxl
```

## Ejecución recursiva por meses

```powershell
py .\reportos\extractor_santander.py "C:\RPA\SANTANDER\REPORTOS"
py .\posicion_diaria\extractor_posicion_diaria.py "C:\RPA\SANTANDER\POSICION DIARIA" --salida "C:\RPA\RESULTADOS\POSICION_DIARIA"
```

Para BBVA, usar `--recursivo` si los PDF están organizados en subcarpetas:

```powershell
py .\bbva\extractor_bbva.py "C:\RPA\BBVA\Entrada" --recursivo
```

Se recorren las subcarpetas `1. Enero`, `2. Febrero`, etcétera.

## Advertencia

Son copias de los scripts previamente entregados en la conversación. Deben probarse y ajustarse contra los PDF originales, especialmente reconocimiento de tablas, múltiples operaciones y conciliación de cifras. No confiar en resultados sin auditoría manual y automatizada. No subir PDF reales ni salidas financieras a este repositorio público.
