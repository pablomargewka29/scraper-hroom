# Hiring Room Scraper

Aplicación de escritorio en Python que extrae ofertas laborales desde portales de [Hiring Room](https://hiringroom.com/) y las exporta a Excel.

Forma parte de un sistema de búsqueda laboral automatizada junto con [jobfit](#) (workflow en n8n + IA que evalúa las ofertas extraídas contra un CV).

## Qué hace

- Lee un Excel con una lista de empresas y sus portales de Hiring Room.
- Scrapea cada portal en paralelo (hasta 5 vacantes simultáneas) usando Playwright.
- Filtra ofertas por **sector**, **fecha de publicación** (hoy / semana / mes) y **palabras clave**.
- Exporta los resultados a un Excel formateado con: empresa, puesto, descripción, requisitos, fecha y enlace.

Incluye una interfaz gráfica (Tkinter) con pausa, reanudación, exportación parcial y log en vivo.

## Stack

- Python 3.10+
- Playwright (scraping asíncrono)
- Tkinter (GUI)
- pandas + openpyxl (lectura/exportación de Excel)

## Cómo correrlo

```bash
# 1. Clonar el repo
git clone https://github.com/pablomargewka29/scraper-hroom.git
cd scraper-hroom

# 2. Instalar dependencias
pip install -r requirements.txt
playwright install chromium

# 3. Ejecutar
python scraper.py
```

## Formato del Excel de entrada

El archivo debe tener al menos dos columnas:

| empresa     | enlace                            | sector (opcional) |
|-------------|-----------------------------------|-------------------|
| Empresa A   | https://empresaA.hiringroom.com   | Tecnología        |
| Empresa B   | https://empresaB.hiringroom.com   | Salud             |

Si no hay columna `sector`, todos los portales se procesan sin filtrado por rubro.

## Notas

- Solo funciona con portales que usan la plataforma Hiring Room.
- Pensado para uso personal en búsqueda laboral.

---

Creado por [Pablo Margewka](https://www.linkedin.com/in/pablomargewka).
