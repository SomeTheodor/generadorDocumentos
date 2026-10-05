# gendoc — Plantilla .docx (Jinja2) + JSON → PDF y/o Word

Dos formas de usarlo: por línea de comandos (`gendoc.py`) o con una
interfaz web local para probar más cómodo (`app.py`).

## Requisitos
```bash
pip install docxtpl flask     # flask solo si vas a usar la interfaz
```
Además necesitás **LibreOffice** instalado para el paso docx→PDF (si sólo
generás Word no hace falta). No hace falta
tocar el PATH: gendoc lo busca solo en el PATH y en las rutas habituales
(`C:\Program Files\LibreOffice\program\soffice.exe`, `/Applications/LibreOffice.app/...`,
`/usr/bin/soffice`). Si lo tenés en otro lado:

```bash
python gendoc.py plantilla.docx datos.json --soffice "C:/Program Files/LibreOffice/program/soffice.exe"
```

o dejalo fijo en una variable de entorno (también la usa `app.py`):

```powershell
$env:GENDOC_SOFFICE = "C:\Program Files\LibreOffice\program\soffice.exe"   # PowerShell, sesión actual
setx GENDOC_SOFFICE "C:\Program Files\LibreOffice\program\soffice.exe"     # permanente
```

## Interfaz web (recomendado para probar)
```bash
python app.py            # abrí http://localhost:5000
```
Subís la plantilla, pegás el JSON, elegís el formato (**PDF**, **Word** o
**Ambos**) y generás. El PDF se previsualiza ahí mismo; el Word se descarga.
Botón **Validar** para ver qué variables espera la plantilla sin generar. Si algo falla, el panel de la derecha te muestra la etapa, la
ubicación y el detalle — y si el error es del JSON, te marca la línea en el
editor. La interfaz reutiliza `gendoc.py`, así que cualquier mejora al script
la toma automáticamente.

## Línea de comandos
```bash
python gendoc.py plantilla.docx datos.json -o salida.pdf               # PDF
python gendoc.py plantilla.docx datos.json -o salida.docx              # Word (por la extensión)
python gendoc.py plantilla.docx datos.json -o salida --formato ambos   # salida.pdf + salida.docx
python gendoc.py plantilla.docx datos.json --check                     # solo valida y lista variables
```
`--formato` (`-f`) acepta `pdf`, `docx` o `ambos`; si no lo pasás, se deduce
de la extensión de `-o` (por defecto, PDF). `--keep-docx` sigue andando y
equivale a `--formato ambos`. Si no indicás `-o` y pedís Word, se guarda como
`<plantilla>_generado.docx` para no pisar la plantilla.

## Sintaxis en el Word (docxtpl)
- `{{ cliente.nombre }}` — variable / atributo
- `{% if x %}...{% endif %}` — condicional inline
- `{%p for it in items %}` ... `{%p endfor %}` — repetir párrafos
- `{%tr for it in items %}` ... `{%tr endfor %}` — repetir filas de tabla

**Gotcha de tablas:** poné `{%tr for %}` en la primera celda y `{%tr endfor %}`
en la última celda de **la misma fila**. Si te tira "unknown tag endfor",
suele ser que las etiquetas quedaron descolocadas entre celdas.

## JSON envuelto en una clave
Si el JSON trae todo adentro de una clave contenedora:

```json
{ "pol": { "pol": {...}, "aseg_1": {...}, "cobertura": [...] } }
```

y la plantilla usa los nombres de adentro (`{{ aseg_1.businessName }}`,
`{% for c in cobertura %}`), gendoc abre ese nivel automáticamente: sube el
contenido al nivel superior y avisa cuál envoltorio abrió. Lo de adentro pisa
a lo de afuera, así que en el ejemplo `pol` termina apuntando al objeto interno
`pol`, no al envoltorio. Sólo abre una clave si eso resuelve variables que
faltaban, y soporta hasta 3 niveles de envoltura.

## Bloques opcionales
Si una variable sólo se usa dentro de su propio `{% if %}`, por ejemplo
`{% if coaseguro1 %}…{{ coaseguro1.businessName }}…{% endif %}`, se la toma
como opcional: si no viene en el JSON, no cuenta como faltante y el bloque
simplemente no se imprime. Si se usa también fuera de su `{% if %}`, sigue
siendo obligatoria. Sirven tanto `{% if x %}` como `{% if x is defined %}`.

Lo mismo para listas de ítems opcionales:

```
{% set lista = [datagen_1, datagen_2, …, datagen_10] %}
{% for p in lista %}{% if p %} … {{ p.attr3 }} … {% endif %}{% endfor %}
```

Los `datagen_N` que no vengan en el JSON se saltean.

## Qué te dice cuando falla
- JSON mal formado → línea y columna (las comas de más antes de un `}` o `]` no frenan la generación: se ignoran y te avisa)
- Variable que la plantilla espera y no está en el JSON → nombre exacto (y `--check`/Validar las listan todas)
- Error de sintaxis Jinja → mensaje + fragmento del XML
- Falla de LibreOffice → su salida de error

## Archivos
- `app.py` — interfaz web (Flask)
- `gendoc.py` — motor + CLI
- `plantilla.docx`, `datos.json` — ejemplo listo para probar
- `presupuesto.pdf` — salida de ejemplo
