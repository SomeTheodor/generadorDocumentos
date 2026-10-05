#!/usr/bin/env python3
"""
gendoc.py — Genera un PDF y/o un Word (.docx) a partir de una plantilla Word
con sintaxis Jinja2 y un archivo JSON de datos.

Uso:
    python gendoc.py plantilla.docx datos.json -o salida.pdf
    python gendoc.py plantilla.docx datos.json -o salida.docx     # Word (sin LibreOffice)
    python gendoc.py plantilla.docx datos.json -o salida --formato ambos
    python gendoc.py plantilla.docx datos.json --check      # sólo valida, no genera
    python gendoc.py plantilla.docx datos.json -o out.pdf --keep-docx
    python gendoc.py plantilla.docx datos.json --soffice "C:/Program Files/LibreOffice/program/soffice.exe"

LibreOffice se busca solo: PATH, rutas habituales de instalación, o la ruta que
se indique con --soffice / la variable de entorno GENDOC_SOFFICE.

En la plantilla .docx se escriben las variables con la sintaxis de docxtpl:
    {{ cliente.nombre }}          -> variable / atributo
    {% for item in items %}...{% endfor %}   -> bucles
    {%p ... %}, {%tr ... %}       -> etiquetas de párrafo / fila de tabla

Si el JSON viene envuelto en una clave contenedora —p. ej. todo adentro de
{"pol": {...}}— ese nivel se abre solo, para que la plantilla pueda usar
directamente los nombres de adentro ({{ aseg_1.businessName }}, {{ cobertura }}).

Si algo falla, el script te dice EN QUÉ ETAPA y CON QUÉ DETALLE:
    - JSON mal formado  -> línea y columna
    - variables que la plantilla espera y faltan en el JSON -> lista de nombres
    - error de sintaxis Jinja -> mensaje + fragmento del documento
    - error al convertir a PDF -> salida de LibreOffice
"""

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
from pathlib import Path
from typing import NamedTuple

from docxtpl import DocxTemplate
from jinja2 import Environment, StrictUndefined
from jinja2.exceptions import TemplateSyntaxError, UndefinedError


# --------------------------------------------------------------------------- #
# Manejo de errores: una excepción propia que lleva etapa + detalle + ayuda.
# --------------------------------------------------------------------------- #
class GenError(Exception):
    def __init__(self, etapa, detalle, ubicacion=None, sugerencia=None):
        self.etapa = etapa
        self.detalle = detalle
        self.ubicacion = ubicacion
        self.sugerencia = sugerencia
        super().__init__(detalle)

    def imprimir(self):
        print(f"\n✖ Error en la etapa: {self.etapa}", file=sys.stderr)
        if self.ubicacion:
            print(f"  Ubicación: {self.ubicacion}", file=sys.stderr)
        print(f"  Detalle:   {self.detalle}", file=sys.stderr)
        if self.sugerencia:
            print(f"  Sugerencia: {self.sugerencia}", file=sys.stderr)
        print("", file=sys.stderr)


# --------------------------------------------------------------------------- #
# 1) Cargar el JSON (con ubicación exacta si está mal formado)
# --------------------------------------------------------------------------- #
def _quitar_comas_sobrantes(texto: str):
    """Saca las comas que quedan justo antes de un } o un ] —JSON no las
    permite, pero es lo que más pasa al pegar o borrar un bloque a mano—.
    Ignora las que están dentro de una cadena. Devuelve (texto, cuántas)."""
    barra = chr(92)
    salida = []
    quitadas = 0
    en_cadena = escape = False
    for i, ch in enumerate(texto):
        if en_cadena:
            salida.append(ch)
            if escape:
                escape = False
            elif ch == barra:
                escape = True
            elif ch == '"':
                en_cadena = False
            continue
        if ch == '"':
            en_cadena = True
        elif ch == ",":
            resto = texto[i + 1:]
            if resto.lstrip()[:1] in ("}", "]"):
                quitadas += 1
                continue
        salida.append(ch)
    return "".join(salida), quitadas


def parsear_json(texto: str):
    """Parsea el texto JSON. Si lo único que falla son comas sobrantes antes de
    } o ], las saca y reintenta, avisando. Devuelve (datos, avisos)."""
    try:
        datos, avisos = json.loads(texto), []
    except json.JSONDecodeError as e:
        limpio, quitadas = _quitar_comas_sobrantes(texto)
        if not quitadas:
            raise _error_json(e)
        try:
            datos = json.loads(limpio)
        except json.JSONDecodeError:
            raise _error_json(e)   # el problema era otro: reportamos el original
        avisos = [f"Se ignoraron {quitadas} coma(s) de más antes de un cierre "
                  "de llave o corchete. Conviene sacarlas del JSON igual."]
    if not isinstance(datos, dict):
        raise GenError(
            "parseo del JSON",
            f"El JSON de nivel superior debe ser un objeto {{...}}, "
            f"pero es un {type(datos).__name__}.",
        )
    return datos, avisos


def _error_json(e: json.JSONDecodeError) -> GenError:
    """GenError para un JSON mal formado, con la línea marcada para el editor."""
    err = GenError(
        "parseo del JSON",
        e.msg,
        ubicacion=f"línea {e.lineno}, columna {e.colno} (car. {e.pos})",
        sugerencia="Revisá comas, comillas y llaves alrededor de esa posición.",
    )
    err.lineno, err.colno = e.lineno, e.colno
    return err


def cargar_json(path: Path):
    """Lee y parsea el archivo JSON. Devuelve (datos, avisos)."""
    if not path.exists():
        raise GenError("lectura de datos", f"No existe el archivo JSON: {path}")
    try:
        return parsear_json(path.read_text(encoding="utf-8"))
    except GenError as e:
        e.ubicacion = f"{path}" + (f" — {e.ubicacion}" if e.ubicacion else "")
        raise


# --------------------------------------------------------------------------- #
# 2) Cargar la plantilla .docx
# --------------------------------------------------------------------------- #
def cargar_plantilla(path: Path) -> DocxTemplate:
    if not path.exists():
        raise GenError("lectura de plantilla", f"No existe la plantilla: {path}")
    try:
        return DocxTemplate(str(path))
    except Exception as e:
        raise GenError(
            "apertura de la plantilla",
            f"{type(e).__name__}: {e}",
            ubicacion=str(path),
            sugerencia="¿Es un .docx válido? Los .doc viejos hay que convertirlos primero.",
        )


# --------------------------------------------------------------------------- #
# 3) Comparar variables esperadas por la plantilla vs. las del JSON
# --------------------------------------------------------------------------- #
class Analisis(NamedTuple):
    """Resultado de comparar la plantilla con los datos."""
    esperadas: list      # todas las variables de nivel superior que usa la plantilla
    faltantes: list      # las que la plantilla espera y el JSON no tiene
    sin_usar: list       # las que el JSON trae y la plantilla no usa
    contexto: dict       # los datos ya listos para renderizar (ver desanidar)
    envoltorios: list    # claves contenedoras que hubo que abrir, en orden
    opcionales: list = []  # bloques {% if x %} cuyo x no vino: se omiten
    atributos: list = []   # x.y que la plantilla usa y el JSON no trae (textos)


def desanidar(datos: dict, esperadas: set, max_niveles: int = 3):
    """Abre envoltorios: si el JSON viene como {"pol": {...todo...}} pero la
    plantilla usa los nombres de adentro (aseg_1, cobertura, ...), sube ese
    contenido al nivel superior para que las variables se resuelvan.

    Sólo abre una clave si eso resuelve variables que hoy faltan, y lo repite
    por si hay más de un nivel de envoltura. Lo de adentro pisa a lo de afuera,
    así {"pol": {"pol": {...}}} deja el objeto interno accesible como `pol`.

    Devuelve (contexto, envoltorios_abiertos).
    """
    ctx = dict(datos)
    abiertos = []
    for _ in range(max_niveles):
        faltan = esperadas - set(ctx)
        if not faltan:
            break
        candidatos = {k: v for k, v in ctx.items() if isinstance(v, dict)}
        # Primero, el que resuelve más variables al abrirlo.
        mejor = max(candidatos, key=lambda k: len(faltan & set(candidatos[k])),
                    default=None)
        if mejor is None or not (faltan & set(candidatos[mejor])):
            # Si ninguno resuelve directo, quizá la envoltura tiene otra adentro.
            mejor = max(
                candidatos,
                key=lambda k: max(
                    (len(faltan & set(sub)) for sub in candidatos[k].values()
                     if isinstance(sub, dict)),
                    default=0,
                ),
                default=None,
            )
            if mejor is None or not any(
                faltan & set(sub) for sub in candidatos[mejor].values()
                if isinstance(sub, dict)
            ):
                break
        ctx = {**ctx, **candidatos[mejor]}
        abiertos.append(mejor)
    return ctx, abiertos


def analizar_variables(doc: DocxTemplate, datos: dict) -> Analisis:
    """Compara plantilla y datos. Devuelve un Analisis (desempaquetable como
    tupla: esperadas, faltantes, sin_usar, contexto, envoltorios)."""
    try:
        env = Environment()
        esperadas = doc.get_undeclared_template_variables(env)
    except TemplateSyntaxError as e:
        _raise_syntaxis(doc, e)

    contexto, envoltorios = desanidar(datos, esperadas)
    presentes = set(contexto.keys())
    faltan = esperadas - presentes
    # Las que sólo se usan como {% if x %} (y adentro de su propio bloque) son
    # opcionales: si no vienen, valen None y el bloque no se imprime.
    opcionales = sorted(faltan & variables_opcionales(doc)) if faltan else []
    contexto = {**contexto, **dict.fromkeys(opcionales)}
    faltantes = sorted(faltan - set(opcionales))
    sin_usar = sorted(presentes - esperadas)
    return Analisis(sorted(esperadas), faltantes, sin_usar, contexto, envoltorios,
                    opcionales, atributos_faltantes(doc, contexto))


def _cadena(nodo):
    """Para `a.b.c` devuelve ("a", ["b", "c"]); si no es una cadena de
    atributos sobre un nombre, None."""
    from jinja2 import nodes
    attrs = []
    while isinstance(nodo, nodes.Getattr):
        attrs.append(nodo.attr)
        nodo = nodo.node
    return (nodo.name, attrs[::-1]) if isinstance(nodo, nodes.Name) else None


def atributos_faltantes(doc: DocxTemplate, contexto: dict) -> list:
    """Recorre la plantilla y, para cada `x.y.z` que use, se fija que el JSON
    tenga esos atributos. También dentro de los {% for a in lista %}, contra
    cada elemento de la lista. Devuelve textos como
    "cont_1.commissionPercent — «cont_1» no tiene «commissionPercent»"."""
    import difflib
    from jinja2 import nodes

    problemas = {}   # ruta -> texto (sin repetir)

    def chequear(raiz, attrs, valores, origen):
        fallos, detalle = 0, None
        for v in valores:
            ruta = raiz
            for attr in attrs:
                if not isinstance(v, dict):
                    break          # None (bloque opcional), lista, etc.: no opinamos
                if attr not in v:
                    fallos += 1
                    parecidos = difflib.get_close_matches(attr, list(v), n=1)
                    detalle = (f"«{ruta}» no tiene «{attr}»"
                               + (f" (¿quisiste decir «{parecidos[0]}»?)" if parecidos else ""))
                    break
                v, ruta = v[attr], f"{ruta}.{attr}"
        if fallos:
            clave = raiz + "." + ".".join(attrs)
            extra = (f" — en {fallos} de {len(valores)} elemento(s) de «{origen}»"
                     if origen else "")
            problemas.setdefault(clave, f"{clave} — {detalle}{extra}")

    def recorrer(nodo, alcance):
        if isinstance(nodo, nodes.For):
            recorrer(nodo.iter, alcance)
            nuevo = dict(alcance)
            c = _cadena(nodo.iter)
            if isinstance(nodo.target, nodes.Name):
                nuevo.pop(nodo.target.name, None)
                if c and c[0] in alcance:
                    items = []
                    for v in alcance[c[0]][0]:
                        for attr in c[1]:
                            v = v.get(attr) if isinstance(v, dict) else None
                        if isinstance(v, list):
                            items.extend(v)
                    nombre_lista = ".".join([c[0], *c[1]])
                    nuevo[nodo.target.name] = (items, nombre_lista)
            for hijo in nodo.body:
                recorrer(hijo, nuevo)
            for hijo in nodo.else_:
                recorrer(hijo, alcance)
            return
        if isinstance(nodo, nodes.Getattr):
            c = _cadena(nodo)
            if c and c[0] in alcance:
                valores, origen = alcance[c[0]]
                chequear(c[0], c[1], valores, origen)
                return
        for hijo in nodo.iter_child_nodes():
            recorrer(hijo, alcance)

    alcance = {k: ([v], None) for k, v in contexto.items()}
    env = Environment()
    try:
        for _, xml in _fuentes_jinja(doc):
            recorrer(env.parse(xml), alcance)
    except Exception:
        return []
    return list(problemas.values())


def _fuentes_jinja(doc: DocxTemplate):
    """XML de cada parte del .docx (cuerpo, encabezados, pies) ya preparado por
    docxtpl, que es justo lo que le pasa a Jinja. Devuelve [(nombre, xml), ...]."""
    from docx import Document as _Document
    from docx.oxml import parse_xml

    temp = _Document(doc.template_file)
    partes = [("cuerpo", doc.patch_xml(doc.xml_to_string(temp._element.body)))]
    for etiqueta, uri in (("encabezado", doc.HEADER_URI), ("pie de página", doc.FOOTER_URI)):
        for _, val in temp._part.rels.items():
            if val.reltype == uri and val.target_part.blob:
                partes.append(
                    (etiqueta, doc.patch_xml(doc.xml_to_string(parse_xml(val.target_part.blob))))
                )
    return partes


def _guardas(test) -> set:
    """Nombres que, si son falsos/None, hacen que no se entre al {% if %}:
    `x`, `x is defined`, `x and ...`."""
    from jinja2 import nodes
    if isinstance(test, nodes.Name):
        return {test.name}
    if isinstance(test, nodes.Test) and test.name == "defined" and isinstance(test.node, nodes.Name):
        return {test.node.name}
    if isinstance(test, nodes.And):
        return _guardas(test.left) | _guardas(test.right)
    return set()


def variables_opcionales(doc: DocxTemplate) -> set:
    """Variables que la plantilla sólo usa protegidas por su propio {% if %}:
    p. ej. {% if coaseguro1 %}...{{ coaseguro1.businessName }}...{% endif %}.
    Si no vienen en el JSON, el bloque simplemente no sale."""
    from jinja2 import nodes

    guardadas, libres = set(), set()

    def recorrer(nodo, protegidas):
        if isinstance(nodo, nodes.If):
            g = _guardas(nodo.test)
            guardadas.update(g)
            recorrer(nodo.test, protegidas | g)
            for hijo in nodo.body:
                recorrer(hijo, protegidas | g)
            # elif / else corren justamente cuando la guarda es falsa.
            for hijo in nodo.elif_ + nodo.else_:
                recorrer(hijo, protegidas)
            return
        if isinstance(nodo, nodes.Name) and nodo.ctx == "load" and nodo.name not in protegidas:
            libres.add(nodo.name)
        for hijo in nodo.iter_child_nodes():
            recorrer(hijo, protegidas)

    env = Environment()
    try:
        for _, xml in _fuentes_jinja(doc):
            recorrer(env.parse(xml), frozenset())
    except Exception:
        return set()   # ante la duda, nada es opcional (comportamiento estricto)
    return guardadas - libres


# Etiquetas Jinja tal como quedan en el documento, ya sin el XML del medio.
_RE_ETIQUETA = re.compile(r"\{%-?\s*(?:tr |p |tc |r )?\s*(\w+)(.*?)-?%\}", re.S)
_RE_USO = re.compile(r"\{\{-?\s*(\w+)")


def _partes_de_texto(doc: DocxTemplate):
    """Texto plano de cada parte del .docx (cuerpo, encabezados, pies), que es
    justo lo que docxtpl le pasa a Jinja al analizar variables.
    Devuelve [(nombre_parte, texto), ...]."""
    return [(nombre, re.sub(r"<[^>]*>", "", xml)) for nombre, xml in _fuentes_jinja(doc)]


def usos_fuera_del_bucle(doc: DocxTemplate, nombre: str):
    """Busca dónde se usa `nombre` fuera del {% for %} que lo define. Es el caso
    típico de "falta la variable e": una etiqueta {{e.algo}} que quedó suelta
    —muchas veces en un encabezado o en una fila copiada— y Jinja la ve como una
    variable de nivel superior. Devuelve una lista de textos ubicándolas."""
    hallazgos = []
    es_de_un_for = False
    for parte, texto in _partes_de_texto(doc):
        pila = []          # bucles abiertos, cada uno con sus variables
        for m in re.finditer(r"\{[{%].*?[%}]\}", texto, re.S):
            frag = m.group(0)
            etiq = _RE_ETIQUETA.match(frag)
            if etiq and etiq.group(1) == "for":
                objetivo = etiq.group(2).split(" in ")[0]
                pila.append({v.strip() for v in objetivo.split(",") if v.strip()})
                es_de_un_for |= nombre in pila[-1]
                continue
            if etiq and etiq.group(1) == "endfor":
                if pila:
                    pila.pop()
                continue
            en_alcance = any(nombre in marco for marco in pila)
            usa = (_RE_USO.match(frag) or [None, None])[1] if frag.startswith("{{") else None
            if not en_alcance and (usa == nombre or
                                   re.search(r"\b" + re.escape(nombre) + r"\b", frag)):
                contexto = texto[max(0, m.start() - 40):m.end() + 20].replace("\n", " ")
                hallazgos.append(f"en el {parte}: …{contexto.strip()}…")
    # Si ningún {% for %} la define, no es una variable de bucle suelta.
    return hallazgos if es_de_un_for else []


def diagnosticar_faltantes(doc: DocxTemplate, faltantes):
    """Para cada variable faltante que en realidad es la de un {% for %},
    explica dónde se usa fuera del bucle. Devuelve {variable: explicación}."""
    pistas = {}
    for nombre in faltantes:
        try:
            usos = usos_fuera_del_bucle(doc, nombre)
        except Exception:
            continue
        if usos:
            pistas[nombre] = (
                f"«{nombre}» es la variable de un {{% for {nombre} in ... %}}: no va en el "
                f"JSON. Hay {len(usos)} uso(s) fuera del bucle, que es lo que la deja "
                f"suelta — " + "; ".join(usos[:3])
            )
    return pistas


def _raise_syntaxis(doc: DocxTemplate, e: TemplateSyntaxError):
    """Construye un GenError descriptivo para un error de sintaxis Jinja,
    mostrando el fragmento del XML del documento donde está la etiqueta."""
    contexto = ""
    try:
        xml = doc.patch_xml(doc.get_xml())
        lineas = xml.splitlines()
        n = e.lineno or 0
        ini, fin = max(0, n - 2), min(len(lineas), n + 1)
        frag = "\n".join(
            f"    {'>' if i + 1 == n else ' '} {lineas[i][:200]}"
            for i in range(ini, fin)
        )
        if frag.strip():
            contexto = "\n  Fragmento del documento (XML interno):\n" + frag
    except Exception:
        pass
    raise GenError(
        "sintaxis de la plantilla (Jinja2)",
        e.message + contexto,
        ubicacion=f"línea {e.lineno} del XML interno del .docx",
        sugerencia="Revisá que cada {% ... %} tenga su cierre "
        "({% endfor %}, {% endif %}) y que las llaves estén balanceadas.",
    )


# --------------------------------------------------------------------------- #
# 4) Renderizar la plantilla con los datos
# --------------------------------------------------------------------------- #
def renderizar(doc: DocxTemplate, datos: dict, salida_docx: Path):
    env = Environment(undefined=StrictUndefined)
    try:
        doc.render(datos, jinja_env=env)
    except TemplateSyntaxError as e:
        _raise_syntaxis(doc, e)
    except UndefinedError as e:
        raise GenError(
            "renderizado (variable indefinida)",
            str(e),
            sugerencia="Falta ese dato en el JSON, o el nombre/atributo está mal escrito. "
            "Corré con --check para ver todas las variables que la plantilla espera.",
        )
    except Exception as e:
        raise GenError("renderizado", f"{type(e).__name__}: {e}")
    try:
        doc.save(str(salida_docx))
    except Exception as e:
        raise GenError("guardado del .docx intermedio", f"{type(e).__name__}: {e}")


# --------------------------------------------------------------------------- #
# 5) Convertir el .docx renderizado a PDF con LibreOffice headless
# --------------------------------------------------------------------------- #
# Dónde suele estar LibreOffice cuando no está en el PATH.
RUTAS_SOFFICE = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/lib/libreoffice/program/soffice",
    "/snap/bin/libreoffice",
]


def buscar_soffice(ruta: str | None = None) -> str:
    """Devuelve el ejecutable de LibreOffice a usar. Orden: la ruta que se pasa
    (--soffice), la variable de entorno GENDOC_SOFFICE, el PATH, y por último
    las ubicaciones habituales de cada sistema."""
    explicita = ruta or os.environ.get("GENDOC_SOFFICE")
    if explicita:
        if Path(explicita).is_file():
            return explicita
        hallado = shutil.which(explicita)
        if hallado:
            return hallado
        raise GenError(
            "conversión a PDF",
            f"La ruta indicada para LibreOffice no existe: {explicita}",
            sugerencia="Apuntá a soffice.exe (en Windows suele estar en "
            r"C:\Program Files\LibreOffice\program\soffice.exe).",
        )

    for nombre in ("soffice", "soffice.exe"):
        hallado = shutil.which(nombre)
        if hallado:
            return hallado
    for candidato in RUTAS_SOFFICE:
        if Path(candidato).is_file():
            return candidato

    raise GenError(
        "conversión a PDF",
        "No se encontró LibreOffice ('soffice') ni en el PATH ni en las rutas habituales.",
        sugerencia="Indicá la ruta con --soffice, o dejándola fija en la variable "
        "de entorno GENDOC_SOFFICE "
        r"(p. ej. C:\Program Files\LibreOffice\program\soffice.exe).",
    )


def a_pdf(docx_path: Path, pdf_path: Path, soffice: str | None = None):
    exe = buscar_soffice(soffice)
    # UserInstallation único evita el choque de locks si ya hay un soffice abierto.
    perfil = Path(tempfile.gettempdir()) / f"lo_{uuid.uuid4().hex}"
    with tempfile.TemporaryDirectory() as tmp:
        cmd = [
            exe, "--headless",
            f"-env:UserInstallation={perfil.as_uri()}",
            "--convert-to", "pdf", "--outdir", tmp, str(docx_path),
        ]
        # Windows a veces devuelve "archivo en uso" si soffice recién arrancó
        # o si el antivirus lo está mirando: reintentamos una vez.
        for intento in (1, 2):
            try:
                r = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
                break
            except FileNotFoundError:
                raise GenError("conversión a PDF", f"No se pudo ejecutar LibreOffice: {exe}",
                               sugerencia="Revisá la ruta (--soffice o GENDOC_SOFFICE).")
            except subprocess.TimeoutExpired:
                raise GenError("conversión a PDF", "LibreOffice tardó demasiado (timeout 180s).")
            except OSError as e:
                if intento == 1:
                    time.sleep(2)
                    continue
                raise GenError(
                    "conversión a PDF",
                    f"No se pudo lanzar LibreOffice: {type(e).__name__}: {e}",
                    ubicacion=exe,
                    sugerencia="Suele ser LibreOffice abierto o ocupado: cerralo "
                    "y volvé a intentar.",
                )
        generado = Path(tmp) / (docx_path.stem + ".pdf")
        if r.returncode != 0 or not generado.exists():
            raise GenError(
                "conversión a PDF",
                (r.stderr or r.stdout or "LibreOffice no devolvió un PDF.").strip(),
            )
        pdf_path.parent.mkdir(parents=True, exist_ok=True)
        generado.replace(pdf_path)


# --------------------------------------------------------------------------- #
# Formatos de salida: PDF, Word (.docx) o ambos
# --------------------------------------------------------------------------- #
FORMATOS = ("pdf", "docx", "ambos")


def generar(doc: DocxTemplate, datos: dict, formato: str, pdf_path: Path | None,
            docx_path: Path | None, soffice: str | None = None):
    """Renderiza la plantilla y deja el resultado en el/los formato(s) pedidos.
    El Word sale directo de docxtpl (no necesita LibreOffice); el PDF es ese
    mismo .docx convertido. Devuelve la lista de archivos generados."""
    if formato not in FORMATOS:
        raise GenError("entrada", f"Formato desconocido: {formato!r}",
                       sugerencia="Usá pdf, docx o ambos.")
    quiere_docx = formato in ("docx", "ambos")
    quiere_pdf = formato in ("pdf", "ambos")
    generados = []
    with tempfile.TemporaryDirectory() as tmp:
        render = docx_path if quiere_docx else Path(tmp) / "render.docx"
        render.parent.mkdir(parents=True, exist_ok=True)
        renderizar(doc, datos, render)
        if quiere_docx:
            generados.append(render)
        if quiere_pdf:
            a_pdf(render, pdf_path, soffice)
            generados.append(pdf_path)
    return generados


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def main():
    p = argparse.ArgumentParser(
        description="Genera un PDF y/o Word desde una plantilla .docx (Jinja2) + un JSON de datos."
    )
    p.add_argument("plantilla", type=Path, help="Plantilla Word (.docx) con etiquetas Jinja2")
    p.add_argument("datos", type=Path, help="Archivo JSON con los datos")
    p.add_argument("-o", "--salida", type=Path,
                   help="Archivo de salida (por defecto: <plantilla>.pdf). Si termina en "
                        ".docx y no se indica --formato, se genera Word.")
    p.add_argument("-f", "--formato", choices=FORMATOS,
                   help="pdf (por defecto), docx (Word, no necesita LibreOffice) o ambos")
    p.add_argument("--check", action="store_true",
                   help="Sólo valida plantilla + datos y lista las variables; no genera nada")
    p.add_argument("--keep-docx", action="store_true",
                   help="Igual que --formato ambos (se mantiene por compatibilidad)")
    p.add_argument("--soffice", metavar="RUTA",
                   help="Ruta al ejecutable de LibreOffice (si no está en el PATH). "
                        "También se puede fijar con la variable GENDOC_SOFFICE.")
    args = p.parse_args()

    try:
        datos, avisos = cargar_json(args.datos)
        for aviso in avisos:
            print("ℹ " + aviso)
        doc = cargar_plantilla(args.plantilla)
        an = analizar_variables(doc, datos)
        datos = an.contexto  # ya desanidado si el JSON venía envuelto

        if an.envoltorios:
            print("ℹ El JSON venía envuelto en "
                  + ", ".join(f"«{k}»" for k in an.envoltorios)
                  + "; se abrió ese nivel para resolver las variables.")

        if args.check:
            print(f"Variables que la plantilla espera ({len(an.esperadas)}): "
                  + (", ".join(an.esperadas) or "—"))
            print("Faltan en el JSON: "
                  + (", ".join(an.faltantes) if an.faltantes else "ninguna ✔"))
            for pista in diagnosticar_faltantes(doc, an.faltantes).values():
                print("  ↳ " + pista)
            for a in an.atributos:
                print("Atributo faltante: " + a)
            if an.opcionales:
                print("Opcionales que no vienen (se omiten sus bloques {% if %}): "
                      + ", ".join(an.opcionales))
            if an.sin_usar:
                print("En el JSON pero sin usar: " + ", ".join(an.sin_usar))
            if an.faltantes or an.atributos:
                sys.exit(1)
            print("\n✔ La plantilla y los datos son compatibles.")
            return

        # Formato: el explícito, o lo que diga la extensión de -o, o PDF.
        formato = args.formato
        if formato is None:
            if args.keep_docx:
                formato = "ambos"
            elif args.salida and args.salida.suffix.lower() == ".docx":
                formato = "docx"
            else:
                formato = "pdf"

        base = args.salida or args.plantilla
        pdf_out = base.with_suffix(".pdf")
        docx_out = base.with_suffix(".docx")
        # Nunca pisar la plantilla con el Word generado.
        if docx_out.resolve() == args.plantilla.resolve():
            docx_out = base.with_name(base.stem + "_generado.docx")

        if an.atributos:
            raise GenError(
                "datos incompletos",
                "La plantilla usa atributos que el JSON no trae:\n"
                + "\n".join("  • " + a for a in an.atributos),
                sugerencia="Agregalos al JSON, o corregí/sacá esa etiqueta en el Word.",
            )
        for archivo in generar(doc, datos, formato, pdf_out, docx_out, args.soffice):
            print(f"✔ Generado: {archivo}")

    except GenError as e:
        e.imprimir()
        sys.exit(1)


if __name__ == "__main__":
    main()
