#!/usr/bin/env python3
"""Genera fragmentos Quarto desde la matriz maestra del sitio."""

from __future__ import annotations

from datetime import date, datetime
from html import escape
from pathlib import Path
import math
import re

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
MATRIZ = ROOT / "Matriz_contenido_web_Fidel_Vallejo.xlsx"
SALIDA = ROOT / "_generated"


def limpio(valor):
    if valor is None:
        return ""
    if isinstance(valor, float) and math.isnan(valor):
        return ""
    return str(valor).strip()


def es_si(valor):
    return limpio(valor).casefold() in {"si", "sí", "yes", "true", "1"}


def numero(valor):
    if isinstance(valor, bool):
        return str(valor)
    if isinstance(valor, int):
        return str(valor)
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return limpio(valor)


def registros(libro, hoja):
    ws = libro[hoja]
    fila_encabezados = None
    encabezados = []
    for numero_fila, celdas in enumerate(ws.iter_rows(min_row=1, max_row=12), start=1):
        candidatos = [limpio(c.value) for c in celdas]
        if "ID" in candidatos and "Publicar" in candidatos:
            fila_encabezados = numero_fila
            encabezados = candidatos
            break
    if fila_encabezados is None:
        raise ValueError(f"No se encontró la fila de encabezados en la hoja {hoja!r}")
    filas = []
    for valores in ws.iter_rows(min_row=fila_encabezados + 1, values_only=True):
        fila = dict(zip(encabezados, valores))
        if not any(limpio(v) for v in valores):
            continue
        if "Publicar" in fila and not es_si(fila["Publicar"]):
            continue
        filas.append(fila)
    return filas


def orden(fila):
    valor = fila.get("Orden")
    try:
        return float(valor)
    except (TypeError, ValueError):
        return float("inf")


def e(valor):
    return escape(limpio(valor), quote=True)


def enlace(url, texto=None):
    url = limpio(url)
    if not url:
        return ""
    return f"[{texto or url}]({url})"


def escribir(nombre, contenido):
    SALIDA.mkdir(parents=True, exist_ok=True)
    (SALIDA / nombre).write_text(contenido.rstrip() + "\n", encoding="utf-8")


def bloque_metricas(filas, clase):
    partes = [f"::: {{.{clase}}}"]
    for fila in sorted(filas, key=orden):
        valor = f"{limpio(fila.get('Prefijo'))}{numero(fila.get('Valor'))}{limpio(fila.get('Sufijo'))}"
        etiqueta = limpio(fila.get("Indicador"))
        if etiqueta:
            etiqueta = etiqueta[0].lower() + etiqueta[1:]
        partes.extend([
            "::: {}",
            f"<strong>{e(valor)}</strong>",
            f"<span>{e(etiqueta)}</span>",
            ":::",
        ])
    partes.append(":::")
    return "\n".join(partes)


def periodo(inicio, fin, actual):
    meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
             "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

    def fecha_corta(valor):
        if isinstance(valor, (datetime, date)):
            return f"{meses[valor.month - 1]} de {valor.year}"
        return limpio(valor)

    ini = fecha_corta(inicio)
    termino = "actualidad" if es_si(actual) else fecha_corta(fin)
    return "–".join(x for x in [ini, termino] if x)


def principal():
    if not MATRIZ.exists():
        raise SystemExit(f"No se encontró la matriz: {MATRIZ}")

    libro = load_workbook(MATRIZ, data_only=True, read_only=True)
    indicadores = registros(libro, "Indicadores")
    publicaciones = sorted(registros(libro, "Publicaciones"), key=orden)
    proyectos = sorted(registros(libro, "Proyectos"), key=orden)
    experiencia = sorted(registros(libro, "Experiencia"), key=orden)
    herramientas = sorted(registros(libro, "Herramientas"), key=orden)
    vinculos = registros(libro, "Enlaces y contacto")
    textos = registros(libro, "Textos")

    def indicadores_de(seccion):
        return [f for f in indicadores if limpio(f.get("Sección")).casefold() == seccion.casefold()]

    inicio = indicadores_de("Inicio")
    inicios = [f.get("Inicio") for f in experiencia if isinstance(f.get("Inicio"), (date, datetime))]
    if inicios:
        inicio.append({
            "Orden": 999,
            "Valor": f"{min(x.year for x in inicios)}–actualidad",
            "Indicador": "Experiencia en proyectos ambientales",
        })
    escribir("indicadores_inicio.md", bloque_metricas(inicio, "metrics-strip"))
    escribir("indicadores_publicaciones.md", bloque_metricas(indicadores_de("Publicaciones"), "publication-metrics"))
    escribir("indicadores_proyectos.md", bloque_metricas(indicadores_de("Proyectos"), "project-metrics"))

    etiquetas = []
    for fila in herramientas:
        if es_si(fila.get("Mostrar en inicio")):
            etiquetas.append(f"<span>{e(fila.get('Herramienta'))}</span>")
    escribir("herramientas_inicio.md", "::: {.toolkit-tags}\n" + "".join(etiquetas) + "\n:::")

    pubs = []
    for fila in publicaciones:
        titulo_limpio = limpio(fila.get("Título"))
        titulo = e(titulo_limpio)
        cierre_titulo = "" if titulo_limpio.endswith((".", "?", "!")) else "."
        revista = e(fila.get("Revista"))
        anio = numero(fila.get("Año"))
        detalle = e(fila.get("Volumen / número / páginas"))
        referencia = ", ".join(x for x in [anio, detalle] if x)
        url = limpio(fila.get("URL")) or ("https://doi.org/" + limpio(fila.get("DOI")) if limpio(fila.get("DOI")) else "")
        pubs.extend([
            "::: {.publication-item}",
            f"**{titulo}{cierre_titulo}**  ",
            f"*{revista}*{', ' if referencia else ''}{e(referencia)}.  ",
            enlace(url),
            ":::",
            "",
        ])
    escribir("publicaciones.md", "\n".join(pubs))

    enlaces_por_etiqueta = {limpio(f.get("Etiqueta")): limpio(f.get("Valor / URL")) for f in vinculos}
    botones = ["::: {.academic-links}"]
    clases = {"Scopus": "btn-primary", "ResearchGate": "btn-outline-primary", "ORCID": "btn-outline-secondary"}
    for etiqueta in ["Scopus", "ResearchGate", "ORCID"]:
        url = enlaces_por_etiqueta.get(etiqueta, "")
        if url:
            botones.append(f"[{etiqueta}]({url}){{.btn .{clases[etiqueta]}}}")
    botones.append(":::")
    escribir("enlaces_academicos.md", "\n".join(botones))

    proys = []
    tipo_anterior = None
    for fila in proyectos:
        tipo = limpio(fila.get("Tipo")) or "Otros proyectos"
        if tipo != tipo_anterior:
            proys.extend([f"## {e(tipo)}", ""])
            tipo_anterior = tipo
        nombre = limpio(fila.get("Nombre"))
        codigo = limpio(fila.get("Código"))
        titulo = f"{nombre} — {codigo}" if codigo else nombre
        destacado = limpio(fila.get("Observaciones")).casefold() == "proyecto destacado"
        if destacado:
            proys.append("::: {.project-card .project-featured}")
            tecnologias = " · ".join(x.strip() for x in limpio(fila.get("Tecnologías")).split(";") if x.strip()).upper()
            if tecnologias:
                proys.append(f'<p class="eyebrow">{e(tecnologias)}</p>')
            proys.append("")
        proys.extend([f"### {e(titulo)}", "", e(fila.get("Descripción pública")), ""])
        url = limpio(fila.get("URL"))
        if url:
            proys.extend([f"[Repositorio en GitHub]({url}){{.btn .btn-primary}}", ""])
        if destacado:
            proys.extend([":::", ""])
    escribir("proyectos.md", "\n".join(proys))

    exp = ["::: {.timeline}"]
    for fila in experiencia:
        exp.extend([
            "::: {.timeline-item}",
            f"### {e(fila.get('Cargo'))}",
            "",
            f"**{e(fila.get('Organización'))} · {e(periodo(fila.get('Inicio'), fila.get('Fin'), fila.get('Actual')))}**",
            "",
            e(fila.get("Descripción pública")),
            ":::",
            "",
        ])
    exp.append(":::")
    escribir("experiencia.md", "\n".join(exp))

    textos_map = {(limpio(f.get("Página")), limpio(f.get("Campo"))): limpio(f.get("Texto")) for f in textos}
    for pagina, campo, archivo in [
        ("Inicio", "Investigación", "texto_inicio_investigacion.md"),
        ("Publicaciones", "Introducción", "texto_publicaciones_intro.md"),
        ("Proyectos", "Introducción", "texto_proyectos_intro.md"),
        ("Perfil", "Síntesis", "texto_perfil_sintesis.md"),
    ]:
        escribir(archivo, textos_map.get((pagina, campo), ""))

    print(f"Contenido generado desde {MATRIZ.name}: {len(publicaciones)} publicaciones, {len(proyectos)} proyectos.")


if __name__ == "__main__":
    principal()
# La publicación automática se activa al actualizar la matriz.
