#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
codigo_limpieza.py
==================
Auditoría y limpieza reproducible de RespuestasFormsOriginales.xlsx
(Proyecto de personalidad - Sistemas Inteligentes, UNAM).

Principios (según las reglas del proyecto):
  * Solo reglas explícitas, operaciones deterministas y expresiones regulares.
    No se usa NLP, IA ni análisis semántico.
  * El archivo original NUNCA se modifica (solo se lee).
  * Cada hoja se trabaja por separado. NO se unen formularios.
  * No se eliminan duplicados, no se imputa y no se sustituyen valores atípicos:
    solo se marcan para revisión humana.
  * Se conservan las respuestas originales; las versiones homologadas van en
    columnas nuevas.
  * Se retiran identificadores directos (correo, número de cuenta, marca de
    tiempo) y se crean IDs anónimos independientes por hoja (P###, M###, R###).

Uso:
    python codigo_limpieza.py
    python codigo_limpieza.py --entrada RespuestasFormsOriginales.xlsx --salida salida

Dependencias: pandas, numpy, openpyxl  (ver README / requirements.txt)

Salidas (en la carpeta --salida):
    Datos_Limpios.xlsx, Bitacora_Limpieza.xlsx, Reporte_Auditoria.xlsx
"""
import argparse
import re
import sys
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


HOJA_P = "RespuestasTestPersonalidad"
HOJA_M = "RespuestasMusica"
HOJA_R = "RespuestasRelacionesAmorosas"

RE_CORREO = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")
RE_CUENTA = re.compile(r"^\d{9}$")
RE_TEL = re.compile(r"(?<!\d)\d{8,}(?!\d)")

# ----------------------------------------------------------------------------
# Utilidades generales
# ----------------------------------------------------------------------------


def sin_acentos(texto):
    t = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in t if not unicodedata.combining(c))


def norm(x):
    """Minúsculas, sin acentos, sin puntuación, espacios colapsados."""
    if pd.isna(x):
        return ""
    s = sin_acentos(x).lower()
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def luz(x):
    """Normalización ligera (conserva dígitos y signos +, -, /, ',', '.')."""
    s = sin_acentos(x).lower()
    s = re.sub(r"\s+", " ", s).strip()
    return s.rstrip(". ")


def en_blanco(x):
    return bool(pd.isna(x) or (isinstance(x, str) and x.strip() == ""))


def mascara_blancos(serie):
    return serie.map(en_blanco)


def mascara_iqr(serie):
    """True donde el valor cae fuera de [Q1-1.5*IQR, Q3+1.5*IQR]."""
    s = pd.to_numeric(serie, errors="coerce")
    if s.dropna().shape[0] < 4:
        return pd.Series(False, index=serie.index)
    q1, q3 = s.quantile(0.25), s.quantile(0.75)
    iqr = q3 - q1
    return (s < q1 - 1.5 * iqr) | (s > q3 + 1.5 * iqr)


class Registro:
    """Acumula bitácora, hallazgos y tablas de auditoría."""

    def __init__(self):
        self.bitacora = []
        self.atipicos = []
        self.revision = []
        self.homologaciones = []
        self.exclusiones = []
        self.duplicados = []
        self.calidad_id = []
        self.validaciones = []
        self.resumen = []
        self.columnas = []
        self.faltantes = []

    def log(self, hoja, transformacion, columna, regla, n, justificacion):
        self.bitacora.append(
            {
                "n": len(self.bitacora) + 1,
                "hoja": hoja,
                "transformacion": transformacion,
                "columna_afectada": columna,
                "regla_aplicada": regla,
                "registros_afectados": int(n),
                "justificacion": justificacion,
            }
        )

    def atipico(self, hoja, id_, columna, valor, motivo):
        self.atipicos.append(
            {"hoja": hoja, "id_participante": id_, "columna": columna, "valor": valor, "motivo": motivo}
        )

    def a_revision(self, hoja, id_, columna, motivo, sugerencia):
        self.revision.append(
            {
                "hoja": hoja,
                "id_participante": id_,
                "columna": columna,
                "motivo": motivo,
                "accion_sugerida": sugerencia,
            }
        )

    def validar(self, nombre, ok, detalle=""):
        self.validaciones.append({"validacion": nombre, "resultado": "OK" if ok else "FALLA", "detalle": detalle})


def asignar_ids(n, prefijo):
    return [f"{prefijo}{i:03d}" for i in range(1, n + 1)]


def cuenta_texto(serie):
    """Número de cuenta como texto de dígitos (sin '.0')."""
    def f(v):
        if pd.isna(v):
            return ""
        if isinstance(v, (int, np.integer)):
            return str(int(v))
        if isinstance(v, (float, np.floating)):
            return str(int(v)) if float(v).is_integer() else str(v)
        return re.sub(r"\D", "", str(v))
    return serie.map(f)


def revisar_texto_abierto(df, columnas_orig, ids, hoja, reg, posiciones):
    """
    Marca (sin modificar) respuestas abiertas poco informativas y posible
    información personal dentro del texto. Devuelve la columna resumen.
    """
    patron_poco = re.compile(r"^(n ?a|ns|no se|nada|ninguno|ninguna|x+)$")
    resumen = []
    for i in df.index:
        marcadas = []
        for c in columnas_orig:
            v = df.at[i, c]
            if en_blanco(v):
                continue
            t = str(v).strip()
            n = norm(t)
            if len(t) <= 3 or patron_poco.fullmatch(n) or re.fullmatch(r"[\W_]+", t):
                marcadas.append(posiciones[c])
                reg.a_revision(hoja, ids[i], f"col_{posiciones[c]}",
                               "Respuesta abierta muy corta o poco informativa (regla: <=3 caracteres o patrón N/A, ns, nada, ninguno).",
                               "Confirmar si es una respuesta válida; no se modificó.")
            if RE_CORREO.search(t) or RE_TEL.search(t):
                reg.a_revision(hoja, ids[i], f"col_{posiciones[c]}",
                               "El texto contiene algo parecido a un correo o a un número largo (posible dato personal).",
                               "Revisar y, si procede, anonimizar manualmente; no se modificó.")
        resumen.append("; ".join(f"col_{m}" for m in marcadas) if marcadas else "")
    return resumen


# ----------------------------------------------------------------------------
# Hoja de personalidad
# ----------------------------------------------------------------------------


def limpiar_personalidad(raw, reg):
    hoja = HOJA_P
    cols = list(raw.columns)
    assert len(cols) == 18, f"{hoja}: se esperaban 18 columnas y hay {len(cols)}"
    c_ts, c_mail, c_cta, c_gen = cols[:4]
    c_open = cols[4:]
    pos = {c: i + 1 for i, c in enumerate(cols)}
    n0 = len(raw)

    # --- calidad de identificadores (solo conteos) ---
    cta = cuenta_texto(raw[c_cta])
    mail = raw[c_mail].astype(str).str.strip().str.lower()
    reg.calidad_id.append({"hoja": hoja, "revision": "Número de cuenta con formato de 9 dígitos",
                           "registros_ok": int(cta.map(lambda v: bool(RE_CUENTA.match(v))).sum()), "total": n0})
    reg.calidad_id.append({"hoja": hoja, "revision": "Correo con formato válido",
                           "registros_ok": int(mail.map(lambda v: bool(RE_CORREO.fullmatch(v))).sum()), "total": n0})
    prefijo = mail.str.extract(r"^(\d+)@")[0].fillna("")
    reg.calidad_id.append({"hoja": hoja, "revision": "Prefijo del correo coincide con el número de cuenta",
                           "registros_ok": int((prefijo == cta).sum()), "total": n0})

    ids = asignar_ids(n0, "P")
    raw = raw.copy()
    raw.index = range(n0)

    # --- posibles duplicados (solo marcar) ---
    grupo = pd.Series("", index=raw.index)
    excluir = pd.Series(False, index=raw.index)  # decisión del equipo sobre duplicados
    k = 0
    for clave, nombre in ((cta, "mismo número de cuenta"), (mail, "mismo correo")):
        for valor, idx in raw.groupby(clave).groups.items():
            if valor != "" and len(idx) > 1 and (grupo.loc[list(idx)] == "").all():
                k += 1
                etiqueta = f"DUP-P{k:02d}"
                grupo.loc[list(idx)] = etiqueta
                filas = list(idx)
                conservar = raw.loc[filas, c_ts].idxmax()  # el envío más reciente
                for f in filas:
                    if f != conservar:
                        excluir[f] = True
                for a in range(len(filas)):
                    for b in range(a + 1, len(filas)):
                        ra, rb = raw.loc[filas[a], c_open + [c_gen]], raw.loc[filas[b], c_open + [c_gen]]
                        iguales = int(sum(1 for c in c_open + [c_gen]
                                          if str(ra[c]).strip() == str(rb[c]).strip()))
                        dif_min = abs((raw.loc[filas[a], c_ts] - raw.loc[filas[b], c_ts]).total_seconds()) / 60
                        reg.duplicados.append({
                            "hoja": hoja, "grupo": etiqueta, "id_a": ids[filas[a]], "id_b": ids[filas[b]],
                            "motivo": f"Coinciden en: {nombre} (el valor no se publica).",
                            "columnas_de_respuesta_identicas": f"{iguales} de {len(c_open) + 1}",
                            "minutos_entre_envios": round(dif_min, 1),
                            "accion": f"Decisión del equipo: se conserva {ids[conservar]} (el envío más reciente) y se excluye el anterior. "
                                  "El equipo confirmó que fue una persona que respondió dos veces (la primera al crear el formulario).",
                        })
                        reg.a_revision(hoja, ids[filas[a]], "toda la fila",
                                       f"Posible duplicado ({etiqueta}) con {ids[filas[b]]}.",
                                       "Resuelto por decisión del equipo: se conserva el envío más reciente.")
    dup_resp = raw.duplicated(subset=c_open + [c_gen], keep=False)
    if dup_resp.any():
        for i in raw.index[dup_resp & (grupo == "")]:
            reg.a_revision(hoja, ids[i], "toda la fila",
                           "Respuestas idénticas a otra fila (sin coincidir identificadores).", "Revisar.")
    reg.log(hoja, "Detección de posibles duplicados", "número de cuenta / correo",
            "Agrupar por valor idéntico; solo se marca, no se elimina", (grupo != "").sum(),
            "Regla 3: no eliminar duplicados automáticamente (aquí solo se detecta y se marca).")
    if excluir.any():
        reg.exclusiones.append({"hoja": hoja,
                                "motivo": "Respuesta repetida de la misma persona (la primera, enviada al crear el formulario). "
                                          "Decisión del equipo: se conserva el envío más reciente.",
                                "registros_excluidos": int(excluir.sum())})
        reg.log(hoja, "Exclusión de respuesta duplicada (decisión del equipo)", "toda la fila",
                "Misma persona con dos envíos: se conserva el de marca de tiempo más reciente; la otra fila se excluye",
                int(excluir.sum()), "Decisión humana documentada, no automática. Los IDs se asignaron antes de excluir, "
                "por lo que el ID de la fila excluida no vuelve a usarse.")

    # --- homologación de género ---
    def homol_genero(x):
        if en_blanco(x):
            return np.nan
        t = norm(x)
        if re.fullmatch(r"(masculino|hombre|m)", t):
            return "Masculino"
        if re.fullmatch(r"(femenino|mujer|f)", t):
            return "Femenino"
        return np.nan

    genero_h = raw[c_gen].map(homol_genero)
    no_homol = genero_h.isna() & ~mascara_blancos(raw[c_gen])
    for i in raw.index[no_homol]:
        reg.a_revision(hoja, ids[i], f"col_{pos[c_gen]}", "Género no reconocido por la regla.", "Revisar manualmente.")
    conteo = (pd.DataFrame({"o": raw[c_gen].astype(str), "h": genero_h})
              .groupby(["o", "h"], dropna=False).size().reset_index(name="n"))
    for _, r in conteo.iterrows():
        reg.homologaciones.append({"hoja": hoja, "columna": "género", "valor_original": repr(r["o"]),
                                   "valor_homologado": r["h"], "registros": int(r["n"])})
    cambiados = int((genero_h.astype(str) != raw[c_gen].astype(str)).sum())
    reg.log(hoja, "Homologación de categorías", "género → genero_homologado",
            "norm(texto) ∈ {masculino, hombre, m} → Masculino; ∈ {femenino, mujer, f} → Femenino; resto → vacío y a revisión",
            cambiados, "Variantes de escritura inequívocas. La columna original se conserva intacta.")

    # --- revisión de texto abierto ---
    revis_sub = revisar_texto_abierto(raw.loc[~excluir], c_open, ids, hoja, reg, pos)
    revis_s = pd.Series("", index=raw.index, dtype=object)
    revis_s.loc[~excluir] = revis_sub
    revis = revis_s.tolist()
    reg.log(hoja, "Marcado de respuestas poco informativas / posible dato personal", "preguntas abiertas",
            "Longitud <=3, patrones N/A|ns|no sé|nada|ninguno, o regex de correo/teléfono; solo se marca",
            sum(1 for r in revis if r), "Regla 8: no interpretar; las dudas van a revisión humana.")

    # --- construcción de salida ---
    limpio = pd.DataFrame({"id_participante": ids})
    limpio[c_gen] = raw[c_gen].values
    limpio["genero_homologado"] = genero_h.values
    for c in c_open:
        limpio[c] = raw[c].values
    limpio["revision_respuestas_cortas"] = revis
    limpio["grupo_posible_duplicado"] = [g if g else np.nan for g in grupo]
    limpio = limpio.loc[~excluir.values].reset_index(drop=True)
    reg.log(hoja, "Retiro de identificadores", "marca temporal, correo, número de cuenta",
            "Columnas eliminadas del conjunto analítico", n0 * 3,
            "Regla 4: sin identificadores directos. Se creó id_participante (P###).")
    reg.log(hoja, "Creación de ID anónimo", "id_participante", "P + número según el orden de filas de ESTA hoja (asignado antes de excluir)", n0,
            "Regla 5: ID independiente por hoja; no permite relacionar personas entre formularios.")

    retiradas = [c_ts, c_mail, c_cta]
    meta = {"retiradas": retiradas, "original": [c_gen] + c_open, "derivadas": ["genero_homologado",
            "revision_respuestas_cortas", "grupo_posible_duplicado"], "pos": pos, "n0": n0, "excluidos": int(excluir.sum()),
            "indices_conservados": list(raw.index[~excluir])}
    return limpio, meta


# ----------------------------------------------------------------------------
# Hoja de música (tiempo de escucha)
# ----------------------------------------------------------------------------

_NUM = r"(\d+(?:[.,]\d+)?)"
_RE_MIN = re.compile(_NUM + r"\s*(?:(?:horas?|hrs?|h)\b)?\s*(?:o\s+mas|\+|a\s+mas|en\s+adelante)")
_RE_INT = re.compile(_NUM + r"\s*(?:-|–|—|a|o|y)\s*" + _NUM)
_RE_UNIDAD = re.compile(r"\b(horas?|hrs?|h)\b")
_RE_APROX = re.compile(r"\b(unas?|unos|aprox\w*|como|casi|alrededor|cerca de|mas o menos)\b")


def _f(s):
    return float(s.replace(",", "."))


def parsear_tiempo(valor, fila_texto_p1="", fila_texto_p4=""):
    """Devuelve dict con min, max, tipo, contextual, revisar, nota. Determinista."""
    base = {"min": np.nan, "max": np.nan, "tipo": "", "contextual": "No", "revisar": "No", "nota": ""}
    if en_blanco(valor):
        base.update(tipo="faltante", nota="Sin respuesta.")
        return base
    t = luz(valor)
    notas = []
    nums = re.findall(_NUM, t)
    aprox = bool(_RE_APROX.search(t))
    tiene_unidad = bool(_RE_UNIDAD.search(t))

    if not nums:
        base.update(tipo="ambiguo_sin_convertir", revisar="Sí",
                    nota="Sin cifra en dígitos (cantidad escrita con letras o respuesta descriptiva). No se convierte; revisar manualmente.")
        return base
    if "minut" in t:
        base.update(tipo="ambiguo_sin_convertir", revisar="Sí", nota="Unidad en minutos: revisar manualmente.")
        return base

    m_min = _RE_MIN.search(t)
    m_int = _RE_INT.search(t)
    if m_min and len(nums) == 1:
        base.update(min=_f(m_min.group(1)), max=np.nan, tipo="minimo_sin_maximo")
        if aprox:
            notas.append("Expresión aproximada en el texto.")
    elif m_int and len(nums) == 2:
        lo, hi = _f(m_int.group(1)), _f(m_int.group(2))
        if lo > hi:
            base.update(tipo="ambiguo_sin_convertir", revisar="Sí", nota="Intervalo con mínimo mayor que máximo.")
            return base
        base.update(min=lo, max=hi, tipo="intervalo_aproximado" if aprox else "intervalo")
    elif len(nums) == 1:
        v = _f(nums[0])
        base.update(min=v, max=v, tipo="puntual_aproximado" if aprox else "puntual")
    else:
        base.update(tipo="ambiguo_sin_convertir", revisar="Sí",
                    nota="Varias cifras que no corresponden a un intervalo claro. Revisar.")
        return base

    if not tiene_unidad:
        base["contextual"] = "Sí"
        notas.append("Sin unidad escrita: se interpreta como horas por el contexto de la pregunta.")
    if base["tipo"].startswith("puntual"):
        notas.append("Valor puntual: mínimo y máximo se registran iguales.")
    top = base["max"] if not pd.isna(base["max"]) else base["min"]
    if base["min"] == 0 and base["max"] == 0:
        base["revisar"] = "Sí"
        notas.append("Respuesta 0: verificar si significa cero horas u otra cosa.")
        if re.search(r"\bno escucho\b", luz(fila_texto_p1)) or re.search(r"\bno escucho\b", luz(fila_texto_p4)):
            notas.append("Otras respuestas de la misma fila (col_3/col_6) dicen 'no escucho', lo cual es consistente con 0 h; aun así verificar.")
    if not pd.isna(top) and top > 12:
        base["revisar"] = "Sí"
        notas.append("Más de 12 h/día: valor sospechoso (no se modificó).")
    base["nota"] = " ".join(notas)
    return base


def limpiar_musica(raw, reg):
    hoja = HOJA_M
    cols = list(raw.columns)
    assert len(cols) == 18, f"{hoja}: se esperaban 18 columnas y hay {len(cols)}"
    c_ts, c_cta = cols[:2]
    c_open = cols[2:]
    c_tiempo = cols[3]
    c_p1, c_p4 = cols[2], cols[5]
    pos = {c: i + 1 for i, c in enumerate(cols)}
    n0 = len(raw)
    raw = raw.copy()
    raw.index = range(n0)
    ids = asignar_ids(n0, "M")

    cta = cuenta_texto(raw[c_cta])
    reg.calidad_id.append({"hoja": hoja, "revision": "Número de cuenta con formato de 9 dígitos",
                           "registros_ok": int(cta.map(lambda v: bool(RE_CUENTA.match(v))).sum()), "total": n0})

    # duplicados
    grupo = pd.Series("", index=raw.index)
    k = 0
    for valor, idx in raw.groupby(cta).groups.items():
        if valor != "" and len(idx) > 1:
            k += 1
            grupo.loc[list(idx)] = f"DUP-M{k:02d}"
            for i in idx:
                reg.a_revision(hoja, ids[i], "toda la fila", "Posible duplicado (mismo número de cuenta).", "Revisar.")
            reg.duplicados.append({"hoja": hoja, "grupo": f"DUP-M{k:02d}", "id_a": ids[list(idx)[0]],
                                   "id_b": ids[list(idx)[1]], "motivo": "Mismo número de cuenta.",
                                   "columnas_de_respuesta_identicas": "", "minutos_entre_envios": "",
                                   "accion": "NO se eliminó."})
    dup_resp = raw.duplicated(subset=c_open, keep=False)
    for i in raw.index[dup_resp & (grupo == "")]:
        reg.a_revision(hoja, ids[i], "toda la fila", "Respuestas idénticas a otra fila.", "Revisar.")
    reg.log(hoja, "Detección de posibles duplicados", "número de cuenta / respuestas",
            "Agrupar por cuenta idéntica; solo se marca", (grupo != "").sum(), "Regla 3: no eliminar duplicados automáticamente.")

    # tiempo de escucha
    res = [parsear_tiempo(raw.at[i, c_tiempo], raw.at[i, c_p1], raw.at[i, c_p4]) for i in raw.index]
    tdf = pd.DataFrame(res)
    for j, i in enumerate(raw.index):
        if tdf.at[j, "revisar"] == "Sí":
            reg.a_revision(hoja, ids[i], f"col_{pos[c_tiempo]}", tdf.at[j, "nota"], "Verificar con la persona o decidir criterio; no se forzó la conversión.")
            reg.atipico(hoja, ids[i], "tiempo de escucha diario", str(raw.at[i, c_tiempo]), tdf.at[j, "nota"])
    # atípicos estadísticos (IQR) sobre valores puntuales
    puntual = tdf["tipo"].str.startswith("puntual")
    iqr = mascara_iqr(tdf["min"].where(puntual))
    for j in tdf.index[iqr]:
        if tdf.at[j, "revisar"] != "Sí":
            reg.atipico(hoja, ids[j], "tiempo de escucha diario", str(raw.at[j, c_tiempo]),
                        "Atípico estadístico (regla 1.5×IQR sobre valores puntuales).")
    for tipo, n in tdf["tipo"].value_counts().items():
        reg.homologaciones.append({"hoja": hoja, "columna": "tiempo de escucha → tipo_registro_tiempo",
                                   "valor_original": "(ver tiempo_original)", "valor_homologado": tipo, "registros": int(n)})
    reg.log(hoja, "Extracción numérica del tiempo de escucha", "tiempo_min_horas, tiempo_max_horas, tipo_registro_tiempo",
            "Regex determinista: puntual → min=max; intervalo → min y max (sin promedio); 'o más'/'+' → min y max vacío; "
            "'unas/aprox/como' → tipo aproximado; sin cifra en dígitos → ambiguo sin convertir",
            (tdf["tipo"] != "ambiguo_sin_convertir").sum() - (tdf["tipo"] == "faltante").sum(),
            "Instrucciones específicas de la hoja de música. Se conserva el texto original.")
    reg.log(hoja, "Marcado de interpretación contextual", "tiempo_interpretacion_contextual",
            "Sin unidad escrita → se asume 'horas' por contexto y se marca 'Sí'", (tdf["contextual"] == "Sí").sum(),
            "Respuestas como '8' o '1 o 2' no traen unidad.")
    reg.log(hoja, "Marcado de valores a revisar", "tiempo_requiere_revision",
            "Valor 0, >12 h/día, o texto ambiguo → 'Sí'; no se sustituyen", (tdf["revisar"] == "Sí").sum(),
            "Regla 9: no eliminar atípicos; señalarlos.")

    revis = revisar_texto_abierto(raw, [c for c in c_open if c != c_tiempo], ids, hoja, reg, pos)
    reg.log(hoja, "Marcado de respuestas poco informativas / posible dato personal", "preguntas abiertas",
            "Longitud <=3, patrones N/A|ns|no sé|nada|ninguno, o regex de correo/teléfono; solo se marca",
            sum(1 for r in revis if r), "Regla 8.")

    limpio = pd.DataFrame({"id_participante": ids})
    for c in c_open:
        limpio[c] = raw[c].values
        if c == c_tiempo:
            limpio["tiempo_original"] = raw[c].values
            limpio["tiempo_min_horas"] = tdf["min"].values
            limpio["tiempo_max_horas"] = tdf["max"].values
            limpio["tipo_registro_tiempo"] = tdf["tipo"].values
            limpio["tiempo_interpretacion_contextual"] = tdf["contextual"].values
            limpio["tiempo_requiere_revision"] = tdf["revisar"].values
            limpio["nota_limpieza_tiempo"] = tdf["nota"].values
    limpio["revision_respuestas_cortas"] = revis
    limpio["grupo_posible_duplicado"] = [g if g else np.nan for g in grupo]
    reg.log(hoja, "Retiro de identificadores", "marca temporal, número de cuenta", "Columnas eliminadas", n0 * 2,
            "Regla 4.")
    reg.log(hoja, "Creación de ID anónimo", "id_participante", "M + número consecutivo según el orden de filas de ESTA hoja", n0,
            "Regla 5: independiente de P### y R###.")
    der = ["tiempo_original", "tiempo_min_horas", "tiempo_max_horas", "tipo_registro_tiempo",
           "tiempo_interpretacion_contextual", "tiempo_requiere_revision", "nota_limpieza_tiempo",
           "revision_respuestas_cortas", "grupo_posible_duplicado"]
    meta = {"retiradas": [c_ts, c_cta], "original": c_open, "derivadas": der, "pos": pos, "n0": n0, "excluidos": 0}
    return limpio, meta


# ----------------------------------------------------------------------------
# Hoja de relaciones amorosas
# ----------------------------------------------------------------------------


def limpiar_amor(raw, reg):
    hoja = HOJA_R
    cols = list(raw.columns)
    assert len(cols) == 21, f"{hoja}: se esperaban 21 columnas y hay {len(cols)}"
    c_ts, c_mail, c_cons, c_area, c_carr, c_sem, c_par, c_edad = cols[:8]
    c_dif = cols[15]
    c_open = cols[8:]
    pos = {c: i + 1 for i, c in enumerate(cols)}
    n0 = len(raw)
    raw = raw.copy()
    raw.index = range(n0)

    # --- consentimiento ---
    cons = raw[c_cons].map(norm)
    si = cons.eq("si")
    n_no = int(cons.eq("no").sum())
    n_otro = int((~si & ~cons.eq("no")).sum())
    reg.exclusiones.append({"hoja": hoja, "motivo": "No dio consentimiento (respondió 'No' a participar en el estudio)",
                            "registros_excluidos": n_no})
    if n_otro:
        reg.exclusiones.append({"hoja": hoja, "motivo": "Consentimiento vacío o no interpretable", "registros_excluidos": n_otro})
    reg.log(hoja, "Exclusión por falta de consentimiento", c_cons,
            "Se conservan solo las filas cuyo consentimiento normalizado es 'sí'", (~si).sum(),
            "Regla específica de la hoja. La fila excluida no se reproduce en ningún entregable.")
    n_vacias_excl = int(raw.loc[~si].drop(columns=[c_ts, c_mail, c_cons]).isna().all(axis=1).sum())
    reg.exclusiones.append({"hoja": hoja, "motivo": f"Detalle: {n_vacias_excl} de las filas excluidas no contenía ninguna otra respuesta",
                            "registros_excluidos": 0})

    corr = raw.index[si]
    r = raw.loc[corr].copy().reset_index(drop=True)
    ids = asignar_ids(len(r), "R")

    # --- calidad de identificadores ---
    mail = raw[c_mail].astype(str).str.strip().str.lower()
    reg.calidad_id.append({"hoja": hoja, "revision": "Correo con formato válido (todas las filas)",
                           "registros_ok": int(mail.map(lambda v: bool(RE_CORREO.fullmatch(v))).sum()), "total": n0})

    # --- duplicados ---
    mail_r = mail.loc[corr].reset_index(drop=True)
    k = 0
    grupo = pd.Series("", index=r.index)
    for valor, idx in r.groupby(mail_r).groups.items():
        if len(idx) > 1:
            k += 1
            grupo.loc[list(idx)] = f"DUP-R{k:02d}"
            for i in idx:
                reg.a_revision(hoja, ids[i], "toda la fila", "Posible duplicado (mismo correo).", "Revisar.")
            reg.duplicados.append({"hoja": hoja, "grupo": f"DUP-R{k:02d}", "id_a": ids[list(idx)[0]],
                                   "id_b": ids[list(idx)[1]], "motivo": "Mismo correo.",
                                   "columnas_de_respuesta_identicas": "", "minutos_entre_envios": "", "accion": "NO se eliminó."})
    for i in r.index[r.duplicated(subset=cols[3:], keep=False) & (grupo == "")]:
        reg.a_revision(hoja, ids[i], "toda la fila", "Respuestas idénticas a otra fila.", "Revisar.")
    reg.log(hoja, "Detección de posibles duplicados", "correo / respuestas", "Agrupar por valor idéntico; solo se marca",
            (grupo != "").sum(), "Regla 3.")

    # --- carrera ---
    def homol_carrera(x):
        if en_blanco(x):
            return np.nan
        t = norm(x)
        if re.fullmatch(r"(licenciatura en )?matematicas aplicadas (y|en) computacion", t) or t == "mac":
            return "Matemáticas Aplicadas y Computación"
        return np.nan

    carrera_h = r[c_carr].map(homol_carrera)
    for i in r.index[carrera_h.isna() & ~mascara_blancos(r[c_carr])]:
        reg.a_revision(hoja, ids[i], f"col_{pos[c_carr]}", "Carrera no reconocida por la regla.", "Revisar.")
    for (o, h), n in pd.DataFrame({"o": r[c_carr].astype(str), "h": carrera_h}).groupby(["o", "h"], dropna=False).size().items():
        reg.homologaciones.append({"hoja": hoja, "columna": "carrera", "valor_original": repr(o), "valor_homologado": h, "registros": int(n)})
    reg.log(hoja, "Homologación de categorías", "carrera → carrera_homologada",
            "norm(texto) = [licenciatura en] matematicas aplicadas (y|en) computacion, o 'mac' → 'Matemáticas Aplicadas y Computación'",
            int((carrera_h.astype(str) != r[c_carr].astype(str)).sum()),
            "Variantes de escritura y la abreviatura MAC (carrera del grupo). Se conserva el original.")

    # --- pareja ---
    def homol_pareja(x):
        if en_blanco(x):
            return np.nan
        t = norm(x)
        return {"si": "Sí", "no": "No"}.get(t, np.nan)

    pareja_h = r[c_par].map(homol_pareja)
    for i in r.index[pareja_h.isna() & ~mascara_blancos(r[c_par])]:
        reg.a_revision(hoja, ids[i], f"col_{pos[c_par]}", "Respuesta de pareja no reconocida.", "Revisar.")
    for (o, h), n in pd.DataFrame({"o": r[c_par].astype(str), "h": pareja_h}).groupby(["o", "h"], dropna=False).size().items():
        reg.homologaciones.append({"hoja": hoja, "columna": "¿Tienes pareja actualmente?", "valor_original": repr(o),
                                   "valor_homologado": h, "registros": int(n)})
    reg.log(hoja, "Homologación de categorías", "pareja → pareja_homologada",
            "norm(texto) ∈ {si, no} → Sí / No", int((pareja_h.astype(str) != r[c_par].astype(str)).sum()),
            "Variantes de acento y mayúsculas.")

    # --- semestre y edad (tipos y rangos) ---
    sem = pd.to_numeric(r[c_sem], errors="coerce")
    edad = pd.to_numeric(r[c_edad], errors="coerce")
    r[c_sem] = sem.astype("Int64")
    r[c_edad] = edad.astype("Int64")
    reg.log(hoja, "Conversión de tipo de dato", f"{c_sem[:30]}…, {c_edad[:30]}…", "float → entero anulable (Int64), sin cambiar valores",
            int(sem.notna().sum() + edad.notna().sum()), "Los valores numéricos deben tener tipo numérico adecuado.")
    for i in r.index:
        if not pd.isna(sem[i]) and not (1 <= sem[i] <= 12):
            reg.atipico(hoja, ids[i], "semestre", int(sem[i]), "Fuera del rango 1-12.")

    notas_edad = []
    iqr_edad = mascara_iqr(edad)
    for i in r.index:
        v = edad[i]
        msg = ""
        if pd.isna(v):
            msg = ""
        elif v <= 0:
            msg = "Edad 0: valor no válido como edad; podría significar 'nunca he tenido pareja'. No se asume ni se cambia: verificar."
        elif v < 5 or v > 50:
            msg = "Edad fuera del rango plausible (5-50)."
        elif iqr_edad[i]:
            msg = "Atípico estadístico (1.5×IQR). No hay variable de edad actual para contrastarlo."
        if msg:
            reg.atipico(hoja, ids[i], "edad de la primera relación", int(v), msg)
            if v <= 0:
                reg.a_revision(hoja, ids[i], f"col_{pos[c_edad]}", msg, "Confirmar con el equipo; no usar como edad real sin decidir.")
        notas_edad.append(msg)
    reg.log(hoja, "Marcado de edades sospechosas", "edad_primera_relacion_nota",
            "Edad <=0, fuera de 5-50 o atípica por 1.5×IQR → nota; el valor no se sustituye",
            sum(1 for n in notas_edad if n), "Regla de la hoja: señalar valores sospechosos sin sustituirlos.")

    # --- diferencia de edad máxima: solo formatos simples ---
    def parsea_dif(x):
        if en_blanco(x):
            return (np.nan, "", "No")
        t = luz(x)
        m = re.fullmatch(r"(\d+)\s*(anos?)?(\s+de\s+diferencia)?", t)
        if m:
            return (float(m.group(1)), "Número simple; la dirección (mayor/menor) no está especificada.",
                    "Sí" if m.group(2) is None else "No")
        return (np.nan, "Formato no simple (signos, direcciones o texto): no se convierte; revisar.", "No")

    pd_ = [parsea_dif(x) for x in r[c_dif]]
    dif_val = pd.Series([p[0] for p in pd_], dtype="float64")
    dif_nota = [p[1] for p in pd_]
    dif_ctx = [p[2] for p in pd_]
    for i in r.index:
        if dif_nota[i].startswith("Formato no simple"):
            reg.a_revision(hoja, ids[i], f"col_{pos[c_dif]}", dif_nota[i], "Revisión manual del formato.")
    reg.log(hoja, "Extracción numérica de la diferencia de edad máxima", "dif_edad_max_anios",
            "Solo si el texto es 'N', 'N año(s)' o 'N años de diferencia'; cualquier otro formato queda vacío y a revisión",
            int(dif_val.notna().sum()), "No se interpretan signos ni direcciones (+/-, 'más que yo').")

    revis = revisar_texto_abierto(r, [c for c in c_open if c != c_dif], ids, hoja, reg, pos)
    reg.log(hoja, "Marcado de respuestas poco informativas / posible dato personal", "preguntas abiertas",
            "Longitud <=3, patrones N/A|ns|no sé|nada|ninguno, o regex de correo/teléfono; solo se marca",
            sum(1 for x in revis if x), "Regla 8.")

    # variables constantes
    for c in (c_area, c_sem):
        if r[c].nunique(dropna=True) <= 1:
            reg.a_revision(hoja, "(columna)", f"col_{pos[c]}", "Variable constante: sin variabilidad, no aporta al modelado.",
                           "Mantenida; considerar excluirla del modelado.")

    limpio = pd.DataFrame({"id_participante": ids})
    limpio[c_area] = r[c_area].values
    limpio[c_carr] = r[c_carr].values
    limpio["carrera_homologada"] = carrera_h.values
    limpio[c_sem] = r[c_sem].values
    limpio[c_par] = r[c_par].values
    limpio["pareja_homologada"] = pareja_h.values
    limpio[c_edad] = r[c_edad].values
    limpio["edad_primera_relacion_nota"] = notas_edad
    for c in c_open:
        limpio[c] = r[c].values
        if c == c_dif:
            limpio["dif_edad_max_anios"] = dif_val.values
            limpio["dif_edad_max_nota"] = dif_nota
            limpio["dif_edad_max_sin_unidad"] = dif_ctx
    limpio["revision_respuestas_cortas"] = revis
    limpio["grupo_posible_duplicado"] = [g if g else np.nan for g in grupo]
    reg.log(hoja, "Retiro de identificadores y columna de consentimiento", "marca temporal, correo, consentimiento",
            "Columnas eliminadas del conjunto analítico",
            len(r) * 3, "Regla 4. El consentimiento es constante ('Sí') tras la exclusión.")
    reg.log(hoja, "Creación de ID anónimo", "id_participante", "R + número consecutivo según el orden de filas de ESTA hoja", len(r),
            "Regla 5: independiente de P### y M###.")

    der = ["carrera_homologada", "pareja_homologada", "edad_primera_relacion_nota", "dif_edad_max_anios",
           "dif_edad_max_nota", "dif_edad_max_sin_unidad", "revision_respuestas_cortas", "grupo_posible_duplicado"]
    meta = {"retiradas": [c_ts, c_mail, c_cons], "original": [c_area, c_carr, c_sem, c_par, c_edad] + c_open,
            "derivadas": der, "pos": pos, "n0": n0, "excluidos": int((~si).sum()), "indices_conservados": list(corr)}
    return limpio, meta


# ----------------------------------------------------------------------------
# Validaciones finales, reportes y escritura
# ----------------------------------------------------------------------------


def valores_iguales(a, b):
    """Compara dos series (original vs limpia) tratando faltantes como iguales entre sí."""
    a = a.reset_index(drop=True)
    b = b.reset_index(drop=True)
    na_a, na_b = a.isna(), b.isna()
    if (na_a != na_b).any():
        return False
    ok = ~na_a
    try:
        return bool((pd.to_numeric(a[ok]) == pd.to_numeric(b[ok])).all()) if ok.any() else True
    except (ValueError, TypeError):
        return bool((a[ok].astype(str) == b[ok].astype(str)).all())


def validar(reg, originales, limpios, metas):
    for hoja in (HOJA_P, HOJA_M, HOJA_R):
        raw, lim, meta = originales[hoja], limpios[hoja], metas[hoja]
        esperado = len(raw) - meta["excluidos"]
        reg.validar(f"[{hoja}] conteo de filas coherente",
                    len(lim) == esperado, f"iniciales={len(raw)}, excluidas={meta['excluidos']}, finales={len(lim)}")
        reg.validar(f"[{hoja}] IDs anónimos únicos", lim["id_participante"].is_unique, f"{lim['id_participante'].nunique()} IDs únicos")
        prefijo = {HOJA_P: "P", HOJA_M: "M", HOJA_R: "R"}[hoja]
        reg.validar(f"[{hoja}] prefijo de ID correcto e independiente por hoja",
                    lim["id_participante"].str.startswith(prefijo).all(), f"prefijo {prefijo}")
        # columnas originales preservadas con los mismos valores
        if "indices_conservados" in meta:
            base = raw.loc[meta["indices_conservados"]]
        else:
            base = raw
        malas = [c for c in meta["original"] if c not in lim.columns or not valores_iguales(base[c], lim[c])]
        reg.validar(f"[{hoja}] columnas originales de respuesta conservadas sin cambios", not malas,
                    "todas iguales" if not malas else f"difieren: {malas}")
        # faltantes siguen siendo faltantes
        nf_ori = int(sum(base[c].isna().sum() for c in meta["original"]))
        nf_lim = int(sum(lim[c].isna().sum() for c in meta["original"]))
        reg.validar(f"[{hoja}] faltantes originales no se convirtieron en otro valor", nf_ori == nf_lim,
                    f"faltantes originales={nf_ori}, finales={nf_lim}")
        # sin identificadores
        patrones = re.compile(r"marca temporal|correo|n[uú]mero de cuenta", re.I)
        reg.validar(f"[{hoja}] sin columnas identificadoras", not any(patrones.search(c) for c in lim.columns), "")
        hay_id = False
        for c in lim.columns:
            if lim[c].dtype == object:
                if lim[c].dropna().astype(str).map(lambda v: bool(RE_CUENTA.match(v.strip()) or RE_CORREO.fullmatch(v.strip()))).any():
                    hay_id = True
        reg.validar(f"[{hoja}] ninguna celda es un correo o un número de cuenta completo", not hay_id, "")

    m = limpios[HOJA_M]
    ok = m[(m["tiempo_min_horas"].notna()) & (m["tiempo_max_horas"].notna())]
    reg.validar("[Música] mínimo <= máximo cuando ambos existen", bool((ok["tiempo_min_horas"] <= ok["tiempo_max_horas"]).all()),
                f"{len(ok)} registros comparados")
    sin_max = m[m["tipo_registro_tiempo"] == "minimo_sin_maximo"]
    reg.validar("[Música] 'mínimo sin máximo' deja el máximo vacío", bool(sin_max["tiempo_max_horas"].isna().all()),
                f"{len(sin_max)} registros")
    reg.validar("[Música] tiempo_min/max son numéricos",
                pd.api.types.is_float_dtype(m["tiempo_min_horas"]) and pd.api.types.is_float_dtype(m["tiempo_max_horas"]), "float64")
    amb = m[m["tipo_registro_tiempo"] == "ambiguo_sin_convertir"]
    reg.validar("[Música] respuestas ambiguas no se forzaron (min/max vacíos)",
                bool(amb["tiempo_min_horas"].isna().all() and amb["tiempo_max_horas"].isna().all()), f"{len(amb)} registros")
    a = limpios[HOJA_R]
    reg.validar("[Amor] tipos numéricos correctos (semestre, edad)",
                str(a.iloc[:, [i for i, c in enumerate(a.columns) if c.startswith("¿En qué semestre")][0]].dtype) == "Int64"
                and str(a.iloc[:, [i for i, c in enumerate(a.columns) if c.startswith("¿A qué edad")][0]].dtype) == "Int64", "Int64")
    n_si = int(originales[HOJA_R].iloc[:, 2].map(norm).eq("si").sum())
    reg.validar("[Amor] solo permanecen filas con consentimiento 'sí'", len(a) == n_si,
                f"filas con consentimiento en el original={n_si}, filas finales={len(a)}")


def resumenes(reg, originales, limpios, metas):
    for hoja in (HOJA_P, HOJA_M, HOJA_R):
        raw, lim, meta = originales[hoja], limpios[hoja], metas[hoja]
        reg.resumen.append({
            "hoja": hoja, "registros_iniciales": len(raw), "registros_excluidos": meta["excluidos"],
            "registros_finales": len(lim), "columnas_originales_en_hoja_fuente": raw.shape[1],
            "identificadores_retirados": len(meta["retiradas"]),
            "columnas_en_hoja_limpia": lim.shape[1],
            "columnas_originales_conservadas": len(meta["original"]),
            "columnas_derivadas": len(meta["derivadas"]),
        })
        for c in raw.columns:
            reg.columnas.append({"hoja": hoja, "posicion_original": f"col_{meta['pos'][c]}", "columna_original": c,
                                 "tratamiento": "Identificador/consentimiento: retirado del conjunto analítico" if c in meta["retiradas"]
                                 else "Respuesta original conservada"})
        for c in meta["derivadas"] + ["id_participante"]:
            reg.columnas.append({"hoja": hoja, "posicion_original": "(nueva)", "columna_original": c,
                                 "tratamiento": "Columna analítica creada por la limpieza"})
        for c in lim.columns:
            vacios = int(lim[c].isna().sum() + (lim[c].map(lambda v: isinstance(v, str) and v.strip() == "").sum()))
            reg.faltantes.append({"hoja": hoja, "columna": c, "faltantes": vacios,
                                  "porcentaje": round(100 * vacios / max(len(lim), 1), 1),
                                  "nota": "Los vacíos en columnas derivadas pueden ser intencionales (p. ej. sin marca = vacío)."
                                  if c in meta["derivadas"] else ""})


def guardar_xlsx(ruta, hojas, ancho_max=60):
    with pd.ExcelWriter(ruta, engine="openpyxl") as w:
        for nombre, df in hojas.items():
            df.to_excel(w, sheet_name=nombre[:31], index=False)
            ws = w.sheets[nombre[:31]]
            fuente_h = Font(name="Arial", bold=True, color="FFFFFF", size=10)
            relleno = PatternFill("solid", start_color="1F4E78")
            for cell in ws[1]:
                cell.font, cell.fill = fuente_h, relleno
                cell.alignment = Alignment(wrap_text=True, vertical="center")
            for fila in ws.iter_rows(min_row=2):
                for cell in fila:
                    cell.font = Font(name="Arial", size=10)
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        cell.data_type = "s"  # evitar que una respuesta se interprete como fórmula
            for j, col in enumerate(df.columns, start=1):
                largo = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)]) if len(df) else len(str(col))
                ws.column_dimensions[get_column_letter(j)].width = min(max(12, largo + 2), ancho_max)
            ws.freeze_panes = "B2" if "id_participante" in df.columns else "A2"
            ws.row_dimensions[1].height = 45


def main():
    ap = argparse.ArgumentParser(description="Limpieza y auditoría de las respuestas de los formularios.")
    ap.add_argument("--entrada", default="RespuestasFormsOriginales.xlsx", help="Libro original (solo lectura).")
    ap.add_argument("--salida", default=".", help="Carpeta donde se escriben los resultados.")
    args = ap.parse_args()

    entrada = Path(args.entrada)
    salida = Path(args.salida)
    salida.mkdir(parents=True, exist_ok=True)
    if not entrada.exists():
        sys.exit(f"No se encontró el archivo original: {entrada}")

    originales = {h: pd.read_excel(entrada, sheet_name=h) for h in (HOJA_P, HOJA_M, HOJA_R)}
    reg = Registro()
    limpios, metas = {}, {}
    for hoja, f in ((HOJA_P, limpiar_personalidad), (HOJA_M, limpiar_musica), (HOJA_R, limpiar_amor)):
        limpios[hoja], metas[hoja] = f(originales[hoja], reg)

    validar(reg, originales, limpios, metas)
    resumenes(reg, originales, limpios, metas)

    guardar_xlsx(salida / "Datos_Limpios.xlsx", limpios)
    guardar_xlsx(salida / "Bitacora_Limpieza.xlsx", {"Bitacora": pd.DataFrame(reg.bitacora)}, ancho_max=70)

    def df_o(lista, cols):
        return pd.DataFrame(lista) if lista else pd.DataFrame(columns=cols)

    guardar_xlsx(salida / "Reporte_Auditoria.xlsx", {
        "Resumen": pd.DataFrame(reg.resumen),
        "Columnas": pd.DataFrame(reg.columnas),
        "Faltantes": pd.DataFrame(reg.faltantes),
        "Duplicados": df_o(reg.duplicados, ["hoja", "grupo", "id_a", "id_b", "motivo", "accion"]),
        "Atipicos_Sospechosos": df_o(reg.atipicos, ["hoja", "id_participante", "columna", "valor", "motivo"]),
        "Homologaciones": pd.DataFrame(reg.homologaciones),
        "Exclusiones": pd.DataFrame(reg.exclusiones),
        "Casos_Revision": df_o(reg.revision, ["hoja", "id_participante", "columna", "motivo", "accion_sugerida"]),
        "Calidad_Identificadores": pd.DataFrame(reg.calidad_id),
        "Validaciones": pd.DataFrame(reg.validaciones),
    })

    fallas = [v for v in reg.validaciones if v["resultado"] != "OK"]
    print("Registros (iniciales → finales):")
    for r in reg.resumen:
        print(f"  {r['hoja']}: {r['registros_iniciales']} → {r['registros_finales']} (excluidos: {r['registros_excluidos']})")
    print(f"Validaciones: {len(reg.validaciones) - len(fallas)} OK, {len(fallas)} con falla")
    for v in fallas:
        print("  FALLA:", v["validacion"], v["detalle"])
    print(f"Casos para revisión humana: {len(reg.revision)} | Atípicos/sospechosos: {len(reg.atipicos)}")
    print(f"Archivos escritos en: {salida.resolve()}")
    if fallas:
        sys.exit(1)


if __name__ == "__main__":
    main()
