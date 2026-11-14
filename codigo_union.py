#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
codigo_union.py
===============
Une los tres formularios (personalidad, música y relaciones amorosas) en un
solo conjunto de datos, UNA FILA POR PERSONA, con un ID anónimo compartido (U###).

Contexto y autorización
-----------------------
La limpieza (codigo_limpieza.py) trata cada hoja por separado y con IDs
independientes. La maestra indicó que SÍ se puede cruzar la información de los
formularios, así que este script hace ese cruce de forma controlada:

  * La llave de unión es el número de cuenta (en el formulario de relaciones
    amorosas, el prefijo numérico del correo). La llave se usa SOLO en memoria.
  * El archivo de salida NO contiene correos, números de cuenta ni marcas de
    tiempo, y este script NO guarda ninguna tabla de equivalencias.
  * Se parte de la hoja de personalidad (tema principal) y se agregan las
    columnas de música y de relaciones amorosas (unión por la izquierda).
  * Se aplican las mismas exclusiones de la limpieza (duplicado de personalidad
    y falta de consentimiento en relaciones amorosas).

IMPORTANTE: este script necesita el libro ORIGINAL (con correos y cuentas), por
lo que solo puede ejecutarse en la computadora del equipo. Al repositorio de
GitHub se sube únicamente el resultado (Datos_Unidos.xlsx).

Uso:
    python codigo_union.py
    python codigo_union.py --entrada RespuestasFormsOriginales.xlsx --salida .
"""
import argparse
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import codigo_limpieza as cl  # noqa: E402  (reutiliza las mismas reglas de limpieza)


def llave_desde_correo(serie):
    """Prefijo numérico del correo (antes de la @) como texto."""
    return (serie.astype(str).str.strip().str.lower().str.extract(r"^(\d+)@")[0])


def construir_llaves(originales, limpios, metas):
    """Devuelve, para cada hoja limpia, la llave de persona alineada fila a fila."""
    llaves = {}

    raw = originales[cl.HOJA_P]
    cta = cl.cuenta_texto(raw[raw.columns[2]])
    pos = limpios[cl.HOJA_P]["id_participante"].str[1:].astype(int) - 1
    llaves[cl.HOJA_P] = cta.iloc[pos.values].reset_index(drop=True)

    raw = originales[cl.HOJA_M]
    cta = cl.cuenta_texto(raw[raw.columns[1]])
    pos = limpios[cl.HOJA_M]["id_participante"].str[1:].astype(int) - 1
    llaves[cl.HOJA_M] = cta.iloc[pos.values].reset_index(drop=True)

    raw = originales[cl.HOJA_R]
    lla = llave_desde_correo(raw[raw.columns[1]])
    conservados = metas[cl.HOJA_R]["indices_conservados"]
    pos = limpios[cl.HOJA_R]["id_participante"].str[1:].astype(int) - 1
    llaves[cl.HOJA_R] = lla.iloc[[conservados[i] for i in pos.values]].reset_index(drop=True)
    return llaves


def renombrar(limpio, meta, hoja):
    """
    Renombra a nombres cortos y devuelve (df_renombrado, diccionario).
    Las columnas originales de texto libre se conservan; para las categóricas se
    deja solo la versión homologada (los originales siguen en Datos_Limpios.xlsx).
    """
    dic = []
    out = pd.DataFrame(index=limpio.index)

    def agregar(nombre, serie, origen, descripcion):
        out[nombre] = serie.values
        dic.append({"variable": nombre, "hoja_origen": hoja, "origen": origen, "descripcion": descripcion})

    if hoja == cl.HOJA_P:
        c_gen = meta["original"][0]
        agregar("pers_genero", limpio["genero_homologado"], c_gen, "Género homologado (Masculino/Femenino)")
        for i, c in enumerate(meta["original"][1:], start=2):
            agregar(f"pers_q{i:02d}", limpio[c], c, "Respuesta abierta original (sin modificar)")

    elif hoja == cl.HOJA_M:
        c_tiempo = meta["original"][1]
        for i, c in enumerate(meta["original"], start=1):
            agregar(f"mus_q{i:02d}", limpio[c], c, "Respuesta abierta original (sin modificar)")
        extra = {
            "tiempo_min_horas": ("mus_tiempo_min_horas", "Horas/día mínimas (regla determinista; puntual: min=max)"),
            "tiempo_max_horas": ("mus_tiempo_max_horas", "Horas/día máximas; vacío si solo hay mínimo o es ambiguo"),
            "tipo_registro_tiempo": ("mus_tiempo_tipo", "Tipo de registro: puntual, intervalo, aproximado, mínimo sin máximo, ambiguo"),
            "tiempo_interpretacion_contextual": ("mus_tiempo_contextual", "Sí si no traía unidad y se asumió 'horas' por contexto"),
            "tiempo_requiere_revision": ("mus_tiempo_revisar", "Sí si el valor es 0, >12 h o ambiguo"),
            "nota_limpieza_tiempo": ("mus_tiempo_nota", "Nota de la limpieza del tiempo de escucha"),
        }
        for orig, (nuevo, desc) in extra.items():
            agregar(nuevo, limpio[orig], f"derivada de: {c_tiempo}", desc)

    else:  # relaciones amorosas
        o = meta["original"]
        c_area, c_carr, c_sem, c_par, c_edad = o[:5]
        abiertas = o[5:]
        agregar("amor_area", limpio[c_area], c_area, "Área de conocimiento")
        agregar("amor_carrera", limpio["carrera_homologada"], c_carr, "Carrera homologada")
        agregar("amor_semestre", limpio[c_sem], c_sem, "Semestre")
        agregar("amor_pareja", limpio["pareja_homologada"], c_par, "¿Tiene pareja? homologada (Sí/No)")
        agregar("amor_edad_primera_relacion", limpio[c_edad], c_edad,
                "Edad de la primera relación (0 = valor sospechoso, NO se sustituyó)")
        agregar("amor_edad_nota", limpio["edad_primera_relacion_nota"], f"derivada de: {c_edad}", "Nota de revisión de la edad")
        for i, c in enumerate(abiertas, start=1):
            agregar(f"amor_q{i:02d}", limpio[c], c, "Respuesta abierta original (sin modificar)")
        c_dif = abiertas[7]
        agregar("amor_dif_edad_max_anios", limpio["dif_edad_max_anios"], f"derivada de: {c_dif}",
                "Años de diferencia, solo si el formato era simple")
        agregar("amor_dif_edad_max_nota", limpio["dif_edad_max_nota"], f"derivada de: {c_dif}", "Nota de la conversión")
        agregar("amor_dif_edad_max_sin_unidad", limpio["dif_edad_max_sin_unidad"], f"derivada de: {c_dif}",
                "Sí si no traía unidad y se asumieron años")
    return out, dic


def main():
    ap = argparse.ArgumentParser(description="Une los tres formularios por persona con un ID anónimo compartido.")
    ap.add_argument("--entrada", default="RespuestasFormsOriginales.xlsx", help="Libro original (solo lectura).")
    ap.add_argument("--salida", default=".", help="Carpeta de salida.")
    args = ap.parse_args()

    entrada, salida = Path(args.entrada), Path(args.salida)
    salida.mkdir(parents=True, exist_ok=True)
    if not entrada.exists():
        sys.exit(f"No se encontró el archivo original: {entrada}")

    hojas = (cl.HOJA_P, cl.HOJA_M, cl.HOJA_R)
    originales = {h: pd.read_excel(entrada, sheet_name=h) for h in hojas}
    reg = cl.Registro()
    limpios, metas = {}, {}
    for hoja, f in ((cl.HOJA_P, cl.limpiar_personalidad), (cl.HOJA_M, cl.limpiar_musica), (cl.HOJA_R, cl.limpiar_amor)):
        limpios[hoja], metas[hoja] = f(originales[hoja], reg)

    llaves = construir_llaves(originales, limpios, metas)
    val = []

    def check(nombre, ok, detalle=""):
        val.append({"validacion": nombre, "resultado": "OK" if ok else "FALLA", "detalle": detalle})

    # llaves únicas y válidas dentro de cada hoja (si no, la unión sería ambigua)
    for h in hojas:
        k = llaves[h]
        check(f"[{h}] llaves de persona válidas (9 dígitos)", bool(k.map(lambda v: isinstance(v, str) and bool(cl.RE_CUENTA.match(v))).all()),
              f"{len(k)} filas")
        check(f"[{h}] una sola fila por persona (llaves únicas)", bool(k.is_unique), f"{k.nunique()} llaves únicas de {len(k)}")

    renombrados, diccionario = {}, []
    for h in hojas:
        renombrados[h], d = renombrar(limpios[h], metas[h], h)
        diccionario += d
        renombrados[h]["_llave"] = llaves[h].values

    base = renombrados[cl.HOJA_P]
    claves_p = set(base["_llave"])
    for h in (cl.HOJA_M, cl.HOJA_R):
        fuera = set(renombrados[h]["_llave"]) - claves_p
        check(f"[{h}] todas las personas existen en personalidad (nadie se pierde al unir)", len(fuera) == 0,
              f"personas sin pareja en personalidad: {len(fuera)}")

    union = base.copy()
    for h, marca in ((cl.HOJA_M, "tiene_musica"), (cl.HOJA_R, "tiene_amor")):
        otro = renombrados[h].copy()
        otro[marca] = "Sí"
        union = union.merge(otro, on="_llave", how="left", validate="one_to_one")
        union[marca] = union[marca].where(union[marca].notna(), "No")

    n_p, n_m, n_r = (len(renombrados[h]) for h in (cl.HOJA_P, cl.HOJA_M, cl.HOJA_R))
    check("Filas del resultado = filas de personalidad", len(union) == n_p, f"{len(union)} = {n_p}")
    check("Personas con música = filas de música", int((union["tiene_musica"] == "Sí").sum()) == n_m,
          f"{int((union['tiene_musica'] == 'Sí').sum())} de {n_m}")
    check("Personas con relaciones amorosas = filas de esa hoja", int((union["tiene_amor"] == "Sí").sum()) == n_r,
          f"{int((union['tiene_amor'] == 'Sí').sum())} de {n_r}")

    union.insert(0, "id_persona", [f"U{i:03d}" for i in range(1, len(union) + 1)])
    union = union.drop(columns="_llave")
    cols_ini = ["id_persona", "tiene_musica", "tiene_amor"]
    union = union[cols_ini + [c for c in union.columns if c not in cols_ini]]
    check("ID compartido único", bool(union["id_persona"].is_unique), f"{union['id_persona'].nunique()} IDs")

    # valores de las respuestas originales no se alteraron al unir
    for h, prefijo in ((cl.HOJA_P, "pers_q"), (cl.HOJA_M, "mus_q"), (cl.HOJA_R, "amor_q")):
        cols = [c for c in renombrados[h].columns if c.startswith(prefijo)]
        ok = True
        for c in cols:
            sub = union[c]
            esperado = renombrados[h].set_index("_llave")[c]
            # comparar solo filas que tienen dato en la hoja de origen
            llave_union = base["_llave"].values
            merged = pd.Series(esperado.reindex(llave_union).values, index=range(len(llave_union)))
            ok = ok and cl.valores_iguales(sub, merged)
        check(f"[{h}] respuestas abiertas idénticas tras la unión", ok, f"{len(cols)} columnas comparadas")

    # privacidad
    pat = re.compile(r"marca temporal|correo|n[uú]mero de cuenta|cuenta", re.I)
    check("Sin columnas identificadoras", not any(pat.search(c) for c in union.columns), "")
    hay = False
    for c in union.columns:
        if union[c].dtype == object:
            if union[c].dropna().astype(str).map(lambda v: bool(cl.RE_CUENTA.match(v.strip()) or cl.RE_CORREO.fullmatch(v.strip()))).any():
                hay = True
    check("Ninguna celda es un correo o un número de cuenta", not hay, "")
    # nulos en amor solo donde no hay formulario
    cols_amor = [c for c in union.columns if c.startswith("amor_") and not c.endswith(("_nota", "_sin_unidad"))]
    sin = union[union["tiene_amor"] == "No"]
    check("Columnas de relaciones amorosas vacías solo para quien no respondió ese formulario",
          bool(sin[cols_amor].isna().all().all()), f"{len(sin)} personas sin ese formulario")

    # resumen de cobertura
    completas = int(((union["tiene_musica"] == "Sí") & (union["tiene_amor"] == "Sí")).sum())
    cobertura = pd.DataFrame([
        {"concepto": "Personas (filas)", "valor": len(union)},
        {"concepto": "Con formulario de música", "valor": int((union["tiene_musica"] == "Sí").sum())},
        {"concepto": "Con formulario de relaciones amorosas", "valor": int((union["tiene_amor"] == "Sí").sum())},
        {"concepto": "Con los tres formularios completos", "valor": completas},
    ])

    cl.guardar_xlsx(salida / "Datos_Unidos.xlsx", {
        "DatosUnidos": union,
        "Diccionario": pd.DataFrame(diccionario),
        "Cobertura": cobertura,
        "ValidacionUnion": pd.DataFrame(val),
    })

    fallas = [v for v in val if v["resultado"] != "OK"]
    print(f"Personas en el dataset unido: {len(union)} | con música: {int((union['tiene_musica'] == 'Sí').sum())}"
          f" | con relaciones amorosas: {int((union['tiene_amor'] == 'Sí').sum())} | con los tres: {completas}")
    print(f"Columnas: {union.shape[1]} | Validaciones: {len(val) - len(fallas)} OK, {len(fallas)} con falla")
    for v in fallas:
        print("  FALLA:", v["validacion"], v["detalle"])
    print(f"Archivo escrito en: {salida.resolve() / 'Datos_Unidos.xlsx'}")
    if fallas:
        sys.exit(1)


if __name__ == "__main__":
    main()
