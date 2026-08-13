#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generar_reporte.py
==================
Automatiza la integración y el cálculo de usabilidad de licencias de Claude.

Flujo:
  1. Lee las pestañas "Asignación" (inventario administrativo) y
     "Usabilidad" (export de la consola) del archivo de entrada.
  2. Limpia y normaliza la llave de cruce (correo electrónico).
  3. Consolida asignaciones duplicadas (upgrades) en una sola fila.
  4. Hace un LEFT JOIN Asignación -> Usabilidad.
  5. Calcula la fecha de corte (máx. de "Last Active") y las columnas nuevas.
  6. Exporta un archivo Excel con varios bloques/hojas e indicadores.

Uso:
    python3 generar_reporte.py [archivo_entrada.xlsx]

Requisitos: pandas, numpy, xlsxwriter
"""

import sys
import unicodedata
import numpy as np
import pandas as pd

# --------------------------------------------------------------------------- #
# Configuración
# --------------------------------------------------------------------------- #
ARCHIVO_ENTRADA_DEFAULT = "Usuarios de licencias Claude con CeCo.xlsx"
ARCHIVO_SALIDA = "Analisis_Usabilidad_Final.xlsx"

HOJA_ASIGNACION = "Asignación"
HOJA_USABILIDAD = "Usabilidad"

# Ranking de licencias: mayor número = licencia "más alta / vigente".
RANK_LICENCIAS = {"premium": 3, "pro": 2, "standard": 1}

# Columnas numéricas que se asegura tratar como número.
COLS_COSTOS = [
    "Costo anual (USD)",
    "Facturado (USD)",
    "Facturado (MXN)",
    "Presupuesto Ops",
    "Cantidad lic",
]
# Reglas de consolidación cuando un correo tiene varios registros:
#   - Campos vigentes: se toman del ÚLTIMO registro.
#   - Montos: se SUMAN a lo largo de todos los registros del usuario.
CAMPOS_ULTIMO = ["Centro de costos", "CeCo Team", "User Team",
                 "Licencia", "Uso", "Fecha"]
COLS_SUMA = ["Costo anual (USD)", "Facturado (USD)",
             "Facturado (MXN)", "Presupuesto Ops"]

# Métricas de usabilidad numéricas (para rellenar con 0 en el left join).
METRICAS_NUM = [
    "Days Active", "Chats", "Messages", "Projects Created", "Projects Used",
    "Pull Requests", "Code sessions", "File Edits", "Cowork Sessions",
    "Cowork Messages", "Artifacts Created", "Estimated Spend (USD)",
]
# Métricas de usabilidad de texto (para rellenar con "Sin datos").
METRICAS_TXT = ["Name", "Role", "Seat Tier"]


# --------------------------------------------------------------------------- #
# Utilidades
# --------------------------------------------------------------------------- #
def normaliza_correo(serie: pd.Series) -> pd.Series:
    """Normaliza correos: minúsculas, sin espacios y sin acentos accidentales."""
    return (
        serie.astype("string")
        .str.strip()
        .str.lower()
        .str.replace(r"\s+", "", regex=True)
    )


def rank_licencia(texto) -> int:
    """Devuelve un peso numérico para ordenar licencias (mayor = más alta)."""
    if not isinstance(texto, str):
        return 0
    t = texto.lower()
    for clave, peso in RANK_LICENCIAS.items():
        if clave in t:
            return peso
    return 0


# --------------------------------------------------------------------------- #
# 1. Lectura
# --------------------------------------------------------------------------- #
def leer_hojas(ruta: str):
    print(f"[1] Leyendo archivo de entrada: {ruta}")
    asig = pd.read_excel(ruta, sheet_name=HOJA_ASIGNACION)
    usab = pd.read_excel(ruta, sheet_name=HOJA_USABILIDAD)
    # Recorta espacios en los nombres de columna (algunas cabeceras del
    # origen traen espacios iniciales/finales).
    asig.columns = [str(c).strip() for c in asig.columns]
    usab.columns = [str(c).strip() for c in usab.columns]
    print(f"    - Asignación: {asig.shape[0]} filas x {asig.shape[1]} columnas")
    print(f"    - Usabilidad: {usab.shape[0]} filas x {usab.shape[1]} columnas")
    return asig, usab


def resolver_fecha_asignacion(df: pd.DataFrame) -> str:
    """Devuelve el nombre de la columna con la fecha real de asignación.

    Prefiere 'Fecha Asig' (fecha de asignación real de la licencia) si existe;
    si no, cae a 'Fecha'.
    """
    for cand in ("Fecha Asig", "Fecha Asignación", "Fecha asignación", "Fecha Asignacion"):
        if cand in df.columns:
            return cand
    return "Fecha"


# --------------------------------------------------------------------------- #
# 2 y 3. Normalización + consolidación de asignaciones duplicadas
# --------------------------------------------------------------------------- #
def consolidar_asignacion(asig: pd.DataFrame) -> pd.DataFrame:
    print("\n[2] Normalizando llave de cruce y consolidando duplicados...")
    df = asig.copy()
    df["email_key"] = normaliza_correo(df["Correo electrónico"])

    # Fecha de asignación real: usa 'Fecha Asig' si existe; si no, 'Fecha'.
    fuente_fecha = resolver_fecha_asignacion(df)
    df["Fecha"] = pd.to_datetime(df[fuente_fecha], errors="coerce")
    print(f"    - Fecha de asignación tomada de la columna: '{fuente_fecha}'"
          + ("" if fuente_fecha == "Fecha Asig"
             else "  (no se encontró 'Fecha Asig'; se usa 'Fecha')"))

    # Asegura tipos numéricos en costos y fecha en la fecha de asignación.
    for c in COLS_COSTOS:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")

    filas = []
    n_consolidados = 0
    for email, g in df.groupby("email_key", sort=False):
        if len(g) == 1:
            fila = g.iloc[0].to_dict()
            filas.append(fila)
            continue

        # --- Consolidación de un correo con múltiples registros ---
        # Regla:
        #   * Campos vigentes (Centro de costos, CeCo Team, User Team,
        #     Licencia, Uso, Fecha) -> del ÚLTIMO registro.
        #   * Montos (Costo anual USD, Facturado USD, Facturado MXN,
        #     Presupuesto Ops) -> SUMA de todos los registros del usuario.
        #   * Cantidad lic -> SUMA (para reflejar la licencia vigente).
        n_consolidados += 1
        base = g.iloc[-1].to_dict()  # último registro = estado vigente

        for c in COLS_SUMA:
            if c in df.columns:
                base[c] = g[c].sum()
        base["Cantidad lic"] = g["Cantidad lic"].sum()

        # Nota de consolidación en Consideraciones.
        lics = " -> ".join(g["Licencia"].astype(str).tolist())
        nota = (f"Consolidado de {len(g)} registros. Campos vigentes del último "
                f"registro (Licencia: {base['Licencia']}; "
                f"Fecha: {pd.Timestamp(base['Fecha']):%Y-%m-%d}). "
                f"Montos sumados. Historial de licencias: {lics}.")
        prev = base.get("Consideraciones")
        base["Consideraciones"] = (
            f"{prev} | {nota}" if isinstance(prev, str) and prev.strip() else nota
        )
        filas.append(base)

    consol = pd.DataFrame(filas).reset_index(drop=True)
    print(f"    - Correos consolidados (upgrades): {n_consolidados}")
    print(f"    - Asignación consolidada (todos): {consol.shape[0]} filas "
          f"(antes {asig.shape[0]})")
    # Se devuelve el consolidado COMPLETO. El filtro de 'licencias vigentes'
    # (Cantidad lic > 0) se aplica en main() sólo para las vistas de
    # usabilidad; los indicadores de costo usan el total del proyecto.
    n_sin_lic = int((pd.to_numeric(consol["Cantidad lic"], errors="coerce")
                     .fillna(0) <= 0).sum())
    print(f"    - Usuarios sin licencia vigente (Cantidad lic = 0): {n_sin_lic}")
    return consol


# --------------------------------------------------------------------------- #
# 4. Cruce principal (left join)
# --------------------------------------------------------------------------- #
def cruzar(consol: pd.DataFrame, usab: pd.DataFrame):
    print("\n[3] Cruce principal (LEFT JOIN Asignación -> Usabilidad)...")
    u = usab.copy()
    u["email_key"] = normaliza_correo(u["Email"])
    u["Last Active"] = pd.to_datetime(u["Last Active"], errors="coerce")

    # Fecha de corte: máxima "Last Active" de la usabilidad.
    fecha_corte = u["Last Active"].max()
    print(f"    - Fecha de corte (máx. Last Active): {fecha_corte:%Y-%m-%d}")

    # Evita colisión de nombres: la usabilidad trae su propio "Name".
    df = consol.merge(u, on="email_key", how="left", suffixes=("", "_usab"))

    con_uso = df["Days Active"].notna().sum()
    print(f"    - Usuarios de asignación con datos de uso: {con_uso}/{len(df)}")
    print(f"    - Usuarios sin datos de uso: {len(df) - con_uso}")

    # Rellena métricas faltantes.
    for c in METRICAS_NUM:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    for c in METRICAS_TXT:
        if c in df.columns:
            df[c] = df[c].astype("object").where(df[c].notna(), "Sin datos")

    return df, fecha_corte


# --------------------------------------------------------------------------- #
# 5. Cálculos (columnas nuevas)
# --------------------------------------------------------------------------- #
def clasifica_nivel(pct) -> str:
    """Nivel de uso en función del % de uso (pct en porcentaje, no ratio)."""
    if pct >= 101:
        return "1.Muy Alto"
    if pct >= 70:
        return "2.Alto"
    if pct >= 45:
        return "3.Medio"
    if pct >= 10:
        return "4.Bajo"
    return "5.Nulo"


def accion_sugerida(nivel: str, dias_sin_uso: float) -> str:
    if nivel == "Reciente":
        return "Seguimiento (reciente)"
    if nivel == "5.Nulo":
        return "Reasignar / cancelar"
    if nivel == "4.Bajo":
        return "Revisar"
    if pd.notna(dias_sin_uso) and dias_sin_uso > 21:
        return "Revisar (inactividad)"
    return "Mantener"


# Días mínimos con la licencia asignada para poder evaluar el uso.
# Por debajo de este umbral, un usuario "Nulo" se marca como "Reciente".
DIAS_MIN_EVALUACION = 15


def calcular(df: pd.DataFrame, fecha_corte: pd.Timestamp) -> pd.DataFrame:
    print("\n[4] Calculando columnas nuevas...")
    d = df.copy()
    corte = pd.Timestamp(fecha_corte)

    # --- Bloque de días ---
    # Días calendario (mínimo 0).
    dias_cal = (corte - d["Fecha"]).dt.days
    d["Días calendario"] = dias_cal.clip(lower=0).fillna(0).astype(int)

    # Días hábiles (lunes a viernes) entre la fecha de asignación y el corte.
    def busdays(fecha_ini):
        if pd.isna(fecha_ini):
            return 0
        ini = np.datetime64(pd.Timestamp(fecha_ini).date(), "D")
        fin = np.datetime64(corte.date(), "D")
        if fin <= ini:
            return 0
        return int(np.busday_count(ini, fin))
    d["Días hábiles"] = d["Fecha"].apply(busdays)

    # --- Campos calculados ---
    # % de uso = Días activos / Días hábiles (puede superar 100%).
    dias_habiles = d["Días hábiles"].replace(0, np.nan)
    pct = (d["Days Active"] / dias_habiles) * 100
    d["% de uso"] = pct.fillna(0).round(1)

    # Nivel de uso.
    d["Nivel de uso"] = d["% de uso"].apply(clasifica_nivel)
    # Los usuarios "Nulo" con menos de DIAS_MIN_EVALUACION días de licencia
    # asignada aún no son evaluables -> se marcan como "Reciente".
    mask_reciente = (d["Nivel de uso"] == "5.Nulo") & \
                    (d["Días calendario"] < DIAS_MIN_EVALUACION)
    d.loc[mask_reciente, "Nivel de uso"] = "Reciente"
    print(f"    - Usuarios marcados como 'Reciente' "
          f"(< {DIAS_MIN_EVALUACION} días de licencia): {int(mask_reciente.sum())}")

    # Días sin uso: días HÁBILES (lunes a viernes) entre la última conexión
    # y la fecha de corte. Como la última conexión nunca es anterior a la
    # asignación, en el peor de los casos equivale a los "Días hábiles".
    # Si no hay registro de conexión, se toma directamente "Días hábiles".
    d["Días sin uso"] = [
        busdays(la) if pd.notna(la) else dh
        for la, dh in zip(d["Last Active"], d["Días hábiles"])
    ]
    d["Días sin uso"] = pd.Series(d["Días sin uso"], index=d.index).astype("Int64")

    # Días desde la última conexión (naturales): corte - Last Active.
    # Si no hay conexión registrada, se usa "Días calendario".
    dias_desde = (corte - d["Last Active"]).dt.days
    d["Días desde últ. conexión"] = dias_desde.where(
        d["Last Active"].notna(), d["Días calendario"]
    ).clip(lower=0).astype("Int64")

    # Amplitud de adopción (0-5): categorías con actividad > 0.
    cat1 = (d["Chats"] + d["Messages"]) > 0
    cat2 = (d["Projects Created"] + d["Projects Used"]) > 0
    cat3 = (d["Code sessions"] + d["File Edits"] + d["Pull Requests"]) > 0
    cat4 = (d["Cowork Sessions"] + d["Cowork Messages"]) > 0
    cat5 = (d["Artifacts Created"]) > 0
    d["Amplitud de adopción (0-5)"] = (
        cat1.astype(int) + cat2.astype(int) + cat3.astype(int)
        + cat4.astype(int) + cat5.astype(int)
    )

    # Acción sugerida.
    d["Acción sugerida"] = [
        accion_sugerida(nv, ds)
        for nv, ds in zip(d["Nivel de uso"], d["Días sin uso"])
    ]

    return d


# --------------------------------------------------------------------------- #
# Ordenamiento en 3 bloques lógicos
# --------------------------------------------------------------------------- #
def ordenar_columnas(d: pd.DataFrame) -> pd.DataFrame:
    # A) Campos base (identidad + inventario administrativo)
    bloque_a = [
        "Nombre", "Correo electrónico", "Centro de costos", "CeCo Team",
        "User Team", "Licencia", "Uso", "Cantidad lic",
        "Costo anual (USD)", "Facturado (USD)", "Facturado (MXN)",
        "Presupuesto Ops", "Fecha", "Seat Tier", "Role", "Last Active",
        "Days Active", "Chats", "Messages", "Projects Created",
        "Projects Used", "Pull Requests", "Code sessions", "File Edits",
        "Cowork Sessions", "Cowork Messages", "Artifacts Created",
        "Estimated Spend (USD)",
    ]
    # B) Bloque de días
    bloque_b = ["Días calendario", "Días hábiles", "Días sin uso",
                "Días desde últ. conexión"]
    # C) Campos calculados
    bloque_c = [
        "% de uso", "Nivel de uso", "Amplitud de adopción (0-5)",
        "Acción sugerida", "Consideraciones",
    ]
    orden = [c for c in bloque_a + bloque_b + bloque_c if c in d.columns]
    # Columnas auxiliares/redundantes que NO se exportan (llave y duplicados
    # de nombre/correo que aporta la hoja de Usabilidad tras el merge).
    aux = {"email_key", "Name", "Email"}
    resto = [c for c in d.columns if c not in orden and c not in aux]
    return d[orden + resto]


# --------------------------------------------------------------------------- #
# Indicadores / secciones adicionales
# --------------------------------------------------------------------------- #
def volumetria_licencias(consol: pd.DataFrame) -> pd.DataFrame:
    """Crecimiento acumulado de licencias asignadas en el tiempo."""
    g = (
        consol.dropna(subset=["Fecha"])
        .groupby(consol["Fecha"].dt.date)
        .agg(**{
            "Licencias en la fecha": ("Cantidad lic", "sum"),
            "Usuarios en la fecha": ("email_key", "nunique"),
        })
        .reset_index()
        .rename(columns={"Fecha": "Fecha de asignación"})
    )
    g["Fecha de asignación"] = pd.to_datetime(g["Fecha de asignación"])
    g = g.sort_values("Fecha de asignación")
    g["Licencias acumuladas"] = g["Licencias en la fecha"].cumsum()
    g["Usuarios acumulados"] = g["Usuarios en la fecha"].cumsum()
    return g


# Orden canónico de niveles y clasificación de aprovechamiento.
NIVEL_ORDEN = ["1.Muy Alto", "2.Alto", "3.Medio", "4.Bajo", "5.Nulo", "Reciente"]
CLASIF_NIVEL = {
    "1.Muy Alto": "Productivo", "2.Alto": "Productivo", "3.Medio": "Productivo",
    "4.Bajo": "Ocioso", "5.Nulo": "Ocioso", "Reciente": "En adopción",
}


def usabilidad_usuarios(final: pd.DataFrame) -> pd.DataFrame:
    """Usabilidad de TODOS los usuarios (todos los niveles).

    Tabla plana con la última conexión y los días de uso frente a los días
    laborales (hábiles). Se ordena por nivel para permitir filtrar por uno o
    varios niveles (autofiltro en Excel / filtros en el dashboard).
    """
    cols = ["User Team", "Nombre", "Correo electrónico", "Nivel de uso",
            "% de uso", "Última conexión", "Días desde últ. conexión",
            "Días activos", "Días hábiles", "Días sin uso", "Acción sugerida"]
    out = final.copy()
    out["Última conexión"] = out["Last Active"]
    out["Días activos"] = out["Days Active"]
    out["_ord"] = out["Nivel de uso"].map({n: i for i, n in enumerate(NIVEL_ORDEN)})
    out = out.sort_values(
        ["_ord", "User Team", "% de uso"], ascending=[True, True, True]
    )
    return out[cols].reset_index(drop=True)


def costo_por_nivel(final: pd.DataFrame) -> pd.DataFrame:
    """Costo por nivel de uso, con clasificación de monto ocioso vs productivo.

    Suma el Costo anual (USD) y el Facturado (MXN) por nivel y etiqueta cada
    nivel como 'Productivo' (Muy Alto/Alto/Medio), 'Ocioso' (Bajo/Nulo) o
    'En adopción' (Reciente).
    """
    g = (
        final.groupby("Nivel de uso")
        .agg(**{
            "Usuarios": ("Correo electrónico", "count"),
            "Costo anual (USD)": ("Costo anual (USD)", "sum"),
            "Facturado (MXN)": ("Facturado (MXN)", "sum"),
        })
        .reset_index()
    )
    g["Clasificación"] = g["Nivel de uso"].map(CLASIF_NIVEL)
    g["_ord"] = g["Nivel de uso"].map({n: i for i, n in enumerate(NIVEL_ORDEN)})
    g = g.sort_values("_ord").drop(columns="_ord")
    total_mxn = g["Facturado (MXN)"].sum()
    g["% del facturado"] = (g["Facturado (MXN)"] / total_mxn * 100).round(1) \
        if total_mxn else 0
    g = g[["Nivel de uso", "Clasificación", "Usuarios",
           "Costo anual (USD)", "Facturado (MXN)", "% del facturado"]]
    total = pd.DataFrame([{
        "Nivel de uso": "TOTAL", "Clasificación": "",
        "Usuarios": g["Usuarios"].sum(),
        "Costo anual (USD)": g["Costo anual (USD)"].sum(),
        "Facturado (MXN)": g["Facturado (MXN)"].sum(),
        "% del facturado": 100.0,
    }])
    return pd.concat([g, total], ignore_index=True)


def resumen_aprovechamiento(costo_niv: pd.DataFrame) -> pd.DataFrame:
    """Agrupa el facturado en Productivo / Ocioso / En adopción."""
    base = costo_niv[costo_niv["Nivel de uso"] != "TOTAL"]
    g = (
        base.groupby("Clasificación")
        .agg(**{
            "Usuarios": ("Usuarios", "sum"),
            "Facturado (MXN)": ("Facturado (MXN)", "sum"),
            "Costo anual (USD)": ("Costo anual (USD)", "sum"),
        })
        .reset_index()
    )
    orden = {"Productivo": 0, "En adopción": 1, "Ocioso": 2}
    g["_o"] = g["Clasificación"].map(orden)
    g = g.sort_values("_o").drop(columns="_o").reset_index(drop=True)
    total = g["Facturado (MXN)"].sum()
    g["% del facturado"] = (g["Facturado (MXN)"] / total * 100).round(1) if total else 0
    return g


def monto_por_team(asig: pd.DataFrame) -> pd.DataFrame:
    """Monto por grupo (User Team) calculado desde los registros CRUDOS.

    Se calcula sobre la hoja "Asignación" sin consolidar (cada registro se
    suma en el equipo donde se facturó), para coincidir con la contabilidad /
    tabla dinámica administrativa. Un usuario con upgrade en dos equipos
    distintos aporta su monto a cada equipo tal como fue facturado.
    """
    df = asig.copy()
    df.columns = [str(c).strip() for c in df.columns]
    for c in ["Costo anual (USD)", "Facturado (USD)", "Facturado (MXN)",
              "Presupuesto Ops", "Cantidad lic"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    g = (
        df.groupby("User Team")
        .agg(**{
            "Registros": ("User Team", "count"),
            "Licencias": ("Cantidad lic", "sum"),
            "Costo anual (USD)": ("Costo anual (USD)", "sum"),
            "Facturado (USD)": ("Facturado (USD)", "sum"),
            "Facturado (MXN)": ("Facturado (MXN)", "sum"),
            "Presupuesto Ops": ("Presupuesto Ops", "sum"),
        })
        .reset_index()
        .sort_values("Facturado (MXN)", ascending=False)
    )
    total = pd.DataFrame([{
        "User Team": "TOTAL",
        "Registros": g["Registros"].sum(),
        "Licencias": g["Licencias"].sum(),
        "Costo anual (USD)": g["Costo anual (USD)"].sum(),
        "Facturado (USD)": g["Facturado (USD)"].sum(),
        "Facturado (MXN)": g["Facturado (MXN)"].sum(),
        "Presupuesto Ops": g["Presupuesto Ops"].sum(),
    }])
    return pd.concat([g, total], ignore_index=True)


def resumen_niveles(final: pd.DataFrame) -> pd.DataFrame:
    g = (
        final.groupby("Nivel de uso")
        .agg(**{
            "Usuarios": ("Correo electrónico", "count"),
            "% de uso promedio": ("% de uso", "mean"),
        })
        .reset_index()
        .sort_values("Nivel de uso")
    )
    g["% de uso promedio"] = g["% de uso promedio"].round(1)
    return g


def uso_por_team(final: pd.DataFrame) -> pd.DataFrame:
    """Promedio de % de uso por grupo (User Team)."""
    g = (
        final.groupby("User Team")
        .agg(**{
            "Usuarios": ("Correo electrónico", "count"),
            "% de uso promedio": ("% de uso", "mean"),
            "Días activos promedio": ("Days Active", "mean"),
        })
        .reset_index()
        .sort_values("% de uso promedio", ascending=False)
    )
    g["% de uso promedio"] = g["% de uso promedio"].round(1)
    g["Días activos promedio"] = g["Días activos promedio"].round(1)
    return g


def presupuesto_ops(consol: pd.DataFrame) -> pd.DataFrame:
    """Consumo de los usuarios con 'Presupuesto Ops' > 0, por fecha de asignación.

    Muestra, por cada fecha en que se asignaron licencias, la cantidad de
    usuarios y el monto (en pesos) correspondiente a esa fecha, más el total.
    """
    base = consol[pd.to_numeric(consol["Presupuesto Ops"], errors="coerce")
                  .fillna(0) > 0].copy()
    base["Fecha"] = pd.to_datetime(base["Fecha"], errors="coerce")
    g = (
        base.groupby(base["Fecha"].dt.date)
        .agg(**{
            "Usuarios": ("email_key", "nunique"),
            "Presupuesto Ops (MXN)": ("Presupuesto Ops", "sum"),
            "Facturado (MXN)": ("Facturado (MXN)", "sum"),
        })
        .reset_index()
        .rename(columns={"Fecha": "Fecha de asignación"})
        .sort_values("Fecha de asignación")
    )
    g["Fecha de asignación"] = pd.to_datetime(g["Fecha de asignación"])
    total = pd.DataFrame([{
        "Fecha de asignación": pd.NaT,
        "Usuarios": g["Usuarios"].sum(),
        "Presupuesto Ops (MXN)": g["Presupuesto Ops (MXN)"].sum(),
        "Facturado (MXN)": g["Facturado (MXN)"].sum(),
    }])
    out = pd.concat([g, total], ignore_index=True)
    return out


# --------------------------------------------------------------------------- #
# Verificación por consola
# --------------------------------------------------------------------------- #
def verificar(final: pd.DataFrame, consol: pd.DataFrame,
              usab: pd.DataFrame, fecha_corte):
    print("\n[5] Verificación de cálculos (muestra):")
    corte = pd.Timestamp(fecha_corte)
    muestra = final.head(4)
    for _, r in muestra.iterrows():
        f = r["Fecha"]
        cal = (corte - f).days if pd.notna(f) else None
        bus = np.busday_count(
            np.datetime64(pd.Timestamp(f).date(), "D"),
            np.datetime64(corte.date(), "D"),
        ) if pd.notna(f) else None
        pct = (r["Days Active"] / r["Días hábiles"] * 100) if r["Días hábiles"] else 0
        print(f"    · {r['Correo electrónico']}: "
              f"asig={f:%Y-%m-%d} | días_cal={r['Días calendario']}(esp {cal}) "
              f"| días_háb={r['Días hábiles']}(esp {int(bus)}) "
              f"| Days Active={int(r['Days Active'])} "
              f"| %uso={r['% de uso']}(esp {round(pct,1)}) "
              f"| nivel={r['Nivel de uso']} | amp={r['Amplitud de adopción (0-5)']}")

    # Chequeos de integridad global.
    print("\n    Chequeos de integridad:")
    assert (final["Días calendario"] >= 0).all(), "Días calendario negativos!"
    assert (final["Días hábiles"] >= 0).all(), "Días hábiles negativos!"
    assert (final["Amplitud de adopción (0-5)"].between(0, 5)).all(), "Amplitud fuera de rango!"
    # Días hábiles nunca debe superar los calendario.
    assert (final["Días hábiles"] <= final["Días calendario"]).all(), \
        "Días hábiles > Días calendario!"
    # Días sin uso (hábiles) nunca debe superar los días hábiles.
    assert (final["Días sin uso"] <= final["Días hábiles"]).all(), \
        "Días sin uso > Días hábiles!"
    # Ningún registro debe tener Cantidad lic = 0.
    assert (final["Cantidad lic"] > 0).all(), "Hay registros con Cantidad lic = 0!"
    niveles_validos = {"1.Muy Alto", "2.Alto", "3.Medio", "4.Bajo", "5.Nulo", "Reciente"}
    assert set(final["Nivel de uso"]).issubset(niveles_validos), "Nivel inválido!"
    # Ningún usuario "Reciente" debe tener 15+ días de licencia asignada.
    rec = final["Nivel de uso"] == "Reciente"
    assert (final.loc[rec, "Días calendario"] < DIAS_MIN_EVALUACION).all(), \
        "Un usuario 'Reciente' tiene 15+ días de licencia!"

    # Cuadre de costos: los campos que se SUMAN no deben alterar el total.
    # La consolidación suma por usuario y luego se excluyen los usuarios cuya
    # Cantidad lic total = 0, por lo que se compara contra el total original
    # de los usuarios que SÍ sobreviven al filtro.
    orig = pd.read_excel(ARCHIVO_ENTRADA_DEFAULT, sheet_name=HOJA_ASIGNACION)
    orig.columns = [str(c).strip() for c in orig.columns]
    orig["_k"] = normaliza_correo(orig["Correo electrónico"])
    orig["Cantidad lic"] = pd.to_numeric(orig["Cantidad lic"], errors="coerce").fillna(0)
    cant_por_correo = orig.groupby("_k")["Cantidad lic"].sum()
    vivos = set(cant_por_correo[cant_por_correo > 0].index)
    orig_vivos = orig[orig["_k"].isin(vivos)]
    excluidos = orig["_k"].nunique() - len(vivos)
    print(f"      (usuarios excluidos por Cantidad lic = 0: {excluidos})")
    for c in COLS_SUMA + ["Cantidad lic"]:
        o = pd.to_numeric(orig_vivos[c], errors="coerce").fillna(0).sum()
        n = pd.to_numeric(final[c], errors="coerce").fillna(0).sum()
        estado = "OK" if abs(o - n) < 0.01 else "DIFERENCIA"
        print(f"      - {c} (suma, usuarios vigentes): "
              f"original={o:,.2f} | consolidado={n:,.2f} -> {estado}")
        assert abs(o - n) < 0.01, f"El total de {c} cambió tras consolidar!"
    # Total bruto (todos los registros) para referencia contable.
    tot_mxn = pd.to_numeric(orig["Facturado (MXN)"], errors="coerce").fillna(0).sum()
    print(f"      Facturado (MXN) bruto (todos los registros): {tot_mxn:,.2f}")

    print(f"\n    Distribución por Nivel de uso:")
    print(final["Nivel de uso"].value_counts().sort_index().to_string())
    print("\n    Todos los chequeos de integridad pasaron correctamente. ✓")


# --------------------------------------------------------------------------- #
# 6. Exportación con formato
# --------------------------------------------------------------------------- #
def exportar(final, volum, usuarios_usab, montos, niveles, usoteam, presupops,
             costo_niv, aprov, fecha_corte):
    print(f"\n[6] Exportando a: {ARCHIVO_SALIDA}")
    with pd.ExcelWriter(ARCHIVO_SALIDA, engine="xlsxwriter",
                        datetime_format="yyyy-mm-dd") as writer:
        wb = writer.book

        # Formatos reutilizables.
        fmt_titulo = wb.add_format({
            "bold": True, "font_size": 14, "font_color": "#1F2937"})
        fmt_sub = wb.add_format({"italic": True, "font_color": "#6B7280"})
        fmt_hdr = wb.add_format({
            "bold": True, "bg_color": "#4338CA", "font_color": "white",
            "border": 1, "align": "center", "valign": "vcenter",
            "text_wrap": True})
        fmt_hdr_a = wb.add_format({
            "bold": True, "bg_color": "#4338CA", "font_color": "white",
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
        fmt_hdr_b = wb.add_format({
            "bold": True, "bg_color": "#0F766E", "font_color": "white",
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
        fmt_hdr_c = wb.add_format({
            "bold": True, "bg_color": "#B45309", "font_color": "white",
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
        fmt_money = wb.add_format({"num_format": "#,##0.00"})
        fmt_pct = wb.add_format({"num_format": "0.0"})
        fmt_int = wb.add_format({"num_format": "#,##0"})
        fmt_total = wb.add_format({"bold": True, "bg_color": "#E5E7EB", "top": 2})

        # ---------------- Hoja: Análisis Detallado (se escribe al final) ----
        def _hoja_detalle():
            hoja = "Análisis Detallado"
            startrow = 3
            final.to_excel(writer, sheet_name=hoja, startrow=startrow, index=False)
            ws = writer.sheets[hoja]
            ws.write(0, 0, "Análisis de Usabilidad de Licencias Claude", fmt_titulo)
            ws.write(1, 0, f"Fecha de corte: {pd.Timestamp(fecha_corte):%Y-%m-%d}  |  "
                           f"Usuarios: {len(final)}", fmt_sub)

            # Encabezados coloreados por bloque (A/B/C).
            bloque_b = {"Días calendario", "Días hábiles", "Días sin uso",
                        "Días desde últ. conexión"}
            bloque_c = {"% de uso", "Nivel de uso", "Amplitud de adopción (0-5)",
                        "Acción sugerida", "Consideraciones"}
            for col, name in enumerate(final.columns):
                if name in bloque_b:
                    f = fmt_hdr_b
                elif name in bloque_c:
                    f = fmt_hdr_c
                else:
                    f = fmt_hdr_a
                ws.write(startrow, col, name, f)

            # Anchos de columna aproximados.
            for col, name in enumerate(final.columns):
                ancho = min(max(len(str(name)) + 2, 12), 34)
                ws.set_column(col, col, ancho)
            ws.freeze_panes(startrow + 1, 2)
            ws.autofilter(startrow, 0, startrow + len(final), len(final.columns) - 1)

            # Escala de color en % de uso (semáforo).
            if "% de uso" in final.columns:
                ci = list(final.columns).index("% de uso")
                ws.conditional_format(
                    startrow + 1, ci, startrow + len(final), ci,
                    {"type": "3_color_scale",
                     "min_color": "#F8696B", "mid_color": "#FFEB84",
                     "max_color": "#63BE7B"})

        # ---------------- Hoja: Volumetría de Licencias ----------------
        hoja2 = "Volumetría Licencias"
        volum_out = volum.copy()
        volum_out.to_excel(writer, sheet_name=hoja2, startrow=2, index=False)
        ws2 = writer.sheets[hoja2]
        ws2.write(0, 0, "Volumetría: crecimiento de licencias asignadas en el tiempo",
                  fmt_titulo)
        for col, name in enumerate(volum_out.columns):
            ws2.write(2, col, name, fmt_hdr)
            ws2.set_column(col, col, 20)
        n = len(volum_out)

        # Gráfico de líneas: licencias y usuarios acumulados.
        chart = wb.add_chart({"type": "line"})
        cats = [hoja2, 3, 0, 2 + n, 0]
        col_lic_acum = list(volum_out.columns).index("Licencias acumuladas")
        col_usr_acum = list(volum_out.columns).index("Usuarios acumulados")
        chart.add_series({
            "name": "Licencias acumuladas",
            "categories": cats,
            "values": [hoja2, 3, col_lic_acum, 2 + n, col_lic_acum],
            "marker": {"type": "circle", "size": 5},
            "line": {"color": "#4338CA", "width": 2.25},
        })
        chart.add_series({
            "name": "Usuarios acumulados",
            "categories": cats,
            "values": [hoja2, 3, col_usr_acum, 2 + n, col_usr_acum],
            "marker": {"type": "square", "size": 5},
            "line": {"color": "#0F766E", "width": 2.25},
        })
        chart.set_title({"name": "Crecimiento acumulado de licencias"})
        chart.set_x_axis({"name": "Fecha de asignación", "num_format": "yyyy-mm-dd"})
        chart.set_y_axis({"name": "Acumulado"})
        chart.set_size({"width": 720, "height": 400})
        ws2.insert_chart(2, len(volum_out.columns) + 1, chart)

        # ---------------- Hoja: Usabilidad Usuarios ----------------
        # Todos los usuarios con autofiltro (permite filtrar por uno o varios
        # niveles de usabilidad de forma nativa en Excel).
        hoja3 = "Usabilidad Usuarios"
        usuarios_usab.to_excel(writer, sheet_name=hoja3, startrow=2, index=False)
        ws3 = writer.sheets[hoja3]
        ws3.write(0, 0, "Usabilidad de todos los usuarios "
                        "(filtra por nivel con el autofiltro)", fmt_titulo)
        for col, name in enumerate(usuarios_usab.columns):
            ws3.write(2, col, name, fmt_hdr)
            ws3.set_column(col, col, max(len(str(name)) + 2, 16))
        ws3.freeze_panes(3, 0)
        ws3.autofilter(2, 0, 2 + len(usuarios_usab), len(usuarios_usab.columns) - 1)

        # ---------------- Hoja: Resumen Niveles ----------------
        hoja5 = "Resumen Niveles"
        niveles.to_excel(writer, sheet_name=hoja5, startrow=2, index=False)
        ws5 = writer.sheets[hoja5]
        ws5.write(0, 0, "Distribución de usuarios por Nivel de uso", fmt_titulo)
        for col, name in enumerate(niveles.columns):
            ws5.write(2, col, name, fmt_hdr)
            ws5.set_column(col, col, max(len(str(name)) + 2, 18))
        chart3 = wb.add_chart({"type": "column"})
        nn = len(niveles)
        chart3.add_series({
            "name": "Usuarios",
            "categories": [hoja5, 3, 0, 2 + nn, 0],
            "values": [hoja5, 3, 1, 2 + nn, 1],
            "fill": {"color": "#0F766E"},
            "data_labels": {"value": True},
        })
        chart3.set_title({"name": "Usuarios por Nivel de uso"})
        chart3.set_size({"width": 640, "height": 380})
        ws5.insert_chart(2, len(niveles.columns) + 1, chart3)

        # ---------------- Hoja: % Uso por Team ----------------
        hoja6 = "% Uso por Team"
        usoteam.to_excel(writer, sheet_name=hoja6, startrow=2, index=False)
        ws6 = writer.sheets[hoja6]
        ws6.write(0, 0, "Promedio de % de uso por User Team", fmt_titulo)
        for col, name in enumerate(usoteam.columns):
            ws6.write(2, col, name, fmt_hdr)
            ws6.set_column(col, col, max(len(str(name)) + 2, 20))
        chart4 = wb.add_chart({"type": "bar"})
        mm = len(usoteam)
        col_pct = list(usoteam.columns).index("% de uso promedio")
        chart4.add_series({
            "name": "% de uso promedio",
            "categories": [hoja6, 3, 0, 2 + mm, 0],
            "values": [hoja6, 3, col_pct, 2 + mm, col_pct],
            "fill": {"color": "#4338CA"},
            "data_labels": {"value": True},
        })
        chart4.set_title({"name": "% de uso promedio por User Team"})
        chart4.set_size({"width": 720, "height": 420})
        ws6.insert_chart(2, len(usoteam.columns) + 1, chart4)

        # ---------------- Hoja: Costo por Nivel (ocioso vs productivo) ------
        hoja8 = "Costo por Nivel"
        # Colores de clasificación.
        col_prod = aprov.loc[aprov["Clasificación"] == "Productivo", "Facturado (MXN)"].sum()
        col_ocio = aprov.loc[aprov["Clasificación"] == "Ocioso", "Facturado (MXN)"].sum()
        col_adop = aprov.loc[aprov["Clasificación"] == "En adopción", "Facturado (MXN)"].sum()
        costo_niv.to_excel(writer, sheet_name=hoja8, startrow=4, index=False)
        ws8 = writer.sheets[hoja8]
        ws8.write(0, 0, "Costo por nivel de uso — monto ocioso vs productivo", fmt_titulo)
        ws8.write(1, 0, f"Productivo (Muy Alto/Alto/Medio): {col_prod:,.2f} MXN   |   "
                        f"Ocioso (Bajo/Nulo): {col_ocio:,.2f} MXN   |   "
                        f"En adopción (Reciente): {col_adop:,.2f} MXN", fmt_sub)
        for col, name in enumerate(costo_niv.columns):
            ws8.write(4, col, name, fmt_hdr)
            ws8.set_column(col, col, max(len(str(name)) + 2, 16))
        for cname in ["Costo anual (USD)", "Facturado (MXN)"]:
            ci = list(costo_niv.columns).index(cname)
            ws8.set_column(ci, ci, 18, fmt_money)
        # Fila TOTAL: header en la fila 4, datos desde la 5; la última fila
        # de datos (TOTAL) queda en 4 + len(costo_niv).
        total_row8 = 4 + len(costo_niv)
        for col in range(len(costo_niv.columns)):
            ws8.write(total_row8, col, costo_niv.iloc[-1, col], fmt_total)
        # Gráfico de columnas: Facturado (MXN) por nivel (sin TOTAL).
        chart8 = wb.add_chart({"type": "column"})
        nn8 = len(costo_niv) - 1  # excluye TOTAL
        col_mxn8 = list(costo_niv.columns).index("Facturado (MXN)")
        chart8.add_series({
            "name": "Facturado (MXN)",
            "categories": [hoja8, 5, 0, 4 + nn8, 0],
            "values": [hoja8, 5, col_mxn8, 4 + nn8, col_mxn8],
            "data_labels": {"value": True},
            "points": [
                {"fill": {"color": "#0ca30c"}}, {"fill": {"color": "#1baf7a"}},
                {"fill": {"color": "#fab219"}}, {"fill": {"color": "#ec835a"}},
                {"fill": {"color": "#d03b3b"}}, {"fill": {"color": "#64748b"}},
            ],
        })
        chart8.set_title({"name": "Facturado (MXN) por nivel de uso"})
        chart8.set_size({"width": 560, "height": 360})
        ws8.insert_chart(4, len(costo_niv.columns) + 1, chart8)

        # Tabla resumen de aprovechamiento + pastel productivo/ocioso.
        rstart = total_row8 + 2
        ws8.write(rstart - 1, 0, "Resumen de aprovechamiento", fmt_titulo)
        aprov.to_excel(writer, sheet_name=hoja8, startrow=rstart, index=False)
        for col, name in enumerate(aprov.columns):
            ws8.write(rstart, col, name, fmt_hdr)
        chart_pie = wb.add_chart({"type": "pie"})
        na = len(aprov)
        col_mxn_a = list(aprov.columns).index("Facturado (MXN)")
        color_map = {"Productivo": "#0ca30c", "Ocioso": "#d03b3b", "En adopción": "#64748b"}
        pts = [{"fill": {"color": color_map.get(c, "#4338CA")}}
               for c in aprov["Clasificación"]]
        chart_pie.add_series({
            "name": "Facturado (MXN)",
            "categories": [hoja8, rstart + 1, 0, rstart + na, 0],
            "values": [hoja8, rstart + 1, col_mxn_a, rstart + na, col_mxn_a],
            "points": pts,
            "data_labels": {"percentage": True, "category": True},
        })
        chart_pie.set_title({"name": "Monto ocioso vs productivo (MXN)"})
        chart_pie.set_size({"width": 460, "height": 340})
        ws8.insert_chart(rstart, len(aprov.columns) + 1, chart_pie)

        # ---------------- Hoja: Consumo Presup. Ops ----------------
        hoja7 = "Consumo Presup. Ops"
        n_usuarios_po = int(presupops.iloc[-1]["Usuarios"]) if len(presupops) else 0
        presupops.to_excel(writer, sheet_name=hoja7, startrow=3, index=False)
        ws7 = writer.sheets[hoja7]
        ws7.write(0, 0, "Consumo de usuarios con Presupuesto Ops > 0", fmt_titulo)
        ws7.write(1, 0, f"Usuarios con Presupuesto Ops: {n_usuarios_po}  |  "
                        f"Total en pesos: "
                        f"{presupops.iloc[-1]['Presupuesto Ops (MXN)']:,.2f} MXN", fmt_sub)
        for col, name in enumerate(presupops.columns):
            ws7.write(3, col, name, fmt_hdr)
            ws7.set_column(col, col, max(len(str(name)) + 2, 18))
        for cname in ["Presupuesto Ops (MXN)", "Facturado (MXN)"]:
            if cname in presupops.columns:
                ci = list(presupops.columns).index(cname)
                ws7.set_column(ci, ci, 20, fmt_money)
        # Resalta fila TOTAL (última).
        total_row7 = 3 + len(presupops)
        ws7.write(total_row7, 0, "TOTAL", fmt_total)
        for col in range(1, len(presupops.columns)):
            ws7.write(total_row7, col, presupops.iloc[-1, col], fmt_total)
        # Gráfico de columnas: Presupuesto Ops por fecha (sin TOTAL).
        chart5 = wb.add_chart({"type": "column"})
        pp = len(presupops) - 1  # excluye TOTAL
        col_po = list(presupops.columns).index("Presupuesto Ops (MXN)")
        chart5.add_series({
            "name": "Presupuesto Ops (MXN)",
            "categories": [hoja7, 4, 0, 3 + pp, 0],
            "values": [hoja7, 4, col_po, 3 + pp, col_po],
            "fill": {"color": "#0F766E"},
            "data_labels": {"value": True},
        })
        chart5.set_title({"name": "Presupuesto Ops (MXN) por fecha de asignación"})
        chart5.set_x_axis({"num_format": "yyyy-mm-dd"})
        chart5.set_size({"width": 720, "height": 400})
        ws7.insert_chart(3, len(presupops.columns) + 1, chart5)

        # ---------------- Hoja: Monto por User Team (al final) ----------------
        hoja4 = "Monto por User Team"
        montos.to_excel(writer, sheet_name=hoja4, startrow=2, index=False)
        ws4 = writer.sheets[hoja4]
        ws4.write(0, 0, "Monto de licencias por grupo (User Team) — desde "
                        "registros crudos (contabilidad)", fmt_titulo)
        for col, name in enumerate(montos.columns):
            ws4.write(2, col, name, fmt_hdr)
            ws4.set_column(col, col, max(len(str(name)) + 2, 16))
        for cname in ["Costo anual (USD)", "Facturado (USD)", "Facturado (MXN)",
                      "Presupuesto Ops"]:
            if cname in montos.columns:
                ci = list(montos.columns).index(cname)
                ws4.set_column(ci, ci, 18, fmt_money)
        total_row = 2 + len(montos)
        for col in range(len(montos.columns)):
            ws4.write(total_row, col, montos.iloc[-1, col], fmt_total)
        chart2 = wb.add_chart({"type": "bar"})
        m = len(montos) - 1  # excluye TOTAL
        col_team = list(montos.columns).index("User Team")
        col_mxn = list(montos.columns).index("Facturado (MXN)")
        chart2.add_series({
            "name": "Facturado (MXN)",
            "categories": [hoja4, 3, col_team, 2 + m, col_team],
            "values": [hoja4, 3, col_mxn, 2 + m, col_mxn],
            "fill": {"color": "#4338CA"},
        })
        chart2.set_title({"name": "Facturado (MXN) por User Team"})
        chart2.set_size({"width": 720, "height": 420})
        ws4.insert_chart(2, len(montos.columns) + 1, chart2)

        # ---------------- Hoja: Análisis Detallado (al final) ----------------
        _hoja_detalle()

    print(f"    - Archivo generado con hojas: Volumetría Licencias, "
          f"Usabilidad Usuarios, Resumen Niveles, % Uso por Team, "
          f"Costo por Nivel, Consumo Presup. Ops, Monto por User Team, "
          f"Análisis Detallado.")


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main():
    global ARCHIVO_ENTRADA_DEFAULT
    ruta = sys.argv[1] if len(sys.argv) > 1 else ARCHIVO_ENTRADA_DEFAULT
    ARCHIVO_ENTRADA_DEFAULT = ruta

    asig, usab = leer_hojas(ruta)
    consol_full = consolidar_asignacion(asig)          # 86: todos los usuarios
    df, fecha_corte = cruzar(consol_full, usab)
    calc = calcular(df, fecha_corte)
    final_full = ordenar_columnas(calc)                # master COMPLETO

    # Master VIGENTE (Cantidad lic > 0) para las vistas de usabilidad.
    vig = pd.to_numeric(final_full["Cantidad lic"], errors="coerce").fillna(0) > 0
    final = final_full[vig].reset_index(drop=True)
    consol = consol_full[pd.to_numeric(consol_full["Cantidad lic"],
             errors="coerce").fillna(0) > 0].reset_index(drop=True)

    # Indicadores / secciones.
    volum = volumetria_licencias(consol)
    usuarios_usab = usabilidad_usuarios(final)
    montos = monto_por_team(asig)  # crudo: coincide con la contabilidad
    niveles = resumen_niveles(final)
    usoteam = uso_por_team(final)
    presupops = presupuesto_ops(consol)
    # Costo por nivel: sobre el master COMPLETO para que el total sea
    # consistente con el costo real del proyecto (incluye licencias de equipo
    # sin usuario nombrado).
    costo_niv = costo_por_nivel(final_full)
    aprov = resumen_aprovechamiento(costo_niv)

    verificar(final, consol, usab, fecha_corte)
    exportar(final.drop(columns=["email_key"], errors="ignore"),
             volum.drop(columns=["email_key"], errors="ignore"),
             usuarios_usab, montos, niveles, usoteam, presupops,
             costo_niv, aprov, fecha_corte)

    print("\n✔ Proceso completado con éxito.")


if __name__ == "__main__":
    main()
