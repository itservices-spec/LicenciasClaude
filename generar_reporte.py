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
ARCHIVO_ENTRADA_DEFAULT = "Usuarios de licencias Claude con CeCo - 07 Julio 2026.xlsx"
ARCHIVO_SALIDA = "Analisis_Usabilidad_Final.xlsx"

HOJA_ASIGNACION = "Asignación"
HOJA_USABILIDAD = "Usabilidad"

# Ranking de licencias: mayor número = licencia "más alta / vigente".
RANK_LICENCIAS = {"premium": 3, "pro": 2, "standard": 1}

# Columnas numéricas de costos que se SUMAN al consolidar duplicados.
COLS_COSTOS = [
    "Costo anual (USD)",
    "Facturado (USD)",
    "Facturado (MXN)",
    "Presupuesto Ops",
    "Cantidad lic",
]

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
    print(f"    - Asignación: {asig.shape[0]} filas x {asig.shape[1]} columnas")
    print(f"    - Usabilidad: {usab.shape[0]} filas x {usab.shape[1]} columnas")
    return asig, usab


# --------------------------------------------------------------------------- #
# 2 y 3. Normalización + consolidación de asignaciones duplicadas
# --------------------------------------------------------------------------- #
def consolidar_asignacion(asig: pd.DataFrame) -> pd.DataFrame:
    print("\n[2] Normalizando llave de cruce y consolidando duplicados...")
    df = asig.copy()
    df["email_key"] = normaliza_correo(df["Correo electrónico"])
    df["_rank_lic"] = df["Licencia"].apply(rank_licencia)

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

        # --- Consolidación de un correo con múltiples asignaciones (upgrade) ---
        n_consolidados += 1
        g_ord = g.sort_values("_rank_lic", ascending=False)
        base = g_ord.iloc[0].to_dict()  # fila con la licencia más alta

        # Suma de costos numéricos.
        for c in COLS_COSTOS:
            if c in df.columns:
                base[c] = g[c].sum()

        # Fecha de asignación más antigua.
        base["Fecha"] = g["Fecha"].min()

        # Nota de consolidación en Consideraciones.
        lics = " + ".join(g_ord["Licencia"].astype(str).tolist())
        nota = (f"Consolidado de {len(g)} asignaciones (upgrade). "
                f"Licencia vigente: {base['Licencia']}. Historial: {lics}.")
        prev = base.get("Consideraciones")
        base["Consideraciones"] = (
            f"{prev} | {nota}" if isinstance(prev, str) and prev.strip() else nota
        )
        filas.append(base)

    consol = pd.DataFrame(filas).drop(columns=["_rank_lic"])
    print(f"    - Correos consolidados (upgrades): {n_consolidados}")
    print(f"    - Asignación consolidada: {consol.shape[0]} filas "
          f"(antes {asig.shape[0]})")
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
    if nivel == "5.Nulo":
        return "Reasignar / cancelar"
    if nivel == "4.Bajo":
        return "Revisar"
    if pd.notna(dias_sin_uso) and dias_sin_uso > 21:
        return "Revisar (inactividad)"
    return "Mantener"


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

    # Días sin uso: corte - Last Active; si Last Active vacío pero hay
    # registro de asignación, se usa Días calendario.
    dias_sin_uso = (corte - d["Last Active"]).dt.days
    d["Días sin uso"] = dias_sin_uso.where(
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
    bloque_b = ["Días calendario", "Días hábiles", "Días sin uso"]
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


def usabilidad_baja(final: pd.DataFrame) -> pd.DataFrame:
    """Usuarios con usabilidad baja (Nivel Bajo o Nulo)."""
    mask = final["Nivel de uso"].isin(["4.Bajo", "5.Nulo"])
    cols = ["Nombre", "Correo electrónico", "User Team", "Nivel de uso",
            "% de uso", "Días sin uso", "Acción sugerida"]
    out = final.loc[mask, cols].sort_values(
        ["Nivel de uso", "% de uso"], ascending=[False, True]
    ).reset_index(drop=True)
    return out


def monto_por_team(consol: pd.DataFrame) -> pd.DataFrame:
    """Monto que corresponde a cada grupo (User Team)."""
    g = (
        consol.groupby("User Team")
        .agg(**{
            "Usuarios": ("email_key", "nunique"),
            "Licencias": ("Cantidad lic", "sum"),
            "Costo anual (USD)": ("Costo anual (USD)", "sum"),
            "Facturado (USD)": ("Facturado (USD)", "sum"),
            "Facturado (MXN)": ("Facturado (MXN)", "sum"),
        })
        .reset_index()
        .sort_values("Facturado (MXN)", ascending=False)
    )
    total = pd.DataFrame([{
        "User Team": "TOTAL",
        "Usuarios": g["Usuarios"].sum(),
        "Licencias": g["Licencias"].sum(),
        "Costo anual (USD)": g["Costo anual (USD)"].sum(),
        "Facturado (USD)": g["Facturado (USD)"].sum(),
        "Facturado (MXN)": g["Facturado (MXN)"].sum(),
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
    niveles_validos = {"1.Muy Alto", "2.Alto", "3.Medio", "4.Bajo", "5.Nulo"}
    assert set(final["Nivel de uso"]).issubset(niveles_validos), "Nivel inválido!"

    # Cuadre de costos: la consolidación no debe alterar el total.
    orig = pd.read_excel(ARCHIVO_ENTRADA_DEFAULT, sheet_name=HOJA_ASIGNACION)
    for c in COLS_COSTOS:
        o = pd.to_numeric(orig[c], errors="coerce").fillna(0).sum()
        n = pd.to_numeric(final[c], errors="coerce").fillna(0).sum()
        estado = "OK" if abs(o - n) < 0.01 else "DIFERENCIA"
        print(f"      - {c}: original={o:,.2f} | consolidado={n:,.2f} -> {estado}")
        assert abs(o - n) < 0.01, f"El total de {c} cambió tras consolidar!"

    print(f"\n    Distribución por Nivel de uso:")
    print(final["Nivel de uso"].value_counts().sort_index().to_string())
    print("\n    Todos los chequeos de integridad pasaron correctamente. ✓")


# --------------------------------------------------------------------------- #
# 6. Exportación con formato
# --------------------------------------------------------------------------- #
def exportar(final, volum, baja, montos, niveles, fecha_corte):
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

        # ---------------- Hoja principal: Análisis Detallado ----------------
        hoja = "Análisis Detallado"
        startrow = 3
        final.to_excel(writer, sheet_name=hoja, startrow=startrow, index=False)
        ws = writer.sheets[hoja]
        ws.write(0, 0, "Análisis de Usabilidad de Licencias Claude", fmt_titulo)
        ws.write(1, 0, f"Fecha de corte: {pd.Timestamp(fecha_corte):%Y-%m-%d}  |  "
                       f"Usuarios: {len(final)}", fmt_sub)

        # Encabezados coloreados por bloque (A/B/C).
        bloque_b = {"Días calendario", "Días hábiles", "Días sin uso"}
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

        # ---------------- Hoja: Usabilidad Baja ----------------
        hoja3 = "Usabilidad Baja"
        baja.to_excel(writer, sheet_name=hoja3, startrow=2, index=False)
        ws3 = writer.sheets[hoja3]
        ws3.write(0, 0, "Usuarios con usabilidad baja (Nivel Bajo / Nulo)", fmt_titulo)
        for col, name in enumerate(baja.columns):
            ws3.write(2, col, name, fmt_hdr)
            ws3.set_column(col, col, max(len(str(name)) + 2, 16))
        ws3.freeze_panes(3, 0)

        # ---------------- Hoja: Monto por User Team ----------------
        hoja4 = "Monto por User Team"
        montos.to_excel(writer, sheet_name=hoja4, startrow=2, index=False)
        ws4 = writer.sheets[hoja4]
        ws4.write(0, 0, "Monto de licencias por grupo (User Team)", fmt_titulo)
        for col, name in enumerate(montos.columns):
            ws4.write(2, col, name, fmt_hdr)
            ws4.set_column(col, col, max(len(str(name)) + 2, 16))
        # Formato de moneda en columnas de costos.
        for cname in ["Costo anual (USD)", "Facturado (USD)", "Facturado (MXN)"]:
            if cname in montos.columns:
                ci = list(montos.columns).index(cname)
                ws4.set_column(ci, ci, 18, fmt_money)
        # Resalta fila TOTAL.
        total_row = 2 + len(montos)
        for col in range(len(montos.columns)):
            ws4.write(total_row, col, montos.iloc[-1, col], fmt_total)

        # Gráfico de barras: Facturado (MXN) por team (sin TOTAL).
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

    print(f"    - Archivo generado con hojas: Análisis Detallado, "
          f"Volumetría Licencias, Usabilidad Baja, Monto por User Team, "
          f"Resumen Niveles.")


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main():
    global ARCHIVO_ENTRADA_DEFAULT
    ruta = sys.argv[1] if len(sys.argv) > 1 else ARCHIVO_ENTRADA_DEFAULT
    ARCHIVO_ENTRADA_DEFAULT = ruta

    asig, usab = leer_hojas(ruta)
    consol = consolidar_asignacion(asig)
    df, fecha_corte = cruzar(consol, usab)
    calc = calcular(df, fecha_corte)
    final = ordenar_columnas(calc)

    # Indicadores / secciones.
    volum = volumetria_licencias(consol)
    baja = usabilidad_baja(final)
    montos = monto_por_team(consol)
    niveles = resumen_niveles(final)

    verificar(final, consol, usab, fecha_corte)
    exportar(final.drop(columns=["email_key"], errors="ignore"),
             volum.drop(columns=["email_key"], errors="ignore"),
             baja, montos, niveles, fecha_corte)

    print("\n✔ Proceso completado con éxito.")


if __name__ == "__main__":
    main()
