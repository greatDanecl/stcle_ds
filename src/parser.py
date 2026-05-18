"""
parser.py — STCLE Dashboard v2
Procesa todos los .xlsx en data/ y genera src/dashboard_data.json
KPIs calculados por mes y por cargo (CC / CCM):
  1. Vuelos nocturnos 00:30–05:30 (%)
  2. PSVNC – Pares de Servicios con Vuelo Nocturno Consecutivo + descanso < 10h
  3. Descansos efectivos DO+DR (%)
  4. Vacaciones VC (%)
  5. Licencias SICK+LNP+OOF (%)
  6. Standby B+ASB (%)
  7. Horas de vuelo promedio (block time)
  8. Distribución CC vs CCM
  9. Capacitaciones ADM+ASB+HSB+CRM (%)
 10. Semáforo de alertas por socio
"""

import os, json, re
from pathlib import Path
from datetime import datetime, timedelta
import pandas as pd
import numpy as np

DATA_DIR = Path(__file__).parent.parent / "data"
OUT_PATH = Path(__file__).parent / "dashboard_data.json"

# ─── helpers de datetime ───────────────────────────────────────────────
def parse_dt_numeric(date_val, time_val):
    try:
        d = pd.to_datetime(str(date_val)[:10])
        t = pd.to_timedelta(str(time_val))
        return d + t
    except:
        return pd.NaT

def parse_dt_feb(date_str, time_str):
    try:
        return pd.to_datetime(str(date_str) + " " + str(time_str), format="%d%b%Y %H:%M")
    except:
        return pd.NaT

# ─── carga archivos ───────────────────────────────────────────────────
def load_file(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)
    cols = list(df.columns)

    is_feb = "First Name" in cols
    if is_feb:
        df["Nombre completo"] = df["First Name"].fillna("") + " " + df["Last Name"].fillna("")
        df["dt_start"] = df.apply(lambda r: parse_dt_feb(r["Str Dt"], r["Str Tm"]), axis=1)
        df["dt_end"]   = df.apply(lambda r: parse_dt_feb(r["End Dt"], r["End Tm"]), axis=1)
    else:
        df["dt_start"] = df.apply(lambda r: parse_dt_numeric(r["Str Dt"], r["Str Tm"]), axis=1)
        df["dt_end"]   = df.apply(lambda r: parse_dt_numeric(r["End Dt"], r["End Tm"]), axis=1)

    # normalizar columnas esenciales
    if "sindicato" not in df.columns:
        df["sindicato"] = "CABLU"
    if "periodo" not in df.columns:
        # inferir del filename o de Str Dt
        df["periodo"] = df["dt_start"].dt.to_period("M").astype(str)

    # block time: puede ser datetime.time, fracción de día (float), o NaN
    if "Block Time" in df.columns:
        def parse_bt(v):
            import datetime as _dt
            if pd.isna(v): return 0.0
            if isinstance(v, _dt.time):
                return v.hour + v.minute / 60 + v.second / 3600
            try:
                f = float(v)
                return f * 24  # fracción de día
            except:
                return 0.0
        df["block_hours"] = df["Block Time"].apply(parse_bt)
    else:
        df["block_hours"] = 0.0

    keep = ["Staff Num", "Nombre completo", "Rank", "Activity", "dt_start", "dt_end",
            "block_hours", "sindicato", "periodo"]
    for c in keep:
        if c not in df.columns:
            df[c] = None
    return df[keep].dropna(subset=["dt_start", "dt_end"])

def load_all() -> pd.DataFrame:
    dfs = []
    for f in sorted(DATA_DIR.glob("*.xlsx")):
        try:
            dfs.append(load_file(f))
        except Exception as e:
            print(f"  SKIP {f.name}: {e}")
    if not dfs:
        raise RuntimeError("No hay archivos .xlsx en data/")
    big = pd.concat(dfs, ignore_index=True)
    # dedup de filas idénticas (mismo socio, actividad, dt_start)
    big = big.drop_duplicates(subset=["Staff Num", "Activity", "dt_start"])
    return big

# ─── clasificadores de actividad ──────────────────────────────────────
def is_flight(act):   return bool(re.match(r"^LA\d", str(act)))
def is_do(act):       return str(act).upper() in ("DO", "DR")
def is_vc(act):       return str(act).upper().startswith("VC")
def is_sick(act):     return str(act).upper() in ("SICK", "LNP", "OOF")
def is_standby(act):  return str(act).upper() in ("B","ASB1","ASB2","ASB3","ASC","BCM","BCN","HB7","HSB1","HSB2")
def is_training(act): return str(act).upper() in ("ADM","ASB1","ASB2","ASB3","HSB1","HSB2","CRM","CLA")

# Ausencia prolongada: cualquier razón (VC, licencias, permisos, capacitación fuera)
_AUSENCIA_SET = {"SICK","LNP","OOF","LMAT","LPAT","LIC","VACL","ML",
                 "INJ","SUS","REST","AUS","CLA","ADM"}
def is_ausencia(act):
    a = str(act).upper()
    return a in _AUSENCIA_SET or a.startswith("VC") or a.startswith("HSB")

def is_nocturnal(ts):
    h = ts.hour + ts.minute / 60
    return 0.5 <= h <= 5.5

# ─── KPI helper: porcentaje seguro ────────────────────────────────────
def pct(num, den):
    return round(num / den * 100, 1) if den else 0.0

# ─── PSVNC: Pares de Servicios con Vuelo Nocturno Consecutivo ─────────
#
# Definición correcta:
#   Un PAIRING es un grupo de vuelos del mismo período de servicio,
#   separados por < 8 h entre sí (multi-leg del mismo día de trabajo).
#
#   Un pairing "toca la franja nocturna 00:30–05:30" si:
#     - Algún vuelo sale en esa franja, O
#     - Algún vuelo llega en esa franja, O
#     - Algún vuelo cruza la medianoche (sale antes de 00:30 y llega al día siguiente)
#
#   PSVNC = pairing nocturno en día calendario D
#           seguido de pairing nocturno en día calendario D+1.

def _pairing_toca_franja(pairing_df: pd.DataFrame) -> bool:
    for _, r in pairing_df.iterrows():
        h_s = r["dt_start"].hour + r["dt_start"].minute / 60
        h_e = r["dt_end"].hour + r["dt_end"].minute / 60
        if 0.5 <= h_s <= 5.5:   return True  # sale en franja
        if 0.5 <= h_e <= 5.5:   return True  # llega en franja
        if r["dt_end"].date() > r["dt_start"].date():  return True  # cruza medianoche
    return False

def _get_pairings(fl_socio: pd.DataFrame, gap_hours: float = 8.0) -> list:
    """Agrupa vuelos en pairings (períodos de servicio) separados por >= gap_hours."""
    fl = fl_socio.sort_values("dt_start").reset_index(drop=True)
    if len(fl) == 0:
        return []
    pairings, cur = [], [0]
    for i in range(1, len(fl)):
        gap = (fl.iloc[i]["dt_start"] - fl.iloc[i-1]["dt_end"]).total_seconds() / 3600
        if gap >= gap_hours:
            pairings.append(fl.iloc[cur].copy())
            cur = []
        cur.append(i)
    if cur:
        pairings.append(fl.iloc[cur].copy())
    return pairings

def compute_psvnc(flights_df: pd.DataFrame) -> list:
    """
    Detecta pares de pairings nocturnos en días calendario consecutivos.
    Retorna lista de dicts con detalle de cada evento PSVNC.
    """
    records = []
    for staff, grp in flights_df.groupby("Staff Num"):
        pairings = _get_pairings(grp)
        for i in range(len(pairings) - 1):
            p1, p2 = pairings[i], pairings[i + 1]

            if not _pairing_toca_franja(p1): continue
            if not _pairing_toca_franja(p2): continue

            # Día calendario del inicio de cada pairing
            day_p1 = p1.iloc[0]["dt_start"].date()
            day_p2 = p2.iloc[0]["dt_start"].date()
            if (day_p2 - day_p1).days != 1:
                continue

            rest_h = (p2.iloc[0]["dt_start"] - p1.iloc[-1]["dt_end"]).total_seconds() / 3600
            records.append({
                "staff":    int(staff),
                "nombre":   str(p1.iloc[0]["Nombre completo"]),
                "rank":     str(p1.iloc[0]["Rank"]),
                # pairing 1
                "dia_p1":   day_p1.strftime("%d/%m/%Y"),
                "vuelos_p1": " → ".join(p1["Activity"].tolist()),
                "inicio_p1": p1.iloc[0]["dt_start"].strftime("%H:%M"),
                "fin_p1":    p1.iloc[-1]["dt_end"].strftime("%H:%M"),
                # pairing 2
                "dia_p2":   day_p2.strftime("%d/%m/%Y"),
                "vuelos_p2": " → ".join(p2["Activity"].tolist()),
                "inicio_p2": p2.iloc[0]["dt_start"].strftime("%H:%M"),
                "fin_p2":    p2.iloc[-1]["dt_end"].strftime("%H:%M"),
                # descanso entre pairings
                "rest_h":   round(rest_h, 2),
            })
    return records

# ─── semáforo por socio ───────────────────────────────────────────────
def semaforo(bt_h, noct_cnt, psvnc_cnt, do_pct, vc_days):
    score = 0
    if bt_h > 100: score += 2
    elif bt_h > 85: score += 1
    if noct_cnt >= 8: score += 2
    elif noct_cnt >= 4: score += 1
    if psvnc_cnt >= 3: score += 2
    elif psvnc_cnt >= 1: score += 1
    if do_pct < 15: score += 1
    if score >= 4: return "rojo"
    if score >= 2: return "amarillo"
    return "verde"

# ─── procesamiento por mes ────────────────────────────────────────────
def process_month(month_str: str, df: pd.DataFrame) -> dict:
    df = df.copy()
    flights = df[df["Activity"].apply(is_flight)].copy()
    flights = flights.drop_duplicates(subset=["Staff Num", "Activity", "dt_start"])

    # todos los socios únicos del mes
    all_staff = df[["Staff Num", "Nombre completo", "Rank"]].drop_duplicates("Staff Num")

    result = {"mes": month_str, "por_cargo": {}, "psvnc_detalle": []}

    for rank in ["CC", "CCM"]:
        staff_rank = all_staff[all_staff["Rank"] == rank]["Staff Num"].tolist()
        n_total = len(staff_rank)
        if n_total == 0:
            continue

        df_r = df[df["Staff Num"].isin(staff_rank)]
        fl_r = flights[flights["Staff Num"].isin(staff_rank)]
        days_in_period = 30  # aproximado; usamos 30 para todos

        # ── KPI 1: vuelos nocturnos ──────────────────────────────────
        total_vuelos = len(fl_r)
        vuelos_noct = fl_r[fl_r["dt_start"].apply(is_nocturnal)]
        n_vuelos_noct = len(vuelos_noct)
        pct_noct = pct(n_vuelos_noct, total_vuelos)
        socios_con_noct = vuelos_noct["Staff Num"].nunique()

        # distribución horaria de salidas nocturnas
        hora_dist = {}
        for _, row in vuelos_noct.iterrows():
            h = row["dt_start"].hour
            hora_dist[h] = hora_dist.get(h, 0) + 1

        # ── PSVNC ───────────────────────────────────────────────────
        psvnc_records = compute_psvnc(fl_r)
        n_psvnc = len(psvnc_records)
        socios_psvnc = len(set(r["staff"] for r in psvnc_records))

        # ── KPI 2: descansos DO+DR ───────────────────────────────────
        n_do_socios = df_r[df_r["Activity"].apply(is_do)]["Staff Num"].nunique()

        # ── KPI 3: vacaciones ────────────────────────────────────────
        n_vc_socios = df_r[df_r["Activity"].apply(is_vc)]["Staff Num"].nunique()

        # ── KPI 4: licencias ─────────────────────────────────────────
        n_sick_socios = df_r[df_r["Activity"].apply(is_sick)]["Staff Num"].nunique()
        n_sick_dias   = df_r[df_r["Activity"].apply(is_sick)].shape[0]

        # ── KPI 5: standby ───────────────────────────────────────────
        standby_rows = df_r[df_r["Activity"].apply(is_standby)]
        n_sb_socios  = standby_rows["Staff Num"].nunique()
        n_sb_dias    = len(standby_rows)
        total_dias   = len(df_r)
        pct_sb       = pct(n_sb_dias, total_dias)

        # ── KPI 6: block time promedio ───────────────────────────────
        bt_by_staff = fl_r.groupby("Staff Num")["block_hours"].sum()
        avg_bt = round(float(bt_by_staff.mean()), 1) if len(bt_by_staff) else 0.0
        socios_sobre_100h = int((bt_by_staff > 100).sum())
        socios_sobre_85h  = int((bt_by_staff > 85).sum())

        # ── KPI 7: capacitaciones ────────────────────────────────────
        n_cap_socios = df_r[df_r["Activity"].apply(is_training)]["Staff Num"].nunique()

        # ── KPI EQ: Equidad de distribución de horas ─────────────────
        # Base: socios con bt > 0 y sin ausencia prolongada (> 5 días)
        aus_by_staff = df_r[df_r["Activity"].apply(is_ausencia)].groupby("Staff Num").size()
        bt_all = {sid: float(bt_by_staff.get(sid, 0)) for sid in staff_rank}

        # Clasificar socios
        eq_excluidos = []   # ausencia > 5d (excluidos del promedio base)
        eq_sin_vuelo = []   # bt == 0 pero no por ausencia (excluidos del base)
        eq_base = []        # participan en el cálculo del promedio

        for sid in staff_rank:
            dias_aus = int(aus_by_staff.get(sid, 0))
            bt_h     = bt_all[sid]
            nombre   = str(all_staff[all_staff["Staff Num"] == sid].iloc[0]["Nombre completo"])
            entry    = {"staff": int(sid), "nombre": nombre, "bt_h": round(bt_h, 1), "dias_aus": dias_aus}
            if dias_aus > 5:
                eq_excluidos.append(entry)
            elif bt_h == 0:
                eq_sin_vuelo.append(entry)
            else:
                eq_base.append(entry)

        # Cálculo estadístico sobre la base
        bt_base_vals = np.array([e["bt_h"] for e in eq_base])
        if len(bt_base_vals) >= 2:
            eq_mean = float(np.mean(bt_base_vals))
            eq_std  = float(np.std(bt_base_vals, ddof=1))
            eq_cv   = round(eq_std / eq_mean * 100, 1)
            tol     = eq_mean * 0.09          # ±9% del promedio
            lo, hi  = eq_mean - tol, eq_mean + tol

            # Clasificar cada socio de la base
            sub_asig  = [e for e in eq_base if e["bt_h"] < lo]
            sobre_asig = [e for e in eq_base if e["bt_h"] > hi]
            dentro     = [e for e in eq_base if lo <= e["bt_h"] <= hi]

            eq_indice = round(len(dentro) / len(eq_base) * 100, 1)  # % dentro del rango

            # Histograma en buckets de 5h para visualización
            min_h = max(0, int(np.floor(bt_base_vals.min() / 5) * 5))
            max_h = int(np.ceil(bt_base_vals.max() / 5) * 5) + 5
            buckets = []
            for b in range(min_h, max_h, 5):
                cnt = int(((bt_base_vals >= b) & (bt_base_vals < b + 5)).sum())
                buckets.append({"desde": b, "hasta": b+5, "n": cnt})

            # Percentiles para diagrama de caja
            p10, p25, p50, p75, p90 = [round(float(x), 1)
                for x in np.percentile(bt_base_vals, [10, 25, 50, 75, 90])]
        else:
            eq_mean = eq_std = eq_cv = 0.0
            lo = hi = tol = 0.0
            eq_indice = 0.0
            sub_asig = sobre_asig = dentro = buckets = []
            p10 = p25 = p50 = p75 = p90 = 0.0

        equidad = {
            "n_base":          len(eq_base),
            "n_excluidos_aus": len(eq_excluidos),
            "n_sin_vuelo":     len(eq_sin_vuelo),
            "mean_h":          round(eq_mean, 1),
            "std_h":           round(eq_std, 1),
            "cv_pct":          eq_cv,                     # coeficiente de variación real
            "umbral_lo":       round(lo, 1),
            "umbral_hi":       round(hi, 1),
            "indice_equidad":  eq_indice,                 # % dentro de ±9%
            "n_dentro":        len(dentro),
            "n_sub":           len(sub_asig),             # por debajo del umbral
            "n_sobre":         len(sobre_asig),           # por encima del umbral
            "brecha_h":        round(float(bt_base_vals.max() - bt_base_vals.min()), 1) if len(bt_base_vals) else 0.0,
            "percentiles":     {"p10": p10, "p25": p25, "p50": p50, "p75": p75, "p90": p90},
            "histograma":      buckets,
            # listas de socios para tabla detalle
            "sub_asig":        sorted(sub_asig,  key=lambda x: x["bt_h"]),
            "sobre_asig":      sorted(sobre_asig, key=lambda x: x["bt_h"], reverse=True),
            "excluidos":       sorted(eq_excluidos, key=lambda x: x["nombre"]),
        }

        # ── KPI 8: semáforo individual ───────────────────────────────
        semaforos = {"verde": 0, "amarillo": 0, "rojo": 0}
        detalle_socios = []
        psvnc_by_staff = {}
        for r in psvnc_records:
            psvnc_by_staff[r["staff"]] = psvnc_by_staff.get(r["staff"], 0) + 1
        noct_by_staff = vuelos_noct.groupby("Staff Num").size().to_dict()

        for sid in staff_rank:
            staff_row = all_staff[all_staff["Staff Num"] == sid].iloc[0]
            nombre = str(staff_row["Nombre completo"])
            bt_h = float(bt_by_staff.get(sid, 0))
            noct_c = int(noct_by_staff.get(sid, 0))
            psvnc_c = int(psvnc_by_staff.get(sid, 0))
            do_rows = df_r[(df_r["Staff Num"] == sid) & df_r["Activity"].apply(is_do)]
            do_p = pct(len(do_rows), days_in_period)
            vc_rows = df_r[(df_r["Staff Num"] == sid) & df_r["Activity"].apply(is_vc)]
            sema = semaforo(bt_h, noct_c, psvnc_c, do_p, len(vc_rows))
            semaforos[sema] += 1
            detalle_socios.append({
                "staff": int(sid),
                "nombre": nombre,
                "rank": rank,
                "bt_h": round(bt_h, 1),
                "noct": noct_c,
                "psvnc": psvnc_c,
                "do_pct": do_p,
                "vc_dias": len(vc_rows),
                "sick_dias": int(df_r[(df_r["Staff Num"] == sid) & df_r["Activity"].apply(is_sick)].shape[0]),
                "sb_dias": int(df_r[(df_r["Staff Num"] == sid) & df_r["Activity"].apply(is_standby)].shape[0]),
                "sema": sema,
            })

        result["por_cargo"][rank] = {
            "n_socios": n_total,
            # KPI 1
            "total_vuelos": total_vuelos,
            "vuelos_nocturnos": n_vuelos_noct,
            "pct_nocturno": pct_noct,
            "socios_con_nocturno": socios_con_noct,
            "hora_dist_noct": hora_dist,
            # PSVNC
            "psvnc_total": n_psvnc,
            "psvnc_socios": socios_psvnc,
            # KPI 2 descansos
            "socios_con_do": int(n_do_socios),
            "pct_socios_do": pct(n_do_socios, n_total),
            # KPI 3 vacaciones
            "socios_con_vc": int(n_vc_socios),
            "pct_socios_vc": pct(n_vc_socios, n_total),
            # KPI 4 licencias
            "socios_con_sick": int(n_sick_socios),
            "dias_sick_total": int(n_sick_dias),
            "pct_socios_sick": pct(n_sick_socios, n_total),
            # KPI 5 standby
            "socios_con_sb": int(n_sb_socios),
            "dias_sb_total": int(n_sb_dias),
            "pct_dias_sb": pct_sb,
            # KPI 6 block time
            "avg_block_hours": avg_bt,
            "socios_sobre_100h": socios_sobre_100h,
            "socios_sobre_85h": socios_sobre_85h,
            # KPI 7 capacitaciones
            "socios_con_cap": int(n_cap_socios),
            "pct_socios_cap": pct(n_cap_socios, n_total),
            # KPI EQ equidad de distribución
            "equidad": equidad,
            # KPI 8 semáforo
            "semaforo": semaforos,
            # detalle socios
            "socios": sorted(detalle_socios, key=lambda x: x["nombre"]),
        }

        # guardar PSVNC detalle para este mes/cargo
        result["psvnc_detalle"].extend([r for r in psvnc_records if r["rank"] == rank])

    return result


# ─── main ─────────────────────────────────────────────────────────────
def main():
    print("📂 Cargando archivos...")
    big = load_all()
    print(f"   {len(big):,} filas cargadas, {big['Staff Num'].nunique()} socios únicos")

    # inferir mes de cada fila desde dt_start
    big["mes_key"] = big["dt_start"].dt.to_period("M").astype(str)

    # excluir socios sin base SCL si hay columna base
    # (en estos archivos no existe la columna, todos son SCL — ok)

    meses_disponibles = sorted(big["mes_key"].unique())
    print(f"   Meses: {meses_disponibles}")

    output = {
        "generado": datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
        "meses": meses_disponibles,
        "datos": {}
    }

    for mes in meses_disponibles:
        print(f"   Procesando {mes}...")
        df_mes = big[big["mes_key"] == mes]
        output["datos"][mes] = process_month(mes, df_mes)

    # guardar JSON
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, separators=(",", ":"))

    size_kb = OUT_PATH.stat().st_size / 1024
    print(f"✅ {OUT_PATH} — {size_kb:.1f} KB")

if __name__ == "__main__":
    main()
