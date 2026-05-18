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

def is_nocturnal(ts):
    h = ts.hour + ts.minute / 60
    return 0.5 <= h <= 5.5

# ─── KPI helper: porcentaje seguro ────────────────────────────────────
def pct(num, den):
    return round(num / den * 100, 1) if den else 0.0

# ─── PSVNC: Pares de Servicios con Vuelo Nocturno Consecutivo ─────────
def compute_psvnc(flights_df: pd.DataFrame) -> list:
    """
    Para cada socio, busca pares de vuelos en días consecutivos
    donde al menos uno es nocturno Y el descanso entre ellos < 10h.
    Retorna lista de dicts con detalle.
    """
    records = []
    for staff, grp in flights_df.groupby("Staff Num"):
        g = grp.sort_values("dt_start").reset_index(drop=True)
        for i in range(len(g) - 1):
            r1, r2 = g.iloc[i], g.iloc[i+1]
            rest_h = (r2["dt_start"] - r1["dt_end"]).total_seconds() / 3600
            date_diff = (r2["dt_start"].date() - r1["dt_start"].date()).days
            if date_diff != 1:
                continue
            h1 = r1["dt_start"].hour + r1["dt_start"].minute / 60
            h2 = r2["dt_start"].hour + r2["dt_start"].minute / 60
            noct1 = 0.5 <= h1 <= 5.5
            noct2 = 0.5 <= h2 <= 5.5
            if not (noct1 or noct2):
                continue
            if rest_h < 0 or rest_h >= 10:
                continue
            records.append({
                "staff": int(staff),
                "nombre": str(r1["Nombre completo"]),
                "rank": str(r1["Rank"]),
                "flight1": str(r1["Activity"]),
                "start1": r1["dt_start"].strftime("%d/%m %H:%M"),
                "end1": r1["dt_end"].strftime("%H:%M"),
                "flight2": str(r2["Activity"]),
                "start2": r2["dt_start"].strftime("%d/%m %H:%M"),
                "end2": r2["dt_end"].strftime("%H:%M"),
                "rest_h": round(rest_h, 2),
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
