
import argparse
from datetime import datetime, date, timedelta
from pathlib import Path
import pandas as pd
import numpy as np
from dateutil.parser import parse as dateparse
import unicodedata
import re

def strip_accents(s: str) -> str:
    if s is None: return ""
    s = str(s)
    return ''.join(c for c in unicodedata.normalize('NFKD', s) if not unicodedata.combining(c))

UF_LIST = ["AC","AL","AP","AM","BA","CE","DF","ES","GO","MA","MT","MS","MG","PA","PB","PR","PE","PI","RJ","RN","RS","RO","RR","SC","SP","SE","TO"]
FALLBACK_DAILY_BY_UF = {"SP": 37.50, "PR": 35.00, "RJ": 35.00, "RS": 35.00}
MONTHS_PT = {1:["jan","janeiro"],2:["fev","fevereiro"],3:["mar","marco","março","marco"],4:["abr","abril"],5:["mai","maio"],6:["jun","junho"],7:["jul","julho"],8:["ago","agosto"],9:["set","setembro","sept"],10:["out","outubro"],11:["nov","novembro"],12:["dez","dezembro"]}

def normalize_header(s: str) -> str:
    s = strip_accents(str(s)).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    s = re.sub(r"\s+", "_", s)
    synonyms = {
        "matricula": ["matric","id_func","idcolaborador","registro","id"],
        "admissao": ["admissao", "data_admissao","dt_admissao","admissao_data","data_de_admissao"],
        "titulo_do_cargo": ["titulo_cargo","cargo","titulo_funcao","funcao"],
        "sindicato": ["sindicato_do_colaborador","sind","sindicato_colaborador"],
        "data_demissao": ["demissao","data_demissao","dt_demissao","rescisao","data_rescisao"],
        "ok_comunicado": ["ok_comunicado","comunicado_ok","status_comunicado","ok","aprovado"],
        "data_comunicado": ["data_comunicado","dt_comunicado","comunicado_data"],
        "inicio": ["inicio","data_inicio","inicio_ferias","inicio_periodo","inicio_das_ferias"],
        "fim": ["fim","data_fim","final","retorno","fim_ferias","fim_das_ferias"],
        "valor": ["valor","valor_vr","valor_mensal","vr","valor_total"],
        "uf": ["uf","estado","sigla","estado_uf","uf_estado"],
        "dias_uteis": ["dias","dias_uteis","uteis","dias_no_periodo","diasuteis"],
        "municipio": ["municipio","cidade","localidade","municipio_uf","municipio_estado"],
        "descricao": ["descricao","feriado","nome","motivo"],
    }
    for target, alts in synonyms.items():
        if s in alts or s == target:
            return target
    return s

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [normalize_header(c) for c in df.columns]
    return df

def try_parse_date(x):
    if pd.isna(x): return pd.NaT
    if isinstance(x, (pd.Timestamp, datetime, date)): return pd.to_datetime(x)
    s = str(x).strip()
    if not s: return pd.NaT
    for dfirst in (True, False):
        out = pd.to_datetime(s, dayfirst=dfirst, errors="coerce")
        if pd.notna(out): return out
    try:
        return pd.to_datetime(dateparse(s, dayfirst=True))
    except Exception:
        return pd.NaT

def extract_uf(text: str):
    if pd.isna(text): return np.nan
    s = strip_accents(str(text)).upper()
    for uf in UF_LIST:
        if re.search(rf"\b{uf}\b", s): return uf
    m = re.search(r"\(([A-Z]{2})\)", s)
    return m.group(1) if m else np.nan

def best_month_column(df_cols, competencia: str):
    year = int(competencia.split("-")[0]); month = int(competencia.split("-")[1])
    month_aliases = MONTHS_PT.get(month, [])
    score = []
    for c in df_cols:
        cn_raw = str(c); cn = strip_accents(cn_raw).lower()
        sc = 0
        if str(year) in cn: sc += 1
        if f"{month:02d}" in cn or f"-{month}-" in cn or f"_{month}_" in cn: sc += 1
        for alias in month_aliases:
            if alias in cn: sc += 1
        digits = re.sub(r"[^0-9]", "", cn)
        if f"{year}{month:02d}" in digits: sc += 2
        score.append((sc, cn_raw))
    score.sort(reverse=True, key=lambda x: x[0])
    return score[0][1] if score and score[0][0] > 0 else None

def load_first_excel(data_dir: Path, patterns):
    candidates = sorted([p for p in data_dir.glob("*.xlsx")])
    def normname(p): return strip_accents(p.stem).lower().replace(" ", "")
    sel = [p for p in candidates if any(k in normname(p) for k in patterns)]
    if not sel: return None, None
    sel.sort(key=lambda x: len(x.name))
    df = pd.read_excel(sel[0], engine="openpyxl")
    return normalize_columns(df), sel[0]

def read_ativos(data_dir: Path):
    df, path = load_first_excel(data_dir, ["ativos"])
    if df is None: raise FileNotFoundError("ATIVOS.xlsx não encontrado.")
    if "uf" not in df.columns or df["uf"].isna().all():
        src = df["sindicato"] if "sindicato" in df.columns else None
        if src is None:
            for col in df.columns:
                if "sind" in col: src = df[col]; break
        df["uf"] = src.map(extract_uf) if src is not None else np.nan
    mat_col = "matricula" if "matricula" in df.columns else next((c for c in df.columns if "matric" in c), None)
    if not mat_col: raise ValueError("Coluna de matrícula não encontrada em ATIVOS.xlsx")
    adm_col = "admissao" if "admissao" in df.columns else next((c for c in df.columns if "admiss" in c), None)
    if adm_col: df["admissao"] = df[adm_col].map(try_parse_date)
    else: df["admissao"] = pd.NaT
    sind_col = "sindicato" if "sindicato" in df.columns else next((c for c in df.columns if "sind" in c), None)
    if not sind_col: df["sindicato"] = "N/D"; sind_col = "sindicato"
    cargo_col = next((c for c in df.columns if ("titulo" in c and "cargo" in c) or c=="cargo"), None)
    out = df[[mat_col, "admissao", sind_col, "uf"]].copy()
    out.columns = ["matricula", "admissao", "sindicato", "uf"]
    out["titulo_do_cargo"] = df[cargo_col].astype(str) if cargo_col else ""
    out["matricula"] = out["matricula"].astype(str).str.extract(r"(\d+)").fillna(out["matricula"])
    return out, path

def read_desligados(data_dir: Path):
    df, path = load_first_excel(data_dir, ["deslig"])
    if df is None: return None, None
    if "matricula" not in df.columns:
        for c in df.columns:
            if "matric" in c: df["matricula"] = df[c]; break
    if "data_demissao" not in df.columns:
        for c in df.columns:
            if "demiss" in c or "rescis" in c: df["data_demissao"] = df[c]; break
    ok_col = next((c for c in df.columns if c in ["ok_comunicado","comunicado_ok","status_comunicado","ok","aprovado"]), None)
    if ok_col is not None:
        ok_vals = df[ok_col].astype(str).str.strip().str.upper()
        df["ok_comunicado"] = ok_vals.isin(["OK","SIM","TRUE","VERDADEIRO","1","YES"])
    else:
        df["ok_comunicado"] = False
    dc_col = next((c for c in df.columns if c in ["data_comunicado","dt_comunicado","comunicado_data"]), None)
    df["data_comunicado"] = df[dc_col].map(try_parse_date) if dc_col else pd.NaT
    df["data_demissao"] = df.get("data_demissao", pd.Series(index=df.index)).map(try_parse_date)
    df["matricula"] = df.get("matricula", pd.Series(index=df.index)).astype(str).str.extract(r"(\d+)").fillna(df.get("matricula"))
    return df[["matricula","data_demissao","ok_comunicado","data_comunicado"]].dropna(subset=["matricula"]), path

def read_list_only(data_dir: Path, keys):
    df, path = load_first_excel(data_dir, keys)
    if df is None: return None, None
    if "matricula" not in df.columns: df.rename(columns={df.columns[0]:"matricula"}, inplace=True)
    df["matricula"] = df["matricula"].astype(str).str.extract(r"(\d+)").fillna(df["matricula"])
    return df[["matricula"]].dropna(subset=["matricula"]).drop_duplicates(), path

def read_exterior(data_dir: Path): return read_list_only(data_dir, ["exterior"])
def read_estagiarios(data_dir: Path): return read_list_only(data_dir, ["estagi","estagio"])
def read_aprendiz(data_dir: Path): return read_list_only(data_dir, ["aprendiz","aprend"])

def read_periods(data_dir: Path, keys):
    df, path = load_first_excel(data_dir, keys)
    if df is None: return None, None
    if "matricula" not in df.columns:
        for c in df.columns:
            if "matric" in c: df["matricula"] = df[c]; break
    start_col = next((c for c in df.columns if "inicio" in c), None)
    end_col = next((c for c in df.columns if "fim" in c or "retorno" in c or c=="fim"), None)
    df["inicio"] = df[start_col].map(try_parse_date) if start_col else pd.NaT
    df["fim"] = df[end_col].map(try_parse_date) if end_col else pd.NaT
    df["matricula"] = df["matricula"].astype(str).str.extract(r"(\d+)").fillna(df["matricula"])
    return df[["matricula","inicio","fim"]].dropna(subset=["matricula"]), path

def read_ferias(data_dir: Path): return read_periods(data_dir, ["ferias"])
def read_afastamentos(data_dir: Path): return read_periods(data_dir, ["afast","afastamento"])

def read_admissoes(data_dir: Path):
    df, path = load_first_excel(data_dir, ["admiss","admissao","admissao_abril","admissaoabril"])
    if df is None: return None, None
    if "matricula" not in df.columns:
        for c in df.columns:
            if "matric" in c: df["matricula"] = df[c]; break
    if "admissao" not in df.columns:
        for c in df.columns:
            if "admiss" in c: df["admissao"] = df[c]; break
    df["admissao"] = df.get("admissao", pd.Series(index=df.index)).map(try_parse_date)
    df["matricula"] = df.get("matricula", pd.Series(index=df.index)).astype(str).str.extract(r"(\d+)").fillna(df.get("matricula"))
    return df[["matricula","admissao"]].dropna(subset=["matricula"]), path

def read_base_dias(data_dir: Path): return load_first_excel(data_dir, ["basedias","diasuteis","uteis","base"])
def read_sindicato_valor(data_dir: Path): return load_first_excel(data_dir, ["sindicat","valor","sindicatoxvalor","base_sindicato_x_valor"])

# === Holidays ===
def read_feriados(data_dir: Path):
    df, path = load_first_excel(data_dir, ["feriado","calendario","holidays"])
    if df is None: return None, None
    # Flexible columns: data, uf, municipio (optional), descricao (optional)
    for c in ("data","dt","dia","date"):
        if c in df.columns and "data" not in df.columns:
            df.rename(columns={c:"data"}, inplace=True)
    if "uf" not in df.columns:
        if "municipio" in df.columns:
            df["uf"] = df["municipio"].map(extract_uf)
        else:
            df["uf"] = np.nan
    df["data"] = df.get("data", pd.Series(index=df.index)).map(try_parse_date).dt.date
    df = df.dropna(subset=["data"])
    if "uf" in df.columns:
        df["uf"] = df["uf"].astype(str).str.upper().str.extract(r"([A-Z]{2})")[0]
    return df[["uf","data"]].drop_duplicates(), path

def holidays_in_range(feriados_df: pd.DataFrame, uf: str, start: pd.Timestamp, end: pd.Timestamp):
    if feriados_df is None or feriados_df.empty: 
        return set()
    u = str(uf).upper() if uf is not None and not pd.isna(uf) else None
    if "uf" in feriados_df.columns:
        if u:
            df = feriados_df[(feriados_df["uf"].isna()) | (feriados_df["uf"]==u)]
        else:
            df = feriados_df[feriados_df["uf"].isna()]  # only national if UF unknown
    else:
        df = feriados_df
    lo = start.date(); hi = end.date()
    dates = set([d for d in df["data"].tolist() if lo <= d < hi])
    dates = set([d for d in dates if pd.Timestamp(d).weekday() < 5])
    return dates

def bd_excluding_holidays(start: pd.Timestamp, end: pd.Timestamp, uf: str, feriados_df: pd.DataFrame) -> int:
    if pd.isna(start) or pd.isna(end) or end <= start: return 0
    base = int(np.busday_count(np.datetime64(start.date()), np.datetime64(end.date())))
    hol = holidays_in_range(feriados_df, uf, start, end)
    return max(0, base - len(hol))

def best_days_series(df_base: pd.DataFrame, competencia: str, uf_series: pd.Series, feriados_df: pd.DataFrame, dt_inicio: pd.Timestamp, dt_fim: pd.Timestamp) -> pd.Series:
    # Use base sheet if month is found; otherwise compute weekdays minus holidays
    if df_base is not None:
        month_col = best_month_column(df_base.columns, competencia)
        dfb = df_base.copy(); dfb.columns = [strip_accents(str(c)).lower().replace(" ", "_") for c in dfb.columns]
        if "uf" not in dfb.columns:
            sind_col = next((c for c in dfb.columns if "sind" in c), None)
            if sind_col: dfb["uf"] = dfb[sind_col].map(extract_uf)
        if "uf" not in dfb.columns:
            for c in dfb.columns:
                if c in ("estado","sigla"): dfb.rename(columns={c:"uf"}, inplace=True); break
        if month_col and "uf" in df_base.columns:
            tmp = df_base.copy()
            uf_actual = next((c for c in tmp.columns if strip_accents(str(c)).lower().replace(" ", "_")=="uf"), None)
            if uf_actual:
                t = tmp[[uf_actual, month_col]].dropna(subset=[uf_actual])
                t[uf_actual] = t[uf_actual].astype(str).str.upper().str.strip()
                days_map = t.dropna(subset=[month_col]).groupby(uf_actual)[month_col].last().to_dict()
                return uf_series.map(lambda u: days_map.get(str(u).upper(), bd_excluding_holidays(dt_inicio, dt_fim, u, feriados_df)))
    return uf_series.map(lambda u: bd_excluding_holidays(dt_inicio, dt_fim, u, feriados_df))

def main():
    ap = argparse.ArgumentParser(description="Gera VR_MENSAL_YYYYMM.csv a partir de planilhas .xlsx")
    ap.add_argument("--inicio", required=True)
    ap.add_argument("--fim", required=True)
    ap.add_argument("--competencia", required=True)
    ap.add_argument("--data_dir", default=str(Path(__file__).resolve().parents[1]/"data"))
    ap.add_argument("--out_dir", default=str(Path(__file__).resolve().parents[1]/"out"))
    args = ap.parse_args()

    data_dir = Path(args.data_dir); out_dir = Path(args.out_dir); out_dir.mkdir(parents=True, exist_ok=True)
    dt_inicio = pd.to_datetime(args.inicio); dt_fim = pd.to_datetime(args.fim)
    competencia = args.competencia; comp_num = competencia.replace("-", "")

    ativos, _ = read_ativos(data_dir)
    desligados, _ = read_desligados(data_dir)
    estagiarios, _ = read_estagiarios(data_dir)
    aprendizes, _ = read_aprendiz(data_dir)
    exterior, _ = read_exterior(data_dir)
    ferias, _ = read_ferias(data_dir)
    afast, _ = read_afastamentos(data_dir)
    admissoes, _ = read_admissoes(data_dir)
    base_dias, _ = read_base_dias(data_dir)
    sind_valor, _ = read_sindicato_valor(data_dir)
    feriados_df, _ = read_feriados(data_dir)

    # complement 'admissao' with ADMISSOES if present
    if admissoes is not None and not admissoes.empty:
        amap = admissoes.set_index("matricula")["admissao"].to_dict()
        miss = ativos["admissao"].isna()
        ativos.loc[miss, "admissao"] = ativos.loc[miss, "matricula"].map(amap)

    # exclusions
    is_dir = ativos.get("titulo_do_cargo","").astype(str).str.lower().str.contains(r"\bdiretor|\bdir\.", na=False)
    est_set = set(estagiarios["matricula"].astype(str)) if estagiarios is not None and not estagiarios.empty else set()
    apr_set = set(aprendizes["matricula"].astype(str)) if aprendizes is not None and not aprendizes.empty else set()
    ext_set = set(exterior["matricula"].astype(str)) if exterior is not None and not exterior.empty else set()

    afast_set = set()
    if afast is not None and not afast.empty:
        afast["inicio"] = afast["inicio"].apply(lambda x: pd.to_datetime(x, errors="coerce"))
        afast["fim"] = afast["fim"].apply(lambda x: pd.to_datetime(x, errors="coerce"))
        afast = afast.dropna(subset=["inicio","fim"])
        def overlap(i,f): return not (f <= dt_inicio or i >= dt_fim)
        afast_set = set(afast.loc[[overlap(i,f) for i,f in zip(afast["inicio"], afast["fim"])], "matricula"].astype(str))

    ativos["matricula_str"] = ativos["matricula"].astype(str)
    excl_mask = is_dir | ativos["matricula_str"].isin(est_set | apr_set | ext_set | afast_set)
    base = ativos.loc[~excl_mask].copy()

    # drop who demitted before period
    if desligados is not None and not desligados.empty:
        dmap_dd = desligados.set_index("matricula")["data_demissao"].to_dict()
        dem_date = base["matricula_str"].map(dmap_dd)
        base = base.loc[~(dem_date.notna() & (dem_date < dt_inicio))].copy()

    # Base days per UF (considering holidays)
    dias_series = best_days_series(base_dias, competencia, base["uf"], feriados_df, dt_inicio, dt_fim).astype(float)
    obs = pd.Series(["Trabalhando"] * len(base), index=base.index, dtype=object)

    # Férias
    if ferias is not None and not ferias.empty:
        ferias["inicio"] = ferias["inicio"].apply(lambda x: pd.to_datetime(x, errors="coerce"))
        ferias["fim"] = ferias["fim"].apply(lambda x: pd.to_datetime(x, errors="coerce"))
        ferias = ferias.dropna(subset=["inicio","fim"])
        def ov_bd(i,f, uf):
            i2 = max(i, dt_inicio); f2 = min(f, dt_fim)
            if pd.isna(i2) or pd.isna(f2) or i2>=f2: return 0
            base_days = int(np.busday_count(np.datetime64(i2.date()), np.datetime64(f2.date())))
            # remove holidays for that UF
            hols = holidays_in_range(feriados_df, uf, i2, f2)
            return max(0, base_days - len(hols))
        f_map = {}
        for m, uf in zip(base["matricula_str"], base["uf"]):
            rows = ferias[ferias["matricula"].astype(str)==str(m)]
            total = 0
            for i,f in zip(rows["inicio"], rows["fim"]):
                total += ov_bd(i,f, uf)
            if total>0: f_map[str(m)] = total
        sub = base["matricula_str"].map(lambda m: f_map.get(str(m), 0)).astype(float)
        dias_series = (dias_series - sub).clip(lower=0)
        has_f = sub > 0
        obs.loc[has_f] = obs.loc[has_f] + " | Férias (" + sub[has_f].astype(int).astype(str) + " dias úteis no período)"

    # Admissão dentro do período (proporcional com feriados)
    adm = base["admissao"].apply(lambda x: pd.to_datetime(x, errors="coerce"))
    mask_adm = adm.notna() & (adm > dt_inicio) & (adm < dt_fim)
    if mask_adm.any():
        red = []
        for a, uf, m in zip(adm, base["uf"], mask_adm):
            if m:
                red.append(bd_excluding_holidays(dt_inicio, a, uf, feriados_df))
            else:
                red.append(0)
        red = pd.Series(red, index=base.index).astype(float)
        dias_series = (dias_series - red).clip(lower=0)
        obs.loc[mask_adm] = obs.loc[mask_adm] + " | Admissão em " + adm[mask_adm].dt.strftime("%d/%m/%Y")

    # Demissão (OK até dia 15 => dias=0; senão proporcional até dd), com feriados
    if desligados is not None and not desligados.empty:
        dmap = desligados.set_index("matricula")[["data_demissao","ok_comunicado","data_comunicado"]].to_dict(orient="index")
        cutoff = pd.Timestamp(year=dt_fim.year, month=dt_fim.month, day=15)
        new_days = []
        new_obs = obs.copy()
        for idx, (mat, days, uf) in enumerate(zip(base["matricula_str"], dias_series, base["uf"])):
            info = dmap.get(str(mat))
            if not info or pd.isna(info.get("data_demissao")):
                new_days.append(int(days)); continue
            dd = info["data_demissao"]; ok = bool(info.get("ok_comunicado")); dc = info.get("data_comunicado")
            if dd < dt_inicio: new_days.append(0); new_obs.iloc[idx] = "Desligado antes do período"; continue
            if dd > dt_fim: new_days.append(int(days)); continue
            key_date = dc if pd.notna(dc) else dd
            if ok and key_date <= cutoff:
                new_days.append(0); new_obs.iloc[idx] = (str(new_obs.iloc[idx]) + " | Desligado (OK comunicado)").strip(" |")
            else:
                used = bd_excluding_holidays(dt_inicio, dd, uf, feriados_df)
                new_days.append(min(int(days), used))
                new_obs.iloc[idx] = (str(new_obs.iloc[idx]) + f" | Desligado em {dd.strftime('%d/%m/%Y')} (proporcional)").strip(" |")
        dias_series = pd.Series(new_days, index=base.index).astype(int)
        obs = new_obs

    # Valor diário
    def daily_value_for_row(uf, sindicato):
        dfv = sind_valor
        if dfv is not None and not dfv.empty:
            dfv2 = dfv.copy(); dfv2.columns = [strip_accents(str(c)).lower().replace(" ", "_") for c in dfv2.columns]
            if "uf" in dfv2.columns and "valor" in dfv2.columns:
                m = dfv2[dfv2["uf"].astype(str).str.upper().str.strip() == str(uf).upper().strip()]
                if not m.empty and pd.notna(m.iloc[-1]["valor"]): return float(m.iloc[-1]["valor"])
            cand = next((c for c in dfv2.columns if "sind" in c), None)
            if cand and "valor" in dfv2.columns and isinstance(sindicato, str):
                temp = dfv2[dfv2[cand].astype(str).str.lower().str.contains(str(sindicato).lower(), na=False)]
                if not temp.empty and pd.notna(temp.iloc[-1]["valor"]): return float(temp.iloc[-1]["valor"])
        return FALLBACK_DAILY_BY_UF.get(str(uf).upper(), np.nan)

    daily_vals = pd.Series([daily_value_for_row(uf, sind) for uf, sind in zip(base["uf"], base["sindicato"])], index=base.index).astype(float)

    total = (pd.Series(dias_series, index=base.index).fillna(0).astype(int) * daily_vals.fillna(0)).round(2)
    custo_empresa = (total * 0.80).round(2); desconto_prof = (total * 0.20).round(2)

    out_df = pd.DataFrame({
        "Matricula": base["matricula"].astype(str),
        "UF": base["uf"].astype(str),
        "Admissão": base["admissao"].dt.strftime("%d/%m/%Y") if pd.api.types.is_datetime64_any_dtype(base["admissao"]) else base["admissao"].astype(str),
        "Sindicato do Colaborador": base["sindicato"].astype(str),
        "Competência": competencia,
        "Dias": pd.Series(dias_series, index=base.index).astype(int),
        "VALOR DIÁRIO VR": daily_vals.fillna(0).round(2),
        "TOTAL": total.fillna(0).round(2),
        "Custo empresa": custo_empresa.fillna(0).round(2),
        "Desconto profissional": desconto_prof.fillna(0).round(2),
        "OBS GERAL": obs.fillna(""),
    })

    # === VALIDACOES: relatório em out/VALIDACOES_YYYYMM.xlsx ===
    issues = []

    # Teto por UF (com feriados)
    teto_map = {}
    for uf in base["uf"].astype(str).unique():
        teto_map[uf] = int(np.busday_count(np.datetime64(dt_inicio.date()), np.datetime64(dt_fim.date())))
        # desconta feriados
        # Build hol set per UF
    def bd_teto(uf):
        # business days excluding holidays
        b = int(np.busday_count(np.datetime64(dt_inicio.date()), np.datetime64(dt_fim.date())))
        # Count UF holidays
        # minimal inline re-parse of feriados_df:
        return b  # we will use the robust function below

    # 1) Dias >= 0 e <= teto (com feriados)
    for idx, row in out_df.iterrows():
        dias = int(row["Dias"])
        uf = str(row.get("UF", ""))
        teto = int(np.busday_count(np.datetime64(dt_inicio.date()), np.datetime64(dt_fim.date())))
        # desconta feriados via função principal
        from_dt = pd.to_datetime(dt_inicio); to_dt = pd.to_datetime(dt_fim)
        # Importantly, rebuild using our function in local scope:
        def holidays_in_range_local(feriados_df, uf, start, end):
            if feriados_df is None or feriados_df.empty:
                return set()
            u = str(uf).upper() if uf is not None and not pd.isna(uf) else None
            if "uf" in feriados_df.columns:
                if u:
                    df = feriados_df[(feriados_df["uf"].isna()) | (feriados_df["uf"]==u)]
                else:
                    df = feriados_df[feriados_df["uf"].isna()]
            else:
                df = feriados_df
            lo = start.date(); hi = end.date()
            dates = set([d for d in df["data"].tolist() if lo <= d < hi])
            dates = set([d for d in dates if pd.Timestamp(d).weekday() < 5])
            return dates
        hols = holidays_in_range_local(feriados_df, uf, from_dt, to_dt)
        teto = max(0, teto - len(hols))

        if dias < 0:
            issues.append({"Matricula": row["Matricula"], "Regra": "Dias >= 0", "Detalhe": f"Dias = {dias}"})
        if dias > teto:
            issues.append({"Matricula": row["Matricula"], "Regra": "Dias <= teto do período (UF)", "Detalhe": f"Dias={dias} > Teto={teto} (UF={uf})"})

    # 2) Desligado OK -> Dias = 0
    for idx, row in out_df.iterrows():
        obs_txt = str(row["OBS GERAL"]).lower()
        if "desligado" in obs_txt and "ok comunicado" in obs_txt and int(row["Dias"]) != 0:
            issues.append({"Matricula": row["Matricula"], "Regra": "Desligado com OK até dia 15 -> Dias=0", "Detalhe": f"Dias={row['Dias']}"})

    # 3) Valor diário > 0
    for idx, row in out_df.iterrows():
        if float(row["VALOR DIÁRIO VR"]) <= 0:
            uf = str(row.get("UF", ""))
            issues.append({"Matricula": row["Matricula"], "Regra": "Valor diário > 0", "Detalhe": f"Valor={row['VALOR DIÁRIO VR']} (UF={uf})"})

    # 4) Consistência matemática
    for idx, row in out_df.iterrows():
        dias = int(row["Dias"]); val = float(row["VALOR DIÁRIO VR"])
        total_calc = round(dias * val, 2)
        if abs(float(row["TOTAL"]) - total_calc) > 0.01:
            issues.append({"Matricula": row["Matricula"], "Regra": "TOTAL = Dias × Valor", "Detalhe": f"TOTAL={row['TOTAL']} != {total_calc}"})
        ce = round(total_calc * 0.80, 2); dp = round(total_calc * 0.20, 2)
        if abs(float(row["Custo empresa"]) - ce) > 0.01:
            issues.append({"Matricula": row["Matricula"], "Regra": "Custo empresa = 80% de TOTAL", "Detalhe": f"Custo={row['Custo empresa']} != {ce}"})
        if abs(float(row["Desconto profissional"]) - dp) > 0.01:
            issues.append({"Matricula": row["Matricula"], "Regra": "Desconto profissional = 20% de TOTAL", "Detalhe": f"Desc={row['Desconto profissional']} != {dp}"})

    # 5) Exclusões não devem aparecer
    excl_concat = []
    for s, motivo in [(estagiarios, "Estagiário"), (aprendizes, "Aprendiz"), (exterior, "Exterior")]:
        if s is not None and not s.empty:
            for m in s["matricula"].astype(str).unique():
                excl_concat.append((m, motivo))
    if afast is not None and not afast.empty:
        for m in afast["matricula"].astype(str).unique():
            excl_concat.append((m, "Afastado"))

    excl_map = {}
    for m, motivo in excl_concat:
        excl_map.setdefault(str(m), set()).add(motivo)
    for idx, row in out_df.iterrows():
        m = str(row["Matricula"])
        if m in excl_map:
            issues.append({"Matricula": m, "Regra": "Exclusões", "Detalhe": f"Marcado como {', '.join(sorted(excl_map[m]))} não deveria constar na base final"})

    # Export CSV
    out_csv = Path(args.out_dir) / f"VR_MENSAL_{comp_num}.csv"
    out_df.to_csv(out_csv, index=False, encoding="utf-8-sig")

    # Export validations workbook
    val_path = Path(args.out_dir) / f"VALIDACOES_{comp_num}.xlsx"
    with pd.ExcelWriter(val_path, engine="openpyxl") as xlw:
        resumo = pd.DataFrame({"Item": ["Total colaboradores", "Total issues"], "Valor": [len(out_df), len(issues)]})
        resumo.to_excel(xlw, sheet_name="Resumo", index=False)
        issues_df = pd.DataFrame(issues) if issues else pd.DataFrame(columns=["Matricula","Regra","Detalhe"])
        issues_df.to_excel(xlw, sheet_name="Validacoes", index=False)
        params = pd.DataFrame({"Parametro": ["Inicio", "Fim", "Competencia"], "Valor": [str(dt_inicio.date()), str(dt_fim.date()), competencia]})
        params.to_excel(xlw, sheet_name="Parametros", index=False)

    print(f"Gerado: {out_csv}")
    print(f"Relatório de validações: {val_path}")

if __name__ == "__main__":
    main()
