
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
def read_sindicato_valor(data_dir: Path): return load_first_excel(data_dir, ["sindicatovalor","sindicatoxvalor","valor"])

def best_days_series(df_base: pd.DataFrame, competencia: str, uf_series: pd.Series, fallback_days: int, sindicato_series: pd.Series = None) -> pd.Series:
    if df_base is None or df_base.empty:
        return pd.Series(fallback_days, index=uf_series.index)
    dfb = df_base.copy()
    dfb.columns = [strip_accents(str(c)).lower().strip().replace(" ", "_") for c in dfb.columns]
    def _month_col(cols, competencia):
        comp = competencia.replace("-", "")
        yy = competencia[:4]; mm = competencia[-2:]
        tokens = {comp, competencia, f"{mm}/{yy}", f"{mm}{yy}", f"{yy}{mm}"}
        months_pt = {"01":["jan","janeiro"],"02":["fev","fevereiro"],"03":["mar","marco","março"],"04":["abr","abril"],"05":["mai","maio"],"06":["jun","junho"],"07":["jul","julho"],"08":["ago","agosto"],"09":["set","setembro"],"10":["out","outubro"],"11":["nov","novembro"],"12":["dez","dezembro"]}
        for c in cols:
            cn = str(c).lower().strip()
            if any(t in cn for t in tokens): return c
        for alias in months_pt.get(mm, []):
            for c in cols:
                cn = str(c).lower().strip()
                if alias in cn: return c
        num_cols = [c for c in cols if pd.api.types.is_numeric_dtype(dfb[c])]
        return num_cols[-1] if num_cols else None
    month_col = _month_col(list(dfb.columns), competencia)
    sind_col = next((c for c in dfb.columns if "sind" in c), None)
    uf_col = "uf" if "uf" in dfb.columns else next((c for c in dfb.columns if c in ("estado","sigla")), None)
    if month_col is None:
        return pd.Series(fallback_days, index=uf_series.index)
    if sindicato_series is not None and sind_col:
        tmp = dfb[[sind_col, month_col]].dropna(subset=[sind_col])
        tmp[sind_col] = tmp[sind_col].astype(str).str.strip().str.lower()
        days_map_sind = tmp.dropna(subset=[month_col]).groupby(sind_col)[month_col].last().to_dict()
        series_norm = sindicato_series.astype(str).str.strip().str.lower()
        mapped = series_norm.map(lambda s: days_map_sind.get(s, pd.NA))
    else:
        mapped = pd.Series(pd.NA, index=uf_series.index)
    if uf_col:
        tmp2 = dfb[[uf_col, month_col]].dropna(subset=[uf_col])
        tmp2[uf_col] = tmp2[uf_col].astype(str).str.upper().str.strip()
        days_map_uf = tmp2.dropna(subset=[month_col]).groupby(uf_col)[month_col].last().to_dict()
        by_uf = uf_series.map(lambda u: days_map_uf.get(str(u).upper(), pd.NA))
    else:
        by_uf = pd.Series(pd.NA, index=uf_series.index)
    out = pd.Series(fallback_days, index=uf_series.index, dtype="float")
    out = out.where(mapped.isna() & by_uf.isna(), mapped.fillna(by_uf))
    return out
def business_days(a: pd.Timestamp, b: pd.Timestamp) -> int:
    if pd.isna(a) or pd.isna(b) or b <= a: return 0
    return int(np.busday_count(np.datetime64(a.date()), np.datetime64(b.date())))

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
        afast["inicio"] = afast["inicio"].map(pd.to_datetime); afast["fim"] = afast["fim"].map(pd.to_datetime)
        afast = afast.dropna(subset=["inicio","fim"])
        def overlap(i,f): return not (f <= dt_inicio or i >= dt_fim)
        afast_set = set(afast.loc[[overlap(i,f) for i,f in zip(afast["inicio"], afast["fim"])], "matricula"].astype(str))

    ativos["matricula_str"] = ativos["matricula"].astype(str)
    excl_mask = is_dir | ativos["matricula_str"].isin(est_set | apr_set | ext_set | afast_set)
    base = ativos.loc[~excl_mask].copy()

    # drop who demitted before period
    if desligados is not None and not desligados.empty:
        dmap = desligados.set_index("matricula")["data_demissao"].to_dict()
        dem_date = base["matricula_str"].map(dmap)
        base = base.loc[~(dem_date.notna() & (dem_date < dt_inicio))].copy()

    fallback_days = int(np.busday_count(np.datetime64(dt_inicio.date()), np.datetime64(dt_fim.date())))
    dias_series = best_days_series(base_dias, competencia, base["uf"], fallback_days, sindicato_series=base["sindicato"]).astype(float)
    obs = pd.Series(["Trabalhando"] * len(base), index=base.index, dtype=object)

    if ferias is not None and not ferias.empty:
        ferias["inicio"] = ferias["inicio"].map(pd.to_datetime); ferias["fim"] = ferias["fim"].map(pd.to_datetime)
        ferias = ferias.dropna(subset=["inicio","fim"])
        def ov_bd(i,f):
            i2 = max(i, dt_inicio); f2 = min(f, dt_fim)
            if pd.isna(i2) or pd.isna(f2) or i2>=f2: return 0
            return int(np.busday_count(np.datetime64(i2.date()), np.datetime64(f2.date())))
        ferias["bdays"] = [ov_bd(i,f) for i,f in zip(ferias["inicio"], ferias["fim"])]
        f_map = ferias.groupby("matricula")["bdays"].sum().to_dict()
        sub = base["matricula_str"].map(lambda m: f_map.get(m, 0)).astype(float)
        dias_series = (dias_series - sub).clip(lower=0)
        has_f = sub > 0
        obs.loc[has_f] = obs.loc[has_f] + " | Férias (" + sub[has_f].astype(int).astype(str) + " dias úteis no período)"

    # Admissão dentro do período
    adm = base["admissao"]
    mask_adm = adm.notna() & (adm > dt_inicio) & (adm < dt_fim)
    if mask_adm.any():
        red = [int(np.busday_count(np.datetime64(dt_inicio.date()), np.datetime64(a.date()))) if m else 0 for a, m in zip(adm, mask_adm)]
        red = pd.Series(red, index=base.index).astype(float)
        dias_series = (dias_series - red).clip(lower=0)
        obs.loc[mask_adm] = obs.loc[mask_adm] + " | Admissão em " + adm[mask_adm].dt.strftime("%d/%m/%Y")

    # Demissão regra 15 + comunicado OK
    if desligados is not None and not desligados.empty:
        dmap = desligados.set_index("matricula")[["data_demissao","ok_comunicado","data_comunicado"]].to_dict(orient="index")
        cutoff = pd.Timestamp(year=dt_fim.year, month=dt_fim.month, day=15)
        new_days = []
        new_obs = obs.copy()
        for idx, (mat, days) in enumerate(zip(base["matricula_str"], dias_series)):
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
                used = int(np.busday_count(np.datetime64(dt_inicio.date()), np.datetime64(dd.date())))
                new_days.append(min(int(days), used))
                new_obs.iloc[idx] = (str(new_obs.iloc[idx]) + f" | Desligado em {dd.strftime('%d/%m/%Y')} (proporcional)").strip(" |")
        dias_series = pd.Series(new_days, index=base.index).astype(int)
        obs = new_obs

    # Valor diário
    def daily_value_for_row(uf, sindicato):
        dfv = sind_valor
        if dfv is not None and not dfv.empty:
            dfv2 = dfv.copy()
            dfv2.columns = [strip_accents(str(c)).lower().replace(" ", "_") for c in dfv2.columns]
            val_col = "valor" if "valor" in dfv2.columns else next((c for c in dfv2.columns if "valor" in c), None)
            if val_col:
                sind_cand = next((c for c in dfv2.columns if "sind" in c), None)
                if sind_cand and isinstance(sindicato, str) and sindicato.strip():
                    temp = dfv2[dfv2[sind_cand].astype(str).str.strip().str.lower() == str(sindicato).strip().lower()]
                    if not temp.empty and pd.notna(temp.iloc[-1][val_col]):
                        return float(temp.iloc[-1][val_col])
                    temp2 = dfv2[dfv2[sind_cand].astype(str).str.lower().str.contains(str(sindicato).strip().lower(), na=False)]
                    if not temp2.empty and pd.notna(temp2.iloc[-1][val_col]):
                        return float(temp2.iloc[-1][val_col])
                uf_cand = "uf" if "uf" in dfv2.columns else next((c for c in dfv2.columns if c in ("estado","sigla")), None)
                if uf_cand:
                    m = dfv2[dfv2[uf_cand].astype(str).str.upper().str.strip() == str(uf).upper().strip()]
                    if not m.empty and pd.notna(m.iloc[-1][val_col]):
                        return float(m.iloc[-1][val_col])
        return FALLBACK_DAILY_BY_UF.get(str(uf).upper(), np.nan)


    daily_vals = pd.Series([daily_value_for_row(uf, sind) for uf, sind in zip(base["uf"], base["sindicato"])], index=base.index).astype(float)
    # === EXTERIOR override (substitui valor calculado quando houver valor em EXTERIOR.xlsx) ===
    ext_path = data_dir / "EXTERIOR.xlsx"
    if ext_path.exists():
        try:
            ext = pd.read_excel(ext_path, engine="openpyxl")
            ext.columns = [strip_accents(str(c)).lower().strip().replace(" ", "_") for c in ext.columns]
            mat_col = next((c for c in ext.columns if "matric" in c or c=="id" or "registro" in c), None)
            val_cols = [c for c in ext.columns if any(k in c for k in ["valor", "total", "vr", "beneficio", "benefício"])]
            if mat_col and val_cols:
                ext2 = ext[[mat_col] + val_cols].copy()
                ext2.columns = ["matricula"] + val_cols
                ext2["matricula"] = ext2["matricula"].astype(str).str.extract(r"(\d+)")[0].fillna(ext2["matricula"].astype(str))
                ext_map = ext2.groupby("matricula").last()
                for idx, row in base.iterrows():
                    mid = str(row["matricula"])
                    if mid in ext_map.index:
                        vals = ext_map.loc[mid]
                        ext_total = next((c for c in val_cols if "total" in c), None)
                        ext_valor = next((c for c in val_cols if "valor" in c), None)
                        if ext_total and pd.notna(vals.get(ext_total, pd.NA)):
                            obs.iloc[idx] = (str(obs.iloc[idx]) + "; " if str(obs.iloc[idx]) not in ["", "nan"] else "") + "Exterior (valor total)"
                            base.loc[idx, "_ext_total_override"] = float(vals[ext_total])
                        elif ext_valor and pd.notna(vals.get(ext_valor, pd.NA)):
                            daily_vals.iloc[idx] = float(vals[ext_valor])
                            obs.iloc[idx] = (str(obs.iloc[idx]) + "; " if str(obs.iloc[idx]) not in ["", "nan"] else "") + "Exterior (valor diário)"
        except Exception:
            pass


    total = (pd.Series(dias_series, index=base.index).fillna(0).astype(int) * daily_vals.fillna(0)).round(2)
    if "_ext_total_override" in base.columns:
        override = pd.to_numeric(base["_ext_total_override"], errors="coerce")
        total = total.where(override.isna(), override)
    custo_empresa = (total * 0.80).round(2); desconto_prof = (total * 0.20).round(2)

    out = pd.DataFrame({
        "Matricula": base["matricula"].astype(str),
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

    out_path = Path(args.out_dir) / f"VR_MENSAL_{comp_num}.csv"
    out.to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"Gerado: {out_path}")

if __name__ == "__main__":
    main()