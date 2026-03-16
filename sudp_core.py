from pathlib import Path
import sys
import pandas as pd
import numpy as np
import calendar
import datetime as dt
import functools
import os, json, time, zipfile, shutil

# ===== 기본값 =====
DEFAULT_ROOT = r"C:\Program Files (x86)\MasterSpace\M-Core\SUDP"
WEEKDAY_KO = {0:"월",1:"화",2:"수",3:"목",4:"금",5:"토",6:"일"}

RESULT_FILENAMES = {
    "입찰량":        "CD_RESULT_AVAIL.txt",
    "발전량":        "CD_POWER_GROSS.txt",
    "송전량":        "CD_POWER.txt",
    "연료사용량":     "CD_RESULT_CAL.txt",
    "발전비용":      "CD_RESULT_COST.txt",
    # 정산금
    "MEP":          "CBP_SEP2.txt",
    "MAP":          "CBP_COFF2.txt",
    "MWP":          "CBP_CON2.txt",
    # SMP
    "SMP":          "UD_SMP.txt",
    "SMPUNIT":      "UD_SMPUnit.txt",
}

REGION_ALIAS_TO_STD = {
    "수도권": "경인", "비수도권": "비경인",
    "경인": "경인", "비경인": "비경인", "제주": "제주",
    "JEJU": "제주", "GI": "경인", "BGI": "비경인"
}

import os, json, time, zipfile, shutil
from pathlib import Path

# 우리가 실제로 사용하는 파일만 스냅샷 (불필요한 대용량 파일 제외)
NEEDED_RESULT_FILES = {
    "CD_RESULT_AVAIL.txt",
    "CD_POWER_GROSS.txt",
    "CD_POWER.txt",
    "CD_RESULT_COST.txt",
    "CD_RESULT_CAL.txt",
    "CBP_SEP2.txt",
    "CBP_CON2.txt",
    "CBP_COFF2.txt",
    "UD_SMP.txt",
    "UD_SMPUnit.txt",
    "CD_RESULT_STUP.txt",     # 기동여부
}
NEEDED_DATA_FILES = {"CD_DATA.txt"}
NEEDED_OPTION_FILES = {"UD_LOAD.txt"}

def snapshot_sudp(root: Path, ys:int, ye:int,
                  out_root: Path, mode:str="zip",
                  include_catalog:bool=True,
                  snap_name: str | None = None) -> Path:
    """
    선택 연-월 범위의 '필요한 텍스트 파일'만 스냅샷.
    mode: "zip" | "copy"
    snap_name: 사용자 지정 이름(빈칸/None이면 자동 생성). zip 모드에서는 .zip 자동 부여.
    반환: 생성된 스냅샷 경로 (zip 파일 경로 또는 폴더 경로)
    """
    import re, time, zipfile, shutil, json

    # 우리가 실제로 사용하는 파일만 스냅샷
    NEEDED_RESULT_FILES = {
        "CD_RESULT_AVAIL.txt",
        "CD_POWER_GROSS.txt",
        "CD_POWER.txt",
        "CD_RESULT_COST.txt",
        "CD_RESULT_CAL.txt",
        "CBP_SEP2.txt",
        "CBP_CON2.txt",
        "CBP_COFF2.txt",
        "UD_SMP.txt",
        "UD_SMPUnit.txt",
        "CD_RESULT_STUP.txt",
    }
    NEEDED_DATA_FILES = {"CD_DATA.txt"}
    NEEDED_OPTION_FILES = {"UD_LOAD.txt"}

    def safe_name(s: str) -> str:
        # Windows 금지 문자 제거
        s = re.sub(r'[\\/:*?"<>|]+', "_", s)
        return s.strip(" .")

    ts = time.strftime("%Y%m%d_%H%M%S")
    out_root = Path(out_root)
    out_root.mkdir(parents=True, exist_ok=True)

    # 파일/폴더 기본 이름
    if snap_name and snap_name.strip():
        base = safe_name(snap_name.strip())
    else:
        base = f"SUDP_snapshot_{ys}_{ye}_{ts}"

    if mode not in ("zip", "copy"):
        mode = "zip"

    if mode == "zip":
        if not base.lower().endswith(".zip"):
            base += ".zip"
        snap_path = out_root / base
        zf = zipfile.ZipFile(snap_path, "w", compression=zipfile.ZIP_DEFLATED)

        def _add_file(src: Path, arc_rel: str):
            zf.write(src, arcname=arc_rel)
    else:
        snap_path = out_root / base
        snap_path.mkdir(parents=True, exist_ok=True)

        def _add_file(src: Path, arc_rel: str):
            dst = snap_path / arc_rel
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)

    # ✅ 누락됐던 초기화
    added: list[str] = []

    # 파일 수집
    for y in range(ys, ye + 1):
        for m in range(1, 13):
            ym_rel = Path(str(y)) / f"{m:02d}"

            # CCDATA / THDATA
            for sub in ("CCDATA", "THDATA"):
                for fn in NEEDED_DATA_FILES:
                    src = root / str(y) / f"{m:02d}" / sub / fn
                    if src.exists():
                        arc = (ym_rel / sub / fn).as_posix()
                        _add_file(src, arc); added.append(arc)

            # RESULT
            for fn in NEEDED_RESULT_FILES:
                src = root / str(y) / f"{m:02d}" / "RESULT" / fn
                if src.exists():
                    arc = (ym_rel / "RESULT" / fn).as_posix()
                    _add_file(src, arc); added.append(arc)

            # OPTION (수요)
            for fn in NEEDED_OPTION_FILES:
                src = root / str(y) / f"{m:02d}" / "OPTION" / fn
                if src.exists():
                    arc = (ym_rel / "OPTION" / fn).as_posix()
                    _add_file(src, arc); added.append(arc)

    # META.json
    meta = {
        "snapshot_name": snap_path.name,
        "created_at": ts,
        "root": str(root),
        "years": [ys, ye],
        "mode": mode,
        "files": added,
        "program": "sudp_core.py",
        "version": "snapshot-1",
        "user_base_name": snap_name or "",
    }
    meta_bytes = json.dumps(meta, ensure_ascii=False, indent=2).encode("utf-8")
    if mode == "zip":
        zf.writestr("META.json", meta_bytes)
        zf.close()
    else:
        (snap_path / "META.json").write_bytes(meta_bytes)

    return snap_path

# sudp_core.py (상단 import 근처에 추가)
from pathlib import Path
import shutil, zipfile, datetime

def _snapshot_sudp(root: Path, ys: int, ye: int,
                   mode: str = "zip",
                   out_dir: Path | None = None,
                   name: str | None = None) -> str:
    """
    SUDP 폴더 전체를 zip 또는 copy로 스냅샷.
    반환: 만들어진 zip 파일 경로 또는 복사된 폴더 경로 (문자열)
    """
    mode = (mode or "zip").lower()
    out_dir = Path(out_dir) if out_dir else Path.cwd()
    out_dir.mkdir(parents=True, exist_ok=True)

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    base = (name or f"SUDP_snapshot_{ys}_{ye}_{ts}").strip()

    src = Path(root)
    if mode == "copy":
        dst = out_dir / base
        if dst.exists():
            shutil.rmtree(dst)
        shutil.copytree(src, dst)
        return str(dst)
    else:
        dst = out_dir / f"{base}.zip"
        with zipfile.ZipFile(dst, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for p in src.rglob("*"):
                zf.write(p, p.relative_to(src))
        return str(dst)


# ===== 공통 유틸 =====
def read_table_flexible(path: Path):
    if not path.exists():
        return None
    for enc in ("utf-8-sig", "cp949", "euc-kr"):
        for sep in ("\t", ",", None):
            try:
                df = pd.read_csv(path, sep=sep, encoding=enc, engine="python")
                if df.shape[1] >= 2:
                    return df
            except Exception:
                continue
    return None

def month_slots_to_calendar(year:int, month:int, nslots:int):
    days = calendar.monthrange(year, month)[1]
    n = min(nslots, days*24)
    idx = np.arange(1, n+1)
    day = ((idx-1)//24) + 1
    hour = ((idx-1)%24) + 1
    weekday = pd.to_datetime({"year":[year]*n,"month":[month]*n,"day":day}).dt.weekday.map(WEEKDAY_KO).values
    return np.array([year]*n), np.array([month]*n), day, weekday, hour

def expand_month_matrix(df: pd.DataFrame, year:int, month:int):
    """첫 열=키(자원ID 또는 지역), 나머지 열=슬롯(1..744) → long화 + 시간복원"""
    if df is None or df.empty:
        return None
    cols = list(df.columns)
    key_col = cols[0]
    slot_cols = [c for c in cols[1:] if str(c).strip().isdigit() or str(c).replace(".","",1).isdigit()]
    if not slot_cols:
        return None
    slot_cols = sorted(slot_cols, key=lambda x: int(float(str(x))))
    nslots = len(slot_cols)
    m = df.melt(id_vars=[key_col], value_vars=slot_cols, var_name="슬롯", value_name="value")
    m["슬롯"] = m["슬롯"].astype(float).astype(int)
    y,mn,d,w,h = month_slots_to_calendar(year, month, nslots)
    time_tbl = pd.DataFrame({"슬롯": np.arange(1, len(y)+1), "연도":y, "월":mn, "일":d, "요일":w, "시간":h})
    out = m.merge(time_tbl, on="슬롯", how="left").rename(columns={key_col:"키"})
    return out

def safe_concat(frames):
    frames = [f for f in frames if f is not None and not f.empty]
    if not frames: return pd.DataFrame()
    out = pd.concat(frames, ignore_index=True)
    if all(c in out.columns for c in ["연도","월","일","시간"]):
        out = out.sort_values(["연도","월","일","시간"])
    return out


def iter_year_months(ys: int, ms: int, ye: int, me: int):
    """시작(연,월)~종료(연,월) 범위의 (연,월) 튜플을 순회한다."""
    ys, ye = int(ys), int(ye)
    ms, me = int(ms), int(me)
    if (ys, ms) > (ye, me):
        ys, ye = ye, ys
        ms, me = me, ms

    for y in range(ys, ye + 1):
        m_start = ms if y == ys else 1
        m_end = me if y == ye else 12
        for m in range(m_start, m_end + 1):
            yield y, m

# ===== 카탈로그 스캔 =====
def scan_catalog_all(root: Path, ys:int, ms:int, ye:int, me:int):
    """
    선택 기간의 모든 연/월/CCDATA·THDATA의 CD_DATA.txt를 끝까지 훑어
    - 자원코드/자원명/자원ID/설비용량/송전단최대용량 누적(최신값 우선)
    - 보조 맵: code->id, id->name, id->capacity
    반환: catalog_df, code_to_id, id_to_name, id_to_capacity_x2
    """
    records = []
    for y, m in iter_year_months(ys, ms, ye, me):
        for sub in ["CCDATA","THDATA"]:
            p = root/str(y)/f"{m:02d}"/sub/"CD_DATA.txt"
            df = read_table_flexible(p)
            if df is None:
                continue
            cols = df.columns
            c_code = "자원코드" if "자원코드" in cols else None
            c_name = "자원명"   if "자원명"   in cols else None
            c_id   = "자원ID"   if "자원ID"   in cols else None
            c_cap  = "설비용량" if "설비용량" in cols else None
            c_tx   = "송전단최대용량" if "송전단최대용량" in cols else None
            c_ru   = "Ramp-up Rate"  if "Ramp-up Rate"  in cols else None
            c_hot  = "HOT기동비용"    if "HOT기동비용"    in cols else None
            if not (c_code and c_name and c_id):
                continue

            for _, r in df.iterrows():
                try:
                    rid = int(float(str(r[c_id])))
                except Exception:
                    continue
                rec = {
                    "스냅연": y, "스냅월": m,
                    "자원코드": str(r[c_code]).strip(),
                    "자원명":   str(r[c_name]).strip(),
                    "자원ID":   rid,
                    "설비용량": float(r[c_cap]) if c_cap in df.columns else np.nan,
                    "송전단최대용량": float(r[c_tx]) if c_tx in df.columns else np.nan,
                    "Ramp-up Rate": float(r[c_ru]) if c_ru in df.columns and pd.notna(r[c_ru]) else np.nan,
                    "HOT기동비용":  float(r[c_hot]) if c_hot in df.columns and pd.notna(r[c_hot]) else np.nan,
                }
                records.append(rec)

    if not records:
        return pd.DataFrame(), {}, {}, {}, {}, {}

    cat = pd.DataFrame(records)
    cat = cat.sort_values(["스냅연","스냅월"]).drop_duplicates(subset=["자원ID"], keep="last")
    code_to_id = dict(zip(cat["자원코드"], cat["자원ID"]))
    id_to_name = dict(zip(cat["자원ID"], cat["자원명"]))
    id_to_capacity_x2 = {rid: (cap*2.0 if pd.notna(cap) else np.nan)
                         for rid, cap in zip(cat["자원ID"], cat["설비용량"])}
    id_to_ru  = {rid: (ru  if pd.notna(ru)  else np.nan) for rid, ru  in zip(cat["자원ID"], cat.get("Ramp-up Rate", np.nan))}
    id_to_hot = {rid: (hot if pd.notna(hot) else np.nan) for rid, hot in zip(cat["자원ID"], cat.get("HOT기동비용",  np.nan))}

    return cat, code_to_id, id_to_name, id_to_capacity_x2, id_to_ru, id_to_hot

def lookup_name_monthly(root: Path, y:int, m:int, rid:int, _cache:dict):
    key = (y,m,rid)
    if key in _cache: return _cache[key]
    for sub in ["CCDATA","THDATA"]:
        p = root/str(y)/f"{m:02d}"/sub/"CD_DATA.txt"
        df = read_table_flexible(p)
        if df is None or "자원ID" not in df.columns or "자원명" not in df.columns:
            continue
        rid_series = df["자원ID"].astype(str).str.replace(".0","", regex=False)
        hit = df.loc[rid_series == str(rid)]
        if not hit.empty:
            nm = str(hit.iloc[0]["자원명"])
            _cache[key] = nm
            return nm
    _cache[key] = None
    return None

# ===== 결과표 생성 (가변 개수 발전기) =====
def pick_units_wide(df_long: pd.DataFrame, target_ids: list, id_to_name: dict):
    """(키=자원ID, value=값) → [연/월/일/요일/시간, 각 발전기명 열...]"""
    if df_long is None or df_long.empty:
        return None
    tmp = df_long[df_long["키"].astype(str).str.replace(".0","", regex=False).str.isdigit()].copy()
    if tmp.empty:
        return pd.DataFrame()
    tmp["자원ID"] = tmp["키"].astype(float).astype(int)
    tmp = tmp[tmp["자원ID"].isin(target_ids)]
    if tmp.empty:
        return pd.DataFrame()
    wide = tmp.pivot_table(index=["연도","월","일","요일","시간"], columns="자원ID", values="value", aggfunc="first").reset_index()
    # 열 이름을 자원명으로 변환(동명이 있으면 "이름 (ID)"로 유니크)
    col_map = {}
    used_names = set()
    for rid in [c for c in wide.columns if isinstance(c, (int, np.integer))]:
        base = id_to_name.get(rid, str(rid))
        name = base
        if name in used_names:
            name = f"{base} ({rid})"
        used_names.add(name)
        col_map[rid] = name
    wide = wide.rename(columns=col_map)
    # 빠진 타깃은 NaN 열 생성
    for rid in target_ids:
        nm = col_map.get(rid, id_to_name.get(rid, str(rid)))
        if nm not in wide.columns: wide[nm] = np.nan
    fixed = ["연도","월","일","요일","시간"] + [nm for rid, nm in sorted(col_map.items(), key=lambda x:x[0])]
    return wide[fixed]

def build_basic(root:Path, ys:int, ms:int, ye:int, me:int, key:str, target_ids:list, id_to_name:dict):
    fname = RESULT_FILENAMES[key]
    frames=[]
    for y, m in iter_year_months(ys, ms, ye, me):
        df = read_table_flexible(root/str(y)/f"{m:02d}"/"RESULT"/fname)
        if df is None:
            continue
        long = expand_month_matrix(df, y, m)
        wide = pick_units_wide(long, target_ids, id_to_name)
        frames.append(wide)
    out = safe_concat(frames)
    if not out.empty:
        for c in ["연도","월","일","시간"]:
            out[c] = out[c].astype(int)
    return out

def build_utilization(root:Path, ys:int, ms:int, ye:int, me:int, target_ids:list, id_to_name:dict, id_to_capacity_x2:dict):
    gross = build_basic(root, ys, ms, ye, me, "발전량", target_ids, id_to_name)  # CD_POWER_GROSS
    if gross.empty: return gross
    # 각 발전기 열에 대해 분모(설비용량×2) 적용
    for rid in target_ids:
        nm = id_to_name.get(rid, str(rid))
        # 이름이 중복되어 pick_units_wide에서 "이름 (ID)"가 되었을 수 있음 → 실제 열명 찾기
        col_candidates = [c for c in gross.columns if c == nm or c.endswith(f"({rid})")]
        if not col_candidates: 
            continue
        denom = id_to_capacity_x2.get(rid, np.nan)
        for col in col_candidates:
            gross[col] = (gross[col] / denom) * 100.0
    return gross

def build_settlement(root:Path, ys:int, ms:int, ye:int, me:int, target_ids:list, id_to_name:dict):
    # 1) 원천 표 생성
    mep = build_basic(root, ys, ms, ye, me, "MEP", target_ids, id_to_name)
    map_ = build_basic(root, ys, ms, ye, me, "MAP", target_ids, id_to_name)
    mwp = build_basic(root, ys, ms, ye, me, "MWP", target_ids, id_to_name)

    on = ["연도","월","일","요일","시간"]

    # 2) 발전기 열만 골라 접미사 부여 (키 열은 그대로 유지)
    def _suffix_values(df, suf):
        if df is None or df.empty:
            return pd.DataFrame(columns=on)
        gen_cols = [c for c in df.columns if c not in on]
        renamed = df[on + gen_cols].copy()
        renamed = renamed.rename(columns={c: f"{c}_{suf}" for c in gen_cols})
        return renamed

    mep_s = _suffix_values(mep, "MEP")
    map_s = _suffix_values(map_, "MAP")
    mwp_s = _suffix_values(mwp, "MWP")

    # 3) 키 기준 외부조인
    df = mep_s.merge(map_s, on=on, how="outer")
    df = df.merge(mwp_s, on=on, how="outer")

    # 4) 숫자 정규화 유틸
    def _to_numeric_col(s):
        s = s.astype(str)\
             .str.replace(",", "", regex=False)\
             .str.replace("\u00A0", "", regex=False)\
             .str.strip()
        s = s.replace({"-": np.nan, "": np.nan})
        return pd.to_numeric(s, errors="coerce")

    # 5) “발전기 표시명” 목록 수집 (접미사 떼고 고유화)
    value_cols = [c for c in df.columns if c not in on]
    base_names = []
    for c in value_cols:
        if c.endswith("_MEP") or c.endswith("_MAP") or c.endswith("_MWP"):
            base_names.append(c.rsplit("_", 1)[0])
    base_names = sorted(set(base_names))  # 예: ["동두천복합2CC", "안산복합CC", "위례열병합CC"]

    # 6) 출력 프레임 구성
    out = df[on].copy()
    for base in base_names:
        mep_col = base + "_MEP" if base + "_MEP" in df.columns else None
        map_col = base + "_MAP" if base + "_MAP" in df.columns else None
        mwp_col = base + "_MWP" if base + "_MWP" in df.columns else None

        mep_series = _to_numeric_col(df[mep_col]) if mep_col else pd.Series(np.nan, index=df.index)
        map_series = _to_numeric_col(df[map_col]) if map_col else pd.Series(np.nan, index=df.index)
        mwp_series = _to_numeric_col(df[mwp_col]) if mwp_col else pd.Series(np.nan, index=df.index)

        pretty = base  # 이미 "자원명" 또는 "자원명 (ID)" 형태
        out[f"MEP({pretty})"] = mep_series
        out[f"MAP({pretty})"] = map_series
        out[f"MWP({pretty})"] = mwp_series
        out[f"합계({pretty})"] = mep_series.fillna(0) + map_series.fillna(0) + mwp_series.fillna(0)

    # 7) 열 순서 정리
    tails = []
    for base in base_names:
        tails += [f"MEP({base})", f"MAP({base})", f"MWP({base})", f"합계({base})"]
    out = out[on + tails].sort_values(["연도","월","일","시간"])

    return out


def build_smp_hourly(root:Path, ys:int, ms:int, ye:int, me:int, id_to_name:dict):
    # 1) SMP 값
    smp_frames=[]; unit_frames=[]
    for y, m in iter_year_months(ys, ms, ye, me):
        ds = read_table_flexible(root/str(y)/f"{m:02d}"/"RESULT"/RESULT_FILENAMES["SMP"])
        du = read_table_flexible(root/str(y)/f"{m:02d}"/"RESULT"/RESULT_FILENAMES["SMPUNIT"])
        if ds is not None:
            smp_frames.append(expand_month_matrix(ds, y, m))
        if du is not None:
            unit_frames.append(expand_month_matrix(du, y, m))
    smp  = safe_concat(smp_frames)
    unit = safe_concat(unit_frames)
    if smp.empty:
        return pd.DataFrame(columns=["연도","월","일","요일","시간","경인","경인 결정 발전기","비경인","비경인 결정 발전기","제주","제주 결정 발전기"])

    # 2) SMP 값 pivot (지역만)
    smp = smp[smp["키"].astype(str).str.isdigit()==False].copy()
    smp_pv = smp.pivot_table(index=["연도","월","일","요일","시간","슬롯"], columns="키", values="value", aggfunc="first").reset_index()
    smp_pv = smp_pv.rename(columns={c: REGION_ALIAS_TO_STD.get(str(c), c) for c in smp_pv.columns})
    for need in ["경인","비경인","제주"]:
        if need not in smp_pv.columns: smp_pv[need] = np.nan

    # 3) 결정 발전기(자원ID→자원명)
    if not unit.empty:
        unit = unit.pivot_table(index=["연도","월","일","요일","시간","슬롯"], columns="키", values="value", aggfunc="first").reset_index()
        unit = unit.rename(columns={c: (REGION_ALIAS_TO_STD.get(str(c), None)+" 결정 발전기") if REGION_ALIAS_TO_STD.get(str(c), None) else c for c in unit.columns})
        # 숫자면 전역 맵으로 먼저 변환
        for nm in ["경인 결정 발전기","비경인 결정 발전기","제주 결정 발전기"]:
            if nm in unit.columns:
                unit[nm] = unit[nm].apply(lambda x: id_to_name.get(int(float(str(x))), str(x))
                                          if str(x).replace(".","",1).isdigit() else x)
        # 남으면 월별 재조회
        _cache={}
        for nm in ["경인 결정 발전기","비경인 결정 발전기","제주 결정 발전기"]:
            if nm in unit.columns:
                def fix(row):
                    v = row[nm]
                    s = str(v).strip()
                    if s.replace(".","",1).isdigit():
                        rid = int(float(s))
                        name = lookup_name_monthly(root, int(row["연도"]), int(row["월"]), rid, _cache)
                        return name if name else s
                    return v
                unit[nm] = unit.apply(fix, axis=1)

        smp_pv = smp_pv.merge(unit[["연도","월","일","요일","시간","슬롯"]+[c for c in ["경인 결정 발전기","비경인 결정 발전기","제주 결정 발전기"] if c in unit.columns]],
                              on=["연도","월","일","요일","시간","슬롯"], how="left")
    for nm in ["경인 결정 발전기","비경인 결정 발전기","제주 결정 발전기"]:
        if nm not in smp_pv.columns: smp_pv[nm]=np.nan

    return smp_pv[["연도","월","일","요일","시간","경인","경인 결정 발전기","비경인","비경인 결정 발전기","제주","제주 결정 발전기"]].sort_values(["연도","월","일","시간"])

def read_load_long(root: Path, ys:int, ms:int, ye:int, me:int):
    """
    각 연/월의 OPTION/UD_LOAD.txt를 읽어서
    - LOAD: 총수요
    - JLD : 제주수요
    를 시간축으로 확장해 반환.
    """
    frames = []
    for y, m in iter_year_months(ys, ms, ye, me):
        p = root/str(y)/f"{m:02d}"/"OPTION"/"UD_LOAD.txt"
        df = read_table_flexible(p)
        if df is None or df.empty:
            continue
        # 첫 열: 지표명(LOAD/NLD/JLD), 나머지 열: 슬롯(1..744)
        long = expand_month_matrix(df, y, m)  # 키=지표명
        if long is None or long.empty:
            continue
        # 필요한 지표만 남긴다: 총수요(LOAD), 제주수요(JLD)
        need = long[ long["키"].astype(str).isin(["LOAD","JLD"]) ].copy()
        if need.empty:
            continue
        pv = need.pivot_table(
            index=["연도","월","일","요일","시간","슬롯"],
            columns="키", values="value", aggfunc="first"
        ).reset_index()
        # 결측 방어
        if "LOAD" not in pv.columns: pv["LOAD"] = np.nan
        if "JLD"  not in pv.columns: pv["JLD"]  = np.nan
        frames.append(pv)
    return safe_concat(frames)


def build_smp_yearly(root:Path, ys:int, ms:int, ye:int, me:int, id_to_name:dict):
    """
    연도별 요약표:
    연도 /
      경인: 최대값, 자원명, 최소값, 자원명, 평균값, 가중평균값 /
      비경인: ... /
      제주: ...
    * 가중평균: 경인/비경인은 총수요(LOAD), 제주는 제주수요(JLD)를 가중치로 사용
    """
    # 시간별 SMP + 결정발전기명
    hourly = build_smp_hourly(root, ys, ms, ye, me, id_to_name)
    if hourly is None or hourly.empty:
        cols = ["연도",
                "경인_최대","경인_자원명(최대)","경인_최소","경인_자원명(최소)","경인_평균","경인_가중평균",
                "비경인_최대","비경인_자원명(최대)","비경인_최소","비경인_자원명(최소)","비경인_평균","비경인_가중평균",
                "제주_최대","제주_자원명(최대)","제주_최소","제주_자원명(최소)","제주_평균","제주_가중평균"]
        return pd.DataFrame(columns=cols)

    # 수요
    load = read_load_long(root, ys, ms, ye, me)   # LOAD, JLD
    # 시간축 결합
    df = hourly.merge(load, on=["연도","월","일","요일","시간"], how="left")

    rows = []
    for y, g in df.groupby("연도"):
        rec = {"연도": int(y)}

        for region, col_name, dcol in [
            ("경인",   "경인",   "경인 결정 발전기"),
            ("비경인", "비경인", "비경인 결정 발전기"),
            ("제주",   "제주",   "제주 결정 발전기"),
        ]:
            if col_name not in g.columns:  # 지역 컬럼이 없으면 생략
                rec.update({
                    f"{region}_최대": np.nan, f"{region}_자원명(최대)": np.nan,
                    f"{region}_최소": np.nan, f"{region}_자원명(최소)": np.nan,
                    f"{region}_평균": np.nan, f"{region}_가중평균": np.nan
                })
                continue

            vals = pd.to_numeric(g[col_name], errors="coerce")
            # 최대/최소 시점
            idx_max = vals.idxmax()
            idx_min = vals.idxmin()
            vmax = float(vals.loc[idx_max]) if pd.notna(idx_max) else np.nan
            vmin = float(vals.loc[idx_min]) if pd.notna(idx_min) else np.nan
            name_max = g.loc[idx_max, dcol] if (dcol in g.columns and pd.notna(idx_max)) else np.nan
            name_min = g.loc[idx_min, dcol] if (dcol in g.columns and pd.notna(idx_min)) else np.nan

            # 평균
            vmean = float(vals.mean()) if len(vals.dropna()) else np.nan
            # 가중평균: 경인/비경인은 LOAD, 제주는 JLD
            if region == "제주":
                w = pd.to_numeric(g["JLD"], errors="coerce")
            else:
                w = pd.to_numeric(g["LOAD"], errors="coerce")
            wavg = float((vals * w).sum() / w.sum()) if (w.notna().any() and w.sum() and vals.notna().any()) else np.nan

            rec.update({
                f"{region}_최대": vmax, f"{region}_자원명(최대)": name_max,
                f"{region}_최소": vmin, f"{region}_자원명(최소)": name_min,
                f"{region}_평균": vmean, f"{region}_가중평균": wavg
            })

        rows.append(rec)

    cols = ["연도",
            "경인_최대","경인_자원명(최대)","경인_최소","경인_자원명(최소)","경인_평균","경인_가중평균",
            "비경인_최대","비경인_자원명(최대)","비경인_최소","비경인_자원명(최소)","비경인_평균","비경인_가중평균",
            "제주_최대","제주_자원명(최대)","제주_최소","제주_자원명(최소)","제주_평균","제주_가중평균"]
    return pd.DataFrame(rows, columns=cols).sort_values("연도")

def build_reserve_capacity_payment(root:Path, ys:int, ms:int, ye:int, me:int,
                                   target_ids:list, id_to_name:dict,
                                   id_to_ru:dict, unit_price:float):
    """예비력용량가치정산금: 시간별로 발전기별 금액 산출"""
    on = ["연도","월","일","요일","시간"]

    # 원천: 입찰량/송전량
    avail = build_basic(root, ys, ms, ye, me, "입찰량", target_ids, id_to_name)
    send  = build_basic(root, ys, ms, ye, me, "송전량", target_ids, id_to_name)

    # 키 기준 외부조인(시간축 정렬)
    base = avail[on].drop_duplicates() if avail is not None and not avail.empty else pd.DataFrame(columns=on)
    if send is not None and not send.empty:
        base = pd.concat([base, send[on]]).drop_duplicates()
    if base.empty:
        return pd.DataFrame(columns=on)

    base = base.sort_values(["연도","월","일","시간"]).reset_index(drop=True)

    # 계산
    out = base.copy()

    # 각 발전기 열 이름 결정(중복명 → "이름 (ID)"로 만들어진 경우까지 포함)
    def colnames(df, rid):
        nm = id_to_name.get(rid, str(rid))
        if df is None or df.empty:
            return []
        cols = []
        for c in df.columns:
            if c in on: continue
            if c == nm or c.endswith(f"({rid})"):
                cols.append(c)
        return cols

    for rid in target_ids:
        # 입찰/송전 열명 찾기
        a_cols = colnames(avail, rid)
        s_cols = colnames(send,  rid)
        # 표시명은 입찰에서 우선, 없으면 송전에서
        disp = (a_cols or s_cols or [id_to_name.get(rid, str(rid))])[0]

        # 시리즈 맞추기
        a = pd.Series(index=out.index, dtype="float64")
        s = pd.Series(index=out.index, dtype="float64")

        if a_cols:
            tmp = base.merge(avail[on + [a_cols[0]]], on=on, how="left")
            a = pd.to_numeric(tmp[a_cols[0]], errors="coerce")
        if s_cols:
            tmp = base.merge(send[on + [s_cols[0]]], on=on, how="left")
            s = pd.to_numeric(tmp[s_cols[0]], errors="coerce")

        # RU (MW/분) 확보
        ru = id_to_ru.get(rid, np.nan)
        if not pd.notna(ru):
            # RU가 없으면 계산 불가 → NaN
            out[disp] = np.nan
            continue

        # 계산식
        # 송전량==0 → 0
        # Δ = max(입찰-송전, 0)
        delta = (a - s)
        delta = delta.where(delta > 0, 0)  # 음수는 0

        tier1 = np.minimum(delta, ru * 5.0)
        tier2 = np.maximum(np.minimum(delta, ru * 10.0) - np.minimum(delta, ru * 5.0), 0.0)

        payment = (tier1 + tier2) * float(unit_price)
        payment = payment.where(s > 0, 0.0)  # 송전량==0 → 0

        out[disp] = payment

    return out

def read_startup_long(root:Path, ys:int, ms:int, ye:int, me:int, target_ids:list, id_to_name:dict):
    """RESULT/CD_RESULT_STUP.txt : 자원ID별 시간당 0/1 → long 형태로."""
    frames = []
    for y, m in iter_year_months(ys, ms, ye, me):
        p = root/str(y)/f"{m:02d}"/"RESULT"/"CD_RESULT_STUP.txt"
        df = read_table_flexible(p)
        if df is None or df.empty:
            continue
        long = expand_month_matrix(df, y, m)
        if long is None or long.empty:
            continue
        long["자원ID"] = pd.to_numeric(long["키"], errors="coerce").astype("Int64")
        long = long[ long["자원ID"].isin(pd.array(target_ids, dtype="Int64")) ].copy()
        if long.empty:
            continue
        long["자원명"] = long["자원ID"].map(id_to_name)
        long = long.rename(columns={"value":"STUP"})
        frames.append(long[["연도","월","일","요일","시간","자원ID","자원명","STUP"]])
    return safe_concat(frames)


def build_generation_cost(
    root: Path,
    ys: int, ms: int, ye: int, me: int,
    target_ids: list[int],
    id_to_name: dict[int, str],
    id_to_hot: dict[int, float],
) -> pd.DataFrame:
    """
    발전비용 시트 생성:
      - 총 발전비용: RESULT/CD_RESULT_COST.txt  (자원ID 행 × 시간슬롯 열)
      - 기동비용:   RESULT/CD_RESULT_STUP.txt (0/1) × HOT기동비용(원/회, CD_DATA)
      - 연료비용 = 총 발전비용 - 기동비용
    출력: [연/월/일/요일/시간] + [발전비용|자원명…] + [기동비용|자원명…] + [연료비용|자원명…]
    """
    on = ["연도","월","일","요일","시간"]

    # 월 행렬을 long으로 안전 변환 → 대상 자원만 wide 피벗
    # (프로젝트 공용 경로와 동일: expand_month_matrix → pick_units_wide)  :contentReference[oaicite:2]{index=2}
    def _collect_wide(fname: str) -> list[pd.DataFrame]:
        frames = []
        for y, m in iter_year_months(ys, ms, ye, me):
            df = read_table_flexible(root/str(y)/f"{m:02d}"/"RESULT"/fname)
            if df is None:
                continue
            long = expand_month_matrix(df, y, m)
            wide = pick_units_wide(long, target_ids, id_to_name)
            if wide is not None and not wide.empty:
                frames.append(wide)
        out = safe_concat(frames)
        if not out.empty:
            for c in ["연도","월","일","시간"]:
                out[c] = out[c].astype(int)
        return out

    # 1) 총 발전비용
    gen_cost = _collect_wide("CD_RESULT_COST.txt")  # 각 자원명 열 = 총 발전비용(원)

    # 2) 기동비용 = STUP(0/1) × HOT(원/회)
    #    STUP도 같은 방식으로 읽되, long 단계에서 HOT을 곱해 'value'를 비용으로 바꾼 후 wide 피벗.
    frames_stup_hot = []
    for y, m in iter_year_months(ys, ms, ye, me):
        df = read_table_flexible(root/str(y)/f"{m:02d}"/"RESULT"/"CD_RESULT_STUP.txt")
        if df is None:
            continue
        long = expand_month_matrix(df, y, m)
        if long is None or long.empty:
            continue
        tmp = long.copy()
        tmp = tmp[tmp["키"].astype(str).str.replace(".0","", regex=False).str.isdigit()]
        if tmp.empty:
            continue
        tmp["자원ID"] = tmp["키"].astype(float).astype(int)
        tmp = tmp[tmp["자원ID"].isin(target_ids)]
        if tmp.empty:
            continue
        tmp["HOT"] = tmp["자원ID"].map(id_to_hot).fillna(0.0)
        tmp["value"] = pd.to_numeric(tmp["value"], errors="coerce").fillna(0.0) * tmp["HOT"]
        wide = tmp.pivot_table(index=on, columns="자원ID", values="value", aggfunc="sum").reset_index()
        col_map, used = {}, set()
        for rid in [c for c in wide.columns if isinstance(c, (int, np.integer))]:
            base = id_to_name.get(int(rid), str(int(rid)))
            name = base if base not in used else f"{base} ({int(rid)})"
            used.add(name)
            col_map[rid] = name
        wide = wide.rename(columns=col_map)
        for rid in target_ids:
            nm = col_map.get(rid, id_to_name.get(rid, str(rid)))
            if nm not in wide.columns:
                wide[nm] = np.nan
        wide = wide[on + [c for c in wide.columns if c not in on]]
        frames_stup_hot.append(wide)

    stup_cost = safe_concat(frames_stup_hot)
    if not stup_cost.empty:
        for c in ["연도","월","일","시간"]:
            stup_cost[c] = stup_cost[c].astype(int)

    # 3) 키 기준 외부조인 후, “연료비용 = 발전비용 - 기동비용”
    if gen_cost is None or gen_cost.empty:
        base = stup_cost.copy()
    elif stup_cost is None or stup_cost.empty:
        base = gen_cost.copy()
    else:
        base = gen_cost.merge(stup_cost, on=on, how="outer", suffixes=("", "__STUP__"))

    if base is None or base.empty:
        return pd.DataFrame(columns=on)  # 데이터 없음

    # 자원명 목록(발전/기동 양쪽 합집합)
    gen_cols = [c for c in base.columns if c not in on and not c.endswith("__STUP__")]
    stp_cols = [c[:-8] for c in base.columns if c.endswith("__STUP__")]  # '__STUP__' 제거
    names = sorted(set(gen_cols) | set(stp_cols))

    # 결과 그리드 구성
    out = base[on].copy()
    for nm in names:
        gcol = nm
        scol = f"{nm}__STUP__"
        g = pd.to_numeric(base.get(gcol), errors="coerce")
        s = pd.to_numeric(base.get(scol), errors="coerce")
        out[f"발전비용|{nm}"] = g
        out[f"기동비용|{nm}"] = s
        out[f"연료비용|{nm}"] = g.fillna(0.0) - s.fillna(0.0)

    # 정렬 및 반환
    out = out.sort_values(on).reset_index(drop=True)
    return out

def build_full_hourly_index(ys:int, ms:int, ye:int, me:int) -> pd.DataFrame:
    """시작~종료연도의 모든 연/월/일/시간 행(1~24h)을 생성."""
    rows = []
    wk_names = ["월","화","수","목","금","토","일"]  # datetime.weekday(): 월=0
    for y, m in iter_year_months(ys, ms, ye, me):
        ndays = calendar.monthrange(y, m)[1]
        for d in range(1, ndays+1):
            w = wk_names[dt.date(y, m, d).weekday()]
            for h in range(1, 25):
                rows.append((y, m, d, w, h))
    return pd.DataFrame(rows, columns=["연도","월","일","요일","시간"])


def aggregate_to_yearly(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    """시간별 결과를 연도별로 집계한다(이용률=평균, 나머지=합계)."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["연도"])

    if "연도" not in df.columns:
        return df.copy()

    key_cols = [c for c in ["연도", "월", "일", "요일", "시간"] if c in df.columns]
    value_cols = [c for c in df.columns if c not in key_cols]
    if not value_cols:
        return df[["연도"]].drop_duplicates().sort_values("연도").reset_index(drop=True)

    agg_func = "mean" if sheet_name == "이용률" else "sum"
    tmp = df.copy()
    tmp[value_cols] = tmp[value_cols].apply(lambda s: pd.to_numeric(s.astype(str).str.replace(",", "", regex=False), errors="coerce"))

    out = (
        tmp.groupby("연도", as_index=False)[value_cols]
        .agg(agg_func)
        .sort_values("연도")
        .reset_index(drop=True)
    )
    return out


def aggregate_fuel_to_yearly(df: pd.DataFrame) -> pd.DataFrame:
    """연료사용량 전용 연도 집계(시간별 합계). 문자열 숫자(콤마)도 안전 변환."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["연도"])
    if "연도" not in df.columns:
        return df.copy()

    key_cols = [c for c in ["연도", "월", "일", "요일", "시간"] if c in df.columns]
    val_cols = [c for c in df.columns if c not in key_cols]
    if not val_cols:
        return df[["연도"]].drop_duplicates().sort_values("연도").reset_index(drop=True)

    tmp = df.copy()
    for c in val_cols:
        tmp[c] = pd.to_numeric(tmp[c].astype(str).str.replace(",", "", regex=False), errors="coerce")

    out = (
        tmp.groupby("연도", as_index=False)[val_cols]
        .sum(min_count=1)
        .sort_values("연도")
        .reset_index(drop=True)
    )
    return out


def build_fuel_sheet_with_ton(df: pd.DataFrame, applied_hhv: float) -> pd.DataFrame:
    """연료사용량 시트: [열량(Mcal)] + 빈 열 + [연료량(ton)] 2개 표를 병렬 생성."""
    if df is None:
        return pd.DataFrame()

    base = df.copy()
    key_cols = [c for c in ["연도", "월", "일", "요일", "시간"] if c in base.columns]
    val_cols = [c for c in base.columns if c not in key_cols]

    if val_cols:
        base[val_cols] = base[val_cols].apply(lambda s: pd.to_numeric(s.astype(str).str.replace(",", "", regex=False), errors="coerce")) * 1000.0

    ton_tbl = base.copy()
    ton_val_cols = [c for c in ton_tbl.columns if c not in key_cols]
    if ton_val_cols:
        if applied_hhv and applied_hhv > 0:
            ton_tbl[ton_val_cols] = ton_tbl[ton_val_cols] / float(applied_hhv)
        else:
            ton_tbl[ton_val_cols] = np.nan

    ton_tbl = ton_tbl.rename(columns={
        c: f"{c}(ton)" for c in ton_tbl.columns if c not in key_cols
    })

    out = pd.concat([base, pd.DataFrame({"": [np.nan] * len(base)}), ton_tbl], axis=1)
    return out

# ===== 런처 =====
def run(root: str = None,
        codes_csv: str = None,
        start_year: int = None,
        start_month: int = None,
        end_year: int = None,
        end_month: int = None,
        out_path: str = None,
        reserve_price: float = None,
        result_mode: str = None,             # "시간별" | "연도별"
        applied_hhv: float = None,            # 적용 발열량(HHV, kcal/kg)
        # GUI/CLI 스냅샷 인자
        snapshot: bool = False,
        snapshot_mode: str = "zip",           # "zip" | "copy"
        snapshot_name: str | None = None,     # 파일명(확장자 없이)
        snapshot_out: str | None = None,      # 저장 폴더
        use_default_snapshot_dir: bool = True # True면 out_path 폴더 사용
        ):
    """
    메인 실행 함수.
    - 자원코드(콤마), 시작/종료연도, (선택) 저장경로, 예비력 단가 입력
    - 스냅샷(zip/copy) 옵션 지원
    - 10개 시트 생성 및 저장
      * SMP(연도별): 2행 병합 헤더 (3행에 컬럼명 표시 안 함)
      * 발전비용: [발전/기동/연료] × [자원명들] 2행 병합 헤더 (1대일 때도 제목 표시)
    """

    # ---------- 1) 우선순위: 함수 인자 > CLI ----------
    do_snapshot = snapshot
    snap_mode   = (snapshot_mode or "zip").lower()
    snap_out    = snapshot_out
    snap_name   = snapshot_name

    # ---------------- argparse/입력 ----------------
    try:
        import argparse
        parser = argparse.ArgumentParser(add_help=False)
        parser.add_argument("--root")
        parser.add_argument("--codes")
        parser.add_argument("--start-year", type=int)
        parser.add_argument("--start-month", type=int)
        parser.add_argument("--end-year", type=int)
        parser.add_argument("--end-month", type=int)
        parser.add_argument("--out")
        parser.add_argument("--ru-price", type=float)
        parser.add_argument("--result-mode", choices=["시간별", "연도별"])
        parser.add_argument("--hhv", type=float)
        parser.add_argument("--snapshot", action="store_true")
        parser.add_argument("--snapshot-mode", choices=["zip","copy"])
        parser.add_argument("--snapshot-out")
        parser.add_argument("--snapshot-name")
        cli, _ = parser.parse_known_args()

        # 기본 입력
        root        = root        or cli.root or DEFAULT_ROOT
        codes_csv   = codes_csv   or cli.codes
        start_year  = start_year  or cli.start_year
        start_month = start_month or cli.start_month
        end_year    = end_year    or cli.end_year
        end_month   = end_month or cli.end_month
        out_path    = out_path    or cli.out
        reserve_price = reserve_price or cli.ru_price
        if getattr(cli, "result_mode", None) is not None and result_mode is None:
            result_mode = cli.result_mode
        if applied_hhv is None and getattr(cli, "hhv", None) is not None:
            applied_hhv = cli.hhv

        # 스냅샷(함수 인자가 우선, 없을 때만 CLI를 채용)
        if snapshot is False and getattr(cli, "snapshot", False):
            do_snapshot = True
        if not snapshot_mode and getattr(cli, "snapshot_mode", None):
            snap_mode = cli.snapshot_mode
        if not snapshot_out and getattr(cli, "snapshot_out", None):
            snap_out = cli.snapshot_out
        if not snapshot_name and getattr(cli, "snapshot_name", None):
            snap_name = cli.snapshot_name
    except Exception:
        # CLI 파서 실패 시에도 기본값 유지
        root = root or DEFAULT_ROOT

    # ---------- 2) 인터랙티브(터미널일 때만 질문) ----------
    try:
        is_tty = bool(sys.stdin) and sys.stdin.isatty()
    except Exception:
        is_tty = False

    if codes_csv is None and is_tty:
        codes_csv = input("자원코드를 콤마로 입력하세요 (예: 2731,2732): ").strip()
    if start_year is None and is_tty:
        start_year = int(input("시작년도를 입력하세요 (예: 2026): ").strip())
    if end_year is None and is_tty:
        end_year = int(input("종료년도를 입력하세요 (예: 2035): ").strip())
    if start_month is None and is_tty:
        start_month = int(input("시작월을 입력하세요 (1~12, 기본=1): ").strip() or "1")
    if end_month is None and is_tty:
        end_month = int(input("종료월을 입력하세요 (1~12, 기본=12): ").strip() or "12")
    if reserve_price is None and is_tty:
        reserve_price = float(input("예비력용량가치 단가를 입력하세요 (예: 1.0): ").strip())
    if result_mode is None and is_tty:
        rm = input("결과 취합 방식(시간별/연도별, 기본=시간별): ").strip()
        result_mode = rm if rm in ("시간별", "연도별") else "시간별"
    if applied_hhv is None and is_tty:
        applied_hhv = float(input("적용 발열량(HHV, kcal/kg)을 입력하세요 (예: 4500): ").strip())

    # 필수 입력 검증 (noconsole/비대화형 실행 시 None 비교 에러 방지)
    if start_year is None or end_year is None:
        raise ValueError(
            "시작연도/종료연도가 지정되지 않았습니다. "
            "EXE 실행 시 --start-year, --end-year 옵션을 입력하세요."
        )
    ys, ye = sorted([int(start_year), int(end_year)])
    ms = int(start_month or 1)
    me = int(end_month or 12)
    if not (1 <= ms <= 12 and 1 <= me <= 12):
        raise ValueError("시작월/종료월은 1~12 범위로 입력하세요.")

    if (ys, ms) > (ye, me):
        ys, ye, ms, me = ye, ys, me, ms

    if not codes_csv:
        raise ValueError(
            "자원코드가 지정되지 않았습니다. "
            "EXE 실행 시 --codes 옵션(예: --codes \"2731,2732\")을 입력하세요."
        )

    # 출력 파일 경로 기본값
    if out_path is None:
        out_path = f"SUDP_{ys}_{ye}.xlsx"

    root = Path(root)
    result_mode = (result_mode or "시간별").strip()
    if result_mode not in ("시간별", "연도별"):
        result_mode = "시간별"
    if applied_hhv is None:
        applied_hhv = 4500.0

    # 스냅샷 저장 폴더 기본값: 결과 엑셀과 같은 폴더
    if not snap_out or use_default_snapshot_dir:
        snap_out = str(Path(out_path).resolve().parent)

    # 터미널에서만 추가 질문 (GUI는 그대로 사용)
    if is_tty and snapshot is False:
        ans = input("현재 SUDP를 스냅샷으로 보존할까요? (Y/n) : ").strip().lower()
        do_snapshot = (ans == "" or ans.startswith("y"))
        if do_snapshot:
            tmp = input(f"스냅샷 모드 선택 [zip/copy] (기본: {snap_mode}) : ").strip().lower()
            if tmp in ("zip","copy"):
                snap_mode = tmp
            tmp = input(f"스냅샷 저장 폴더(Enter=기본:{snap_out}) : ").strip()
            if tmp:
                snap_out = tmp
            tmp = input("스냅샷 이름을 입력하세요 (빈칸=자동 생성) : ").strip()
            if tmp:
                snap_name = tmp

    # ---------- 3) 스냅샷 생성(단 한 번) ----------
    if do_snapshot:
        try:
            # 주의: 프로젝트에 정의된 스냅샷 함수 이름/시그니처에 맞추세요.
            # 제안: def snapshot_sudp(root:Path, ys:int, ye:int, out_dir:Path, mode="zip", snap_name=None) -> str
            snap_path = snapshot_sudp(
                root, ys, ye, Path(snap_out),
                mode=snap_mode, snap_name=snap_name
            )
            print(f"[스냅샷] {snap_mode} → {snap_path}")
        except Exception as e:
            print(f"[경고] 스냅샷 생성 실패: {e} (계속 진행합니다)")
                
    # ---------------- 카탈로그/타깃 선정 ----------------
    # (중요) RU/HOT 포함 전체 스캔
    catalog, code_to_id, id_to_name, id_to_capacity_x2, id_to_ru, id_to_hot = scan_catalog_all(root, ys, ms, ye, me)

    # id_to_code (역매핑) / code_to_hot 만들기
    id_to_code = {rid: code for code, rid in code_to_id.items()}

    # catalog에서 자원코드와 HOT기동비용을 바로 뽑아 code->hot 생성
    # (HOT기동비용 컬럼명이 다르면 여기를 바꿔주세요)
    if "자원코드" in catalog.columns and "HOT기동비용" in catalog.columns:
        code_to_hot = dict(
            catalog.dropna(subset=["자원코드"])  # 코드 없는 행 제외
                .loc[:, ["자원코드", "HOT기동비용"]]
                .dropna(subset=["HOT기동비용"])
                .itertuples(index=False, name=None)
        )
    else:
        code_to_hot = {}  # 없으면 0으로 처리됩니다.

    # 자원코드 -> 자원ID 목록
    input_codes = [c.strip() for c in codes_csv.split(",") if c.strip()]
    target_ids = []
    for code in input_codes:
        rid = code_to_id.get(code)
        if rid is None:
            print(f"[경고] 코드 '{code}'에 해당하는 자원이 기간({ys}~{ye}) 내 CD_DATA에서 발견되지 않았습니다.")
        else:
            target_ids.append(int(rid))
    target_ids = list(dict.fromkeys(target_ids))  # 중복 제거, 순서 보존
    if not target_ids:
        print("[종료] 대상 자원코드에 해당하는 자원이 없습니다.")
        return

    # ---------------- 시트 생성 ----------------
    sheets = {}
    # 기본 4종 (+발전비용은 아래에서 커스텀)
    for key in ["입찰량","발전량","송전량","연료사용량"]:
        sheets[key] = build_basic(root, ys, ms, ye, me, key, target_ids, id_to_name)

    # 발전비용: [발전/기동/연료] 분리 + 2단 헤더 렌더링용 이름목록
    sheets["발전비용"] = build_generation_cost(
    root, ys, ms, ye, me,
    target_ids=target_ids,
    id_to_name=id_to_name,
    id_to_hot=id_to_hot,
    )
    # 이용률
    sheets["이용률"] = build_utilization(root, ys, ms, ye, me, target_ids, id_to_name, id_to_capacity_x2)
    # 정산금
    sheets["정산금"] = build_settlement(root, ys, ms, ye, me, target_ids, id_to_name)
    # 예비력용량가치정산금
    sheets["예비력용량가치정산금"] = build_reserve_capacity_payment(
        root, ys, ms, ye, me, target_ids, id_to_name, id_to_ru, reserve_price)
    # SMP(시간별)
    sheets["SMP(시간별)"] = build_smp_hourly(root, ys, ms, ye, me, id_to_name)
    # SMP(연도별)
    sheets["SMP(연도별)"] = build_smp_yearly(root, ys, ms, ye, me, id_to_name)


    # --- 시간별/연도별 모드 처리 ---
    if result_mode == "시간별":
        # 전체 시간 인덱스(공란 패딩용)
        full_idx = build_full_hourly_index(ys, ms, ye, me)
        key_cols = ["연도","월","일","요일","시간"]

        # 시간 축을 가지는 시트들 (연도 요약 시트 제외)
        hourly_sheets = [
            "입찰량","발전량","송전량","이용률","연료사용량",
            "정산금","SMP(시간별)","예비력용량가치정산금","발전비용"
        ]

        for name in hourly_sheets:
            df = sheets.get(name)
            if df is None or df.empty:
                sheets[name] = full_idx.copy()
                continue
            sheets[name] = full_idx.merge(df, on=key_cols, how="left")
    else:
        yearly_targets = [
            "입찰량","발전량","송전량","발전비용",
            "이용률","정산금","예비력용량가치정산금"
        ]
        for name in yearly_targets:
            sheets[name] = aggregate_to_yearly(sheets.get(name), name)

        # 연료사용량은 전용 경로로 집계(문자열 숫자/콤마 대응 강화)
        sheets["연료사용량"] = aggregate_fuel_to_yearly(sheets.get("연료사용량"))

        # 연도별 결과에서는 SMP(시간별) 제외
        if "SMP(시간별)" in sheets:
            del sheets["SMP(시간별)"]

    # --- 연료사용량: [Mcal] + 공백열 + [ton] 2개 표로 변환 ---
    sheets["연료사용량"] = build_fuel_sheet_with_ton(sheets.get("연료사용량"), applied_hhv)

    # 단위 문자열
    unit_map = {
        "입찰량": "[MWh]",
        "발전량": "[MWh]",
        "송전량": "[MWh]",
        "발전비용": "[천 원]",
        "정산금": "[천 원]",
        "예비력용량가치정산금": "[천 원]",
        "이용률": "[%]",
        "SMP(시간별)": "[원/kWh]",
        "SMP(연도별)": "[원/kWh]",
    }

    # ---------------- 엑셀 저장 ----------------
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as xw:
        for name, df in sheets.items():

            # 공통 포맷
            wb  = xw.book
            fmt_head = wb.add_format({"align":"center","valign":"vcenter","bold":True,"border":1,"bg_color":"#EEEEEE"})
            fmt_key  = wb.add_format({"align":"center","valign":"vcenter","border":1,"num_format":"0"})
            fmt_int  = wb.add_format({"align":"center","valign":"vcenter","border":1,"num_format":"#,##0"})
            fmt_num  = wb.add_format({"align":"center","valign":"vcenter","border":1,"num_format":"#,##0.00"})
            fmt_txt  = wb.add_format({"align":"center","valign":"vcenter","border":1})
            fmt_center_noborder = wb.add_format({"align":"center","valign":"vcenter"})  # 중앙정렬(테두리X)
            fmt_unit = wb.add_format({"align":"right","valign":"vcenter","bold":True})

            # =========================================
            # 1) SMP(연도별) : 병합 헤더 + 데이터 포맷(연도=정수, 나머지=두째자리)
            # =========================================
            if name == "SMP(연도별)":
                base_df = (pd.DataFrame() if df is None else df).copy()

                # 숫자열 float 강제 (연도/자원명 제외)
                num_cols_names = [c for c in base_df.columns if c != "연도" and "자원명" not in c]
                if num_cols_names:
                    base_df[num_cols_names] = base_df[num_cols_names].apply(pd.to_numeric, errors="coerce")

                startrow = 3
                base_df.to_excel(xw, index=False, sheet_name=name, startrow=startrow, header=False)
                ws = xw.sheets[name]

                # 단위 표기(1행)
                ws.write(0, max(base_df.shape[1]-1, 0), unit_map.get(name, ""), fmt_unit)

                # 병합 헤더
                ws.merge_range(1, 0, 2, 0, "연도", fmt_head)
                groups = [("경인", 1), ("비경인", 7), ("제주", 13)]
                for title, c0 in groups:
                    ws.merge_range(1, c0, 1, c0+5, title, fmt_head)
                    for i, s in enumerate(["최대값","자원명","최소값","자원명","평균값","가중평균값"]):
                        ws.write(2, c0+i, s, fmt_head)

                nrows, ncols = base_df.shape

                # (A) 먼저 모든 열 너비만 통일(포맷=None로 덮어쓰기 방지)
                ws.set_column(0, max(ncols-1, 0), 8)  # 포맷 인수 생략

                # (B) 각 열에 ‘표시형식’ 부여: 연도는 정수, 자원명은 텍스트, 그 외 숫자둘째자리
                fmt_num_center = wb.add_format({"align":"center","valign":"vcenter","num_format":"#,##0.00"})
                fmt_int_center = wb.add_format({"align":"center","valign":"vcenter","num_format":"0"})
                fmt_txt_center = wb.add_format({"align":"center","valign":"vcenter"})

                for i, cname in enumerate(base_df.columns):
                    if i == 0:
                        ws.set_column(i, i, 8, fmt_int_center)       # 연도
                    elif "자원명" in cname:
                        ws.set_column(i, i, 8, fmt_txt_center)       # 자원명
                    else:
                        ws.set_column(i, i, 8, fmt_num_center)       # 숫자(둘째자리)

                # (C) 데이터 구간에만 테두리(표처럼)
                if nrows > 0 and ncols > 0:
                    r1, r2 = startrow, startrow + nrows - 1
                    border_only = wb.add_format({"border":1})
                    ws.conditional_format(r1, 0, r2, ncols-1, {"type":"no_blanks","format":border_only})

                continue


            # =========================================
            # 2) 정산금 : 2단 헤더(1행=자원명 merge, 2행=MEP/MAP/MWP/합계) + 본문 포맷
            # =========================================
            if name == "정산금":
                base_df = (pd.DataFrame() if df is None else df).copy()
                startrow = 3  # 데이터는 4행부터 쓴다
                base_df.to_excel(xw, index=False, sheet_name=name, startrow=startrow, header=False)
                ws = xw.sheets[name]

                # 단위 표기(1행)
                ws.write(0, max(base_df.shape[1]-1, 0), unit_map.get(name, ""), fmt_unit)

                # 키 5열(연/월/일/요일/시간) 세로 병합 헤더
                key_cols = [c for c in ["연도","월","일","요일","시간"] if c in base_df.columns]
                for i, lab in enumerate(key_cols):
                    ws.merge_range(1, i, 2, i, lab, fmt_head)

                nrows, ncols = base_df.shape
                ws.set_column(0, max(ncols-1, 0), 14, fmt_center_noborder)  # 전열 중앙정렬

                # 자원명 그룹 추출
                # base_df의 5열 이후는 [MEP(이름), MAP(이름), MWP(이름), 합계(이름)] 반복
                import re
                group_cols = base_df.columns[len(key_cols):]
                names_order = []
                name_to_cols = {}

                for i, col in enumerate(group_cols, start=len(key_cols)):
                    # 괄호 안 이름 추출
                    m = re.search(r"\((.+?)\)", col)
                    name = m.group(1) if m else col
                    # 앞쪽 접두로 MEP/MAP/MWP/합계 판단
                    if name not in name_to_cols:
                        name_to_cols[name] = {"MEP":None,"MAP":None,"MWP":None,"합계":None}
                        names_order.append(name)
                    if col.startswith("MEP("):
                        name_to_cols[name]["MEP"] = i
                    elif col.startswith("MAP("):
                        name_to_cols[name]["MAP"] = i
                    elif col.startswith("MWP("):
                        name_to_cols[name]["MWP"] = i
                    elif col.startswith("합계("):
                        name_to_cols[name]["합계"] = i

                # 1행: 자원명 병합, 2행: MEP/MAP/MWP/합계 라벨
                cur_col = len(key_cols)  # ★ 항상 키 다음 칼럼(=5)에서 시작
                for nm in names_order:
                    # 4개 폭
                    ws.merge_range(1, cur_col, 1, cur_col+3, nm, fmt_head)
                    ws.write(2, cur_col+0, "MEP", fmt_head)
                    ws.write(2, cur_col+1, "MAP", fmt_head)
                    ws.write(2, cur_col+2, "MWP", fmt_head)
                    ws.write(2, cur_col+3, "합계", fmt_head)
                    cur_col += 4

                # 본문 포맷: 키 5열 정수, 나머지 수치 #,##0.00
                if nrows > 0 and ncols > 0:
                    r1, r2 = startrow, startrow + nrows - 1
                    key_last = min(len(key_cols)-1, ncols-1)
                    if key_last >= 0:
                        ws.conditional_format(r1, 0, r2, key_last, {"type":"no_blanks","format":fmt_key})
                    if ncols > len(key_cols):
                        ws.conditional_format(r1, len(key_cols), r2, ncols-1, {"type":"no_blanks","format":fmt_num})

                # 열 폭: 0~4 동일(8), 이후 14
                try:
                    ws.set_column(0, len(key_cols)-1, 8,  fmt_center_noborder)
                    if ncols > len(key_cols):
                        ws.set_column(len(key_cols), ncols-1, 14, fmt_center_noborder)
                except Exception:
                    pass
                continue


            # =========================================
            # 3) 발전비용 : 2단 병합 헤더(1~2행) + 본문 포맷
            # =========================================
            if name == "발전비용":
                base_df = (pd.DataFrame() if df is None else df).copy()

                # --- 1) 컬럼 파싱으로 자원명/지표 매핑 만들기 ---
                import re
                key_cols = [c for c in ["연도","월","일","요일","시간"] if c in base_df.columns]
                data_cols = [c for c in base_df.columns if c not in key_cols]

                metrics = ["발전비용","기동비용","연료비용"]
                metric_to_map = {m:{} for m in metrics}
                names_order = []  # 등장 순서 유지

                for c in data_cols:
                    m = None
                    for mt in metrics:
                        if c.startswith(mt + "|"):
                            m = mt; break
                    if m is None:
                        continue
                    nm = c.split("|",1)[1]
                    if nm not in names_order:
                        names_order.append(nm)
                    metric_to_map[m][nm] = c  # 실제 컬럼명 저장

                # --- 2) 안전한 출력 순서 구성 (없는 조합은 빈 열 생성) ---
                ordered_cols = key_cols[:]
                for nm in names_order:
                    for mt in metrics:
                        colname = metric_to_map[mt].get(nm)
                        if colname is None:
                            base_df[f"{mt}|{nm}"] = pd.NA
                            colname = f"{mt}|{nm}"
                        ordered_cols.append(colname)

                base_df = base_df[ordered_cols]

                # --- 3) 엑셀 쓰기 ---
                startrow = 3
                base_df.to_excel(xw, index=False, sheet_name=name, startrow=startrow, header=False)
                ws = xw.sheets[name]

                # 단위 표기(1행)
                ws.write(0, max(base_df.shape[1]-1, 0), unit_map.get(name, ""), fmt_unit)

                # 2,3행 키 5열 세로 병합
                for i, lab in enumerate(key_cols):
                    ws.merge_range(1, i, 2, i, lab, fmt_head)

                # 1행: 발전기명 병합, 2행: 발전/기동/연료
                col0 = len(key_cols)
                for nm in names_order:
                    ws.merge_range(1, col0, 1, col0 + len(metrics) - 1, nm, fmt_head)
                    for j, mt in enumerate(metrics):
                        ws.write(2, col0 + j, mt, fmt_head)
                    col0 += len(metrics)

                # 본문 포맷 (정렬·서식)
                nrows, ncols = base_df.shape
                ws.set_column(0, ncols-1, 14, fmt_center_noborder)
                if nrows > 0 and ncols > 0:
                    r1, r2 = startrow, startrow + nrows - 1
                    key_last = len(key_cols) - 1
                    # 키 열 정수
                    if key_last >= 0:
                        ws.conditional_format(r1, 0, r2, key_last, {"type":"no_blanks","format":fmt_key})
                    # 값들 #,##0.00
                    if ncols > len(key_cols):
                        ws.conditional_format(r1, len(key_cols), r2, ncols-1, {"type":"no_blanks","format":fmt_num})

                # 키 열 폭 통일
                if ncols > 0:
                    key_last = min(len(key_cols)-1, ncols-1)
                    if key_last >= 0:
                        ws.set_column(0, key_last, 8, fmt_center_noborder)
                    if ncols > len(key_cols):
                        ws.set_column(len(key_cols), ncols-1, 14, fmt_center_noborder)
                continue


            # =========================================
            # 4) 연료사용량 : [Mcal] + 공백열 + [ton], 단위 2개 표기
            # =========================================
            if name == "연료사용량":
                base_df = (pd.DataFrame() if df is None else df).copy()
                base_df.to_excel(xw, index=False, sheet_name=name, startrow=2, header=False)
                ws = xw.sheets[name]

                nrows, ncols = base_df.shape
                split_col = list(base_df.columns).index("") if "" in base_df.columns else -1

                # 헤더(2행)
                for c, col in enumerate(base_df.columns):
                    if split_col >= 0 and c == split_col:
                        ws.write(1, c, "", fmt_center_noborder)
                    else:
                        ws.write(1, c, col, fmt_head)

                # 단위(1행) - 첫 표, 둘째 표 각각 마지막 컬럼
                if split_col > 0:
                    ws.write(0, split_col - 1, "[Mcal]", fmt_unit)
                    ws.write(0, ncols - 1, "[ton]", fmt_unit)
                elif ncols > 0:
                    ws.write(0, ncols - 1, "[Mcal]", fmt_unit)

                # 열 폭
                if ncols > 0:
                    def _col_width(series: pd.Series, header: str, min_w: int = 8, max_w: int = 28) -> int:
                        sval = series.astype(str).replace({"nan": "", "None": ""}) if series is not None else pd.Series(dtype=str)
                        max_len = max([len(str(header))] + ([int(sval.map(len).max())] if not sval.empty else [0]))
                        return max(min_w, min(max_w, max_len + 2))

                    key_names = ["연도", "월", "일", "요일", "시간"]

                    if split_col >= 0:
                        # 왼쪽 표 자동 너비
                        for i in range(0, split_col):
                            cname = str(base_df.columns[i])
                            width = _col_width(base_df.iloc[:, i], cname, min_w=(8 if cname in key_names else 10))
                            ws.set_column(i, i, width, fmt_center_noborder)

                        # 구분 공백열
                        ws.set_column(split_col, split_col, 3, fmt_center_noborder)

                        # 오른쪽 표 자동 너비
                        for i in range(split_col + 1, ncols):
                            cname = str(base_df.columns[i])
                            is_right_key = cname in key_names
                            width = _col_width(base_df.iloc[:, i], cname, min_w=(8 if is_right_key else 10))
                            ws.set_column(i, i, width, fmt_center_noborder)
                    else:
                        for i in range(0, ncols):
                            cname = str(base_df.columns[i])
                            width = _col_width(base_df.iloc[:, i], cname, min_w=(8 if cname in key_names else 10))
                            ws.set_column(i, i, width, fmt_center_noborder)

                # 본문 포맷
                if nrows > 0 and ncols > 0:
                    r1, r2 = 2, nrows + 1
                    key_cols_left = [c for c in ["연도","월","일","요일","시간"] if c in base_df.columns[: (split_col if split_col >= 0 else ncols)]]
                    key_len = len(key_cols_left)
                    if split_col >= 0:
                        left_key_last = min(key_len - 1, split_col - 1)
                        if left_key_last >= 0:
                            ws.conditional_format(r1, 0, r2, left_key_last, {"type":"no_blanks","format":fmt_key})
                        if split_col - 1 >= key_len:
                            ws.conditional_format(r1, key_len, r2, split_col - 1, {"type":"no_blanks","format":fmt_num})

                        right_start = split_col + 1
                        right_key_last = right_start + key_len - 1
                        right_val_start = right_key_last + 1
                        if right_key_last >= right_start:
                            ws.conditional_format(r1, right_start, r2, right_key_last, {"type":"no_blanks","format":fmt_key})
                        if right_val_start <= ncols - 1:
                            ws.conditional_format(r1, right_val_start, r2, ncols-1, {"type":"no_blanks","format":fmt_num})
                    else:
                        key_last = min(key_len - 1, ncols - 1)
                        if key_last >= 0:
                            ws.conditional_format(r1, 0, r2, key_last, {"type":"no_blanks","format":fmt_key})
                        if ncols - 1 >= key_len:
                            ws.conditional_format(r1, key_len, r2, ncols-1, {"type":"no_blanks","format":fmt_num})
                continue


            # =========================================
            # 4) 일반 시트 : 헤더 1행만, 데이터 포맷(정수/소수 분리) + 중앙정렬 기본
            # =========================================
            df = pd.DataFrame() if df is None else df
            df.to_excel(xw, index=False, sheet_name=name, startrow=1)  # 헤더는 2행
            ws = xw.sheets[name]

            # 단위(1행, 마지막 컬럼 우측정렬)
            if df.shape[1] > 0:
                ws.write(0, df.shape[1]-1, unit_map.get(name, ""), fmt_unit)

            # 헤더(2행, 정확 범위만 회색)
            for c, col in enumerate(df.columns):
                ws.write(1, c, col, fmt_head)

            nrows, ncols = df.shape

            # 전열 중앙정렬 기본
            if ncols > 0:
                ws.set_column(0, ncols-1, 14, fmt_center_noborder)

            if nrows > 0 and ncols > 0:
                r1, r2 = 2, nrows + 1  # 데이터 구간
                key_cols = [c for c in ["연도","월","일","요일","시간"] if c in df.columns]
                key_len = len(key_cols)
                last_key = min(key_len - 1, ncols-1)
                # 키 열: 정수
                if last_key >= 0:
                    ws.conditional_format(r1, 0, r2, last_key, {"type":"no_blanks","format":fmt_key})
                # 나머지 열: 자원명(텍스트) vs 값(두째자리)
                for i in range(key_len, ncols):
                    colname = df.columns[i]
                    if "자원명" in colname:
                        ws.conditional_format(r1, i, r2, i, {"type":"no_blanks","format":fmt_txt})
                    else:
                        ws.conditional_format(r1, i, r2, i, {"type":"no_blanks","format":fmt_num})

            # 열 폭: 키 열 8, 이후 14
            try:
                if ncols > 0:
                    key_cols = [c for c in ["연도","월","일","요일","시간"] if c in df.columns]
                    key_len = len(key_cols)
                    key_last = min(key_len-1, ncols-1)
                    if key_last >= 0:
                        ws.set_column(0, key_last, 8, fmt_center_noborder)
                    if ncols > key_len:
                        ws.set_column(key_len, ncols-1, 14, fmt_center_noborder)
            except Exception:
                pass


    print(f"[완료] 엑셀 생성: {out_path}")


if __name__ == "__main__":
    try:
        # 인자 없이 실행(특히 --noconsole EXE 더블클릭)하면 GUI를 띄운다.
        if len(sys.argv) == 1:
            try:
                from sudp_gui import main as gui_main
                gui_main()
                raise SystemExit(0)
            except Exception:
                # GUI 로딩 실패 시 기존 CLI 경로로 폴백
                pass
        run()
    except Exception as e:
        msg = f"[오류] {e}"
        try:
            Path("sudp_error.log").write_text(msg + "\n", encoding="utf-8")
        except Exception:
            pass
        print(msg)
        raise SystemExit(1)
