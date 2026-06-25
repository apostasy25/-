# -*- coding: utf-8 -*-
"""
누적 순매출 원장 + writeback.

· 작품별(표준작품명) 누적 순매출(기타수익 제외)을 정산월별로 누적 관리.
· 월 정산 완료(검수 통과) 후 writeback 으로 갱신 — (표준작품명, 정산월) 키 idempotent
  (재실행 시 같은 월 행을 덮어써서 중복 반영 방지).
· 누적구간 예외 RS 는 이 누적값을 기준으로 한계세율식 적용.
· 감사추적: 적용규칙ID·적용RS_상세·누적순매출(전/후)·임계통과·당월원작료를 행마다 보존.
"""
import datetime
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "gui", "master"))
import master_io as MIO

LEDGER = "누적정산기준매출원장"
LHDR = ["표준작품명", "정산월", "당월_정산기준매출_기타제외",
        "누적정산기준매출_전", "누적정산기준매출_후", "적용규칙ID", "RS적용base",
        "적용RS_상세", "임계통과", "당월원작료", "갱신일시", "비고"]


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return 0.0


def _data_rows(rows):
    """안내/빈 행 제외한 실제 원장 행."""
    out = []
    for r in rows:
        nm = str(r.get("표준작품명") or "").strip()
        if not nm or nm.startswith("[") or nm.startswith("·"):
            continue
        out.append(r)
    return out


def read_ledger(master_path):
    """(history[list[dict]], latest{표준작품명: 최신 누적_후})."""
    try:
        _, rows = MIO.read_sheet(master_path, LEDGER)
    except Exception:
        return [], {}
    hist = _data_rows(rows)
    latest = {}
    for r in hist:
        nm = str(r["표준작품명"]).strip()
        wol = str(r.get("정산월") or "")
        cum = _num(r.get("누적정산기준매출_후"))
        if nm not in latest or wol > latest[nm][0]:
            latest[nm] = (wol, cum)
    return hist, {k: v[1] for k, v in latest.items()}


def prev_cumulative(master_path, 정산월):
    """작품별, 해당 정산월 직전까지의 누적_후 (writeback 누적_전 기준)."""
    hist, _ = read_ledger(master_path)
    prev = {}
    for r in hist:
        nm = str(r["표준작품명"]).strip()
        wol = str(r.get("정산월") or "")
        if wol < str(정산월):
            if nm not in prev or wol > prev[nm][0]:
                prev[nm] = (wol, _num(r.get("누적정산기준매출_후")))
    return {k: v[1] for k, v in prev.items()}


def parse_tiers(rule):
    """예외규칙 행 → [(임계, RS), ...] 오름차순."""
    tiers = []
    for tcol, rcol in [("임계1", "RS1"), ("임계2", "RS2")]:
        t, rs = rule.get(tcol), rule.get(rcol)
        if t not in (None, "") and rs not in (None, ""):
            tiers.append((_num(t), _num(rs)))
    return sorted(tiers)


def tiered_원작료(당월_순매출, 당월_정산기준, 누적_before, base_rs, tiers, base_mode="정산기준매출"):
    """누적 순매출 구간별 한계세율 원작료. 반환: (원작료, 적용RS_상세, 임계통과)."""
    pts = sorted([(0.0, base_rs)] + [(float(t), float(rs)) for t, rs in tiers])
    bounds = [p[0] for p in pts] + [float("inf")]
    start, end = 누적_before, 누적_before + 당월_순매출
    원작료, detail, passed = 0.0, [], []
    for i in range(len(pts)):
        lo, hi = bounds[i], bounds[i + 1]
        seg = max(0.0, min(end, hi) - max(start, lo))
        if seg > 0:
            rs = pts[i][1]
            if base_mode == "순매출":
                원작료 += seg * rs
            else:                                   # 정산기준매출 base
                frac = (seg / 당월_순매출) if 당월_순매출 else 0.0
                원작료 += frac * 당월_정산기준 * rs
            detail.append(f"{int(pts[i][0]):,}↑×{rs}" if pts[i][0] > 0 else f"기본×{rs}")
            if pts[i][0] > 0:
                passed.append(f"{int(pts[i][0]):,}통과")
    return 원작료, " + ".join(detail), ("; ".join(passed) if passed else f"미통과(누적 {int(end):,})")


def writeback(master_path, 정산월, entries, do_backup=True):
    """월 정산 완료 후 원장 갱신. entries: list[dict] with
       표준작품명·당월_정산기준매출·당월_순매출_기타제외·적용규칙ID·RS적용base·적용RS_상세·임계통과·당월원작료.
       (표준작품명, 정산월) idempotent: 같은 월 행 있으면 덮어씀."""
    if do_backup:
        MIO.backup_master(master_path)
    hist, _ = read_ledger(master_path)
    work_set = {e["표준작품명"] for e in entries}
    # 이번 작품들의 '이번 정산월' 기존 행 제거(중복 방지) — 나머지는 보존
    keep = [r for r in hist
            if not (str(r.get("표준작품명")).strip() in work_set
                    and str(r.get("정산월")) == str(정산월))]
    prev = {}
    for r in keep:
        nm = str(r["표준작품명"]).strip()
        wol = str(r.get("정산월") or "")
        if wol < str(정산월):
            if nm not in prev or wol > prev[nm][0]:
                prev[nm] = (wol, _num(r.get("누적정산기준매출_후")))
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    new_rows = []
    for e in entries:
        nm = e["표준작품명"]
        before = prev.get(nm, ("", 0.0))[1]
        당월 = _num(e["당월_정산기준매출_기타제외"])
        after = before + 당월
        new_rows.append({
            "표준작품명": nm, "정산월": 정산월,
            "당월_정산기준매출_기타제외": round(당월),
            "누적정산기준매출_전": round(before), "누적정산기준매출_후": round(after),
            "적용규칙ID": e.get("적용규칙ID", ""), "RS적용base": e.get("RS적용base", ""),
            "적용RS_상세": e.get("적용RS_상세", ""), "임계통과": e.get("임계통과", ""),
            "당월원작료": round(_num(e.get("당월원작료"))), "갱신일시": now, "비고": e.get("비고", ""),
        })
    MIO.save_sheet(master_path, LEDGER, LHDR, keep + new_rows, do_backup=False)
    return new_rows
