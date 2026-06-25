# -*- coding: utf-8 -*-
"""
M2 — 금액 엔진.

지급대상(M1) + 매출/기타수익(M0) + 마스터(업체·선인세MG·이월)를 받아
업체별 정산금액·MG 상계·이월·원천징수·실지급액을 산출한다.

핵심 규칙(확정):
  · 원작사 지급액(연재) = 매출 '원작료(C)' = 정산기준매출 × 원작사RS.
  · 기타수익(광고) 원작료 = 금액 × 원작사RS (금액은 RS 적용 전 순매출).
  · MG 상계(나 방식): 차감 중 작품도 당월 정산금액(= 원작료 C)을 누적상계액에 가산.
        누적_new = min(누적_prev + 당월, MG총액)
        잔여(마스터) = MG총액 − 누적_new   (정산서 표시값 = −잔여)
        실지급(해당작품) = max(0, (누적_prev + 당월) − MG총액)   ← 초과분만 지급
        잔여 > 0 인 동안 표시 "MG 차감중", 지급 0.
  · 원천징수(개인): 3.3%(소득세 3% + 지방 0.3%), 10원 절사. 기타소득이면 8.8%.
  · 부가세(사업자): 국내 면세(계산서)=0 / 해외 과세(세금계산서)=10%. 혼합·해외는 수기 플래그.
  · 이월: 이월마스터에서 (합산월=이번 정산월 & 증빙상태=수취)인 건의 [자동] 이월 정산금액을 가산.
"""
from collections import defaultdict

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "gui", "master"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import master_io as MIO
import ledger as LG


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return 0.0


def round_down_10(x):
    """10원 미만 절사 (원천징수 관행)."""
    return int(x // 10 * 10)


def _load_masters(master_path):
    _, ven = MIO.read_sheet(master_path, "업체마스터")
    vmap = {str(v.get("업체ID")).strip(): v for v in ven if v.get("업체ID")}

    _, mgrows = MIO.read_sheet(master_path, "선인세MG마스터")
    mg = {}
    for r in mgrows:
        vid = str(r.get("업체ID") or "").strip()
        nm = str(r.get("작품명(표준)") or "").strip()
        if vid.startswith("V") and nm:
            mg[(vid, nm)] = {
                "MG총액": _num(r.get("MG총액")),
                "누적상계액": _num(r.get("누적상계액")),
            }

    _, iwrows = MIO.read_sheet(master_path, "이월마스터")
    iwol = []
    for r in iwrows:
        vid = str(r.get("업체ID") or "").strip()
        if vid.startswith("V"):
            iwol.append(r)

    # 예외규칙(작품별) — 하드코딩 없이 마스터 테이블에서 로드
    ex = defaultdict(list)
    try:
        _, exrows = MIO.read_sheet(master_path, "예외규칙마스터")
        for r in exrows:
            nm = str(r.get("표준작품명") or "").strip()
            유형 = str(r.get("규칙유형") or "").strip()
            if nm and 유형 in ("누적순매출구간", "수익종류", "MG완료후", "계약기간"):
                ex[nm].append(r)
    except Exception:
        pass
    return vmap, mg, iwol, ex


def _etc_rs(std, 종류, base_rs, ex):
    """기타수익 한 행의 적용 RS — 수익종류 예외가 있으면 그 RS, 없으면 base."""
    for r in ex.get(std, []):
        if str(r.get("규칙유형")) == "수익종류" and str(r.get("대상스트림")) == "기타수익":
            want = str(r.get("수익종류") or "").strip()
            if want and want in str(종류 or ""):
                ap = r.get("적용RS")
                if ap not in (None, ""):
                    return float(ap)
    return base_rs


def compute(company, 정산서월, master_path, R, m1_result):
    """company 한 곳의 금액 산출. R=M0.load_all 결과, m1_result=M1.match_sales 결과."""
    vmap, mg, iwol, ex = _load_masters(master_path)
    지급월 = R["지급월"]

    # 1) 매출을 표준작품명(std)별로 집계: 정산기준매출·순매출(기타제외)·원작료(C)
    std_of = {}
    for w in m1_result["지급대상"]:
        std_of[w["작품명"]] = (str(w.get("원작사_업체ID") or "").strip(),
                              str(w.get("마스터작품명") or ""))
    agg = defaultdict(lambda: {"정산기준": 0.0, "순매출": 0.0, "C": 0.0, "vid": "",
                               "by_plat": defaultdict(float)})
    for r in R["매출"]:
        if r["회사"] == company and "원작료" in str(r["항목"]) and r["작품명"] in std_of:
            vid, std = std_of[r["작품명"]]
            a = agg[std]
            a["vid"] = vid
            a["정산기준"] += _num(r.get("정산기준매출"))
            a["순매출"] += _num(r.get("정산순매출"))
            a["C"] += _num(r.get("원작료"))
            a["by_plat"][str(r.get("플랫폼") or "").strip()] += _num(r.get("정산기준매출"))

    # 2) 기타수익 당월지급분 원작료 = 금액 × 원작사RS (수익종류 예외 반영)
    rs_by_std = {}
    for w in m1_result["지급대상"]:
        rs_by_std[str(w.get("마스터작품명") or "")] = _num(w.get("RS율"))
    etcC = defaultdict(float)
    for e in R.get("기타수익", []):
        if e.get("회사") != company:
            continue
        if str(e.get("지급월_norm") or "") != str(정산서월):
            continue
        if not any(k in str(e.get("종류구분") or "") for k in ("광고", "IP")):
            continue
        nm = str(e.get("작품명") or "")
        base_rs = rs_by_std.get(nm, 0.0)
        rs = _etc_rs(nm, e.get("종류"), base_rs, ex)
        etcC[nm] += _num(e.get("금액")) * rs

    # 3) 연재 원작료 결정: 누적구간 예외면 원장 누적 기준 한계세율, 아니면 입력 원작료(C)
    tier_rules = {}
    for nm, rules in ex.items():
        for rl in rules:
            if str(rl.get("규칙유형")) == "누적순매출구간":
                tier_rules[nm] = rl
    prev_cum = LG.prev_cumulative(master_path, 정산서월)   # 작품별 직전 누적 정산기준매출
    연재_by_std, audit, ledger_entries = {}, {}, []
    for std, a in agg.items():
        if std in tier_rules:
            rl = tier_rules[std]
            base_rs = rs_by_std.get(std, 0.0)
            tiers = LG.parse_tiers(rl)
            base_mode = str(rl.get("RS적용base") or "정산기준매출")
            before = prev_cum.get(std, 0.0)
            # 계약서 제5조: '순매출' = 플랫폼/유통사 정산액 = 정산기준매출(B). 누적·base 모두 B.
            범위 = str(rl.get("구간적용범위") or "전체").strip()
            if 범위.startswith("최초서비스"):
                # 최초서비스 플랫폼에만 구간 적용, 나머지 플랫폼은 기본 RS
                plat = 범위.split(":", 1)[1].split("(")[0].strip() if ":" in 범위 else ""
                scoped = a["by_plat"].get(plat, 0.0)
                other = a["정산기준"] - scoped
                won_s, detail, passed = LG.tiered_원작료(scoped, scoped, before,
                                                       base_rs, tiers, base_mode)
                won = won_s + other * base_rs
                cum_amt = scoped                          # 누적은 최초서비스 플랫폼만
                detail = f"[최초:{plat}]{detail} + 기타플랫폼×{base_rs}"
            else:                                          # 전체 플랫폼 합산
                cum_amt = a["정산기준"]
                won, detail, passed = LG.tiered_원작료(cum_amt, cum_amt, before,
                                                     base_rs, tiers, base_mode)
            연재_by_std[std] = won
            audit[std] = {"규칙ID": rl.get("규칙ID"), "적용RS_상세": detail, "임계통과": passed,
                          "누적_전": round(before), "누적_후": round(before + cum_amt),
                          "base": base_mode, "범위": 범위}
            ledger_entries.append({
                "표준작품명": std, "당월_정산기준매출_기타제외": cum_amt,
                "적용규칙ID": rl.get("규칙ID"), "RS적용base": base_mode,
                "적용RS_상세": detail, "임계통과": passed, "당월원작료": won})
        else:
            연재_by_std[std] = a["C"]                       # 일반: 입력 원작료(C) 신뢰

    # 업체 → 작품(std)별 정산금액(연재 + 기타) 집계
    by_vendor = defaultdict(lambda: defaultdict(float))
    for std, a in agg.items():
        by_vendor[a["vid"]][std] += 연재_by_std.get(std, a["C"])
    for std, amt in etcC.items():
        for w in m1_result["지급대상"]:
            if str(w.get("마스터작품명") or "") == std:
                by_vendor[str(w.get("원작사_업체ID")).strip()][std] += amt
                break

    # 4) 업체별 계산
    results = []
    for vid, works in by_vendor.items():
        v = vmap.get(vid, {})
        유형 = str(v.get("유형") or "")
        소속 = str(v.get("소속") or "")
        증빙 = str(v.get("증빙구분") or "")
        items = []
        정산금액합 = 0.0
        for std, amt in sorted(works.items()):
            line = {"작품": std, "정산금액": round(amt)}
            key = (vid, std)
            if key in mg:                              # MG 상계
                prev = mg[key]["누적상계액"]
                total = mg[key]["MG총액"]
                누적_new = min(prev + amt, total)
                초과 = max(0.0, (prev + amt) - total)
                잔여 = total - 누적_new                 # 마스터 부호(양수=보류)
                line.update({
                    "MG": True, "MG총액": round(total),
                    "누적상계_전": round(prev), "당월차감": round(min(amt, total - prev) if total > prev else 0),
                    "누적상계_후": round(누적_new),
                    "잔여_마스터": round(잔여), "잔여_정산서": -round(잔여),
                    "실지급분": round(초과),
                    "표시": "MG 차감중" if 잔여 > 0 else "MG 완료(초과분 지급)",
                })
                정산금액합 += 초과
            else:
                line["실지급분"] = round(amt)
                정산금액합 += amt
            items.append(line)

        # 5) 이월 가산 (합산월 = 이번 정산월 & 증빙상태 = 수취)
        이월금액 = 0.0
        for r in iwol:
            if str(r.get("업체ID")).strip() == vid \
               and str(r.get("합산월(이번 합산 정산월)") or "") == 정산서월 \
               and "수취" in str(r.get("증빙상태") or ""):
                이월금액 += _num(r.get("[자동] 이월 정산금액"))

        base = 정산금액합 + 이월금액                      # 원천징수/부가세 대상(공급가·소득)

        # 6) 원천징수 / 부가세
        원천징수 = 부가세 = 0
        수기플래그 = None
        if 유형 == "개인":
            세율 = 0.088 if "기타소득" in str(v.get("정산유형") or "") else 0.033
            원천징수 = round_down_10(base * 세율)
            실지급 = round(base) - 원천징수
        elif 유형 == "사업자":
            if 소속 == "해외" or "혼합" in 증빙:
                수기플래그 = "해외/혼합 증빙 — 부가세·원천징수 수기 검토"
                실지급 = round(base)
            elif "세금계산서" in 증빙:                  # 과세
                부가세 = round(base * 0.1)
                실지급 = round(base) + 부가세
            else:                                       # 계산서(면세) 등
                실지급 = round(base)
        else:
            실지급 = round(base)

        results.append({
            "업체ID": vid, "업체명": v.get("업체명"), "유형": 유형, "소속": 소속, "증빙": 증빙,
            "작품수": len(items), "items": items,
            "정산금액합": round(정산금액합), "이월": round(이월금액),
            "원천징수": 원천징수, "부가세": 부가세,
            "실지급액": 실지급, "수기플래그": 수기플래그,
        })
    results.sort(key=lambda x: x["업체ID"])
    return {"company": company, "정산서월": 정산서월, "지급월": 지급월, "업체": results,
            "누적구간_감사": audit, "ledger_entries": ledger_entries}
