# -*- coding: utf-8 -*-
"""
M1 매칭 엔진 — 전체 정산 엔진(방향 B)의 두 번째 단계.

역할: M0가 읽은 매출(항목=원작료) 작품을 '작품Alias마스터'(= 정답, Single Source
      of Truth)에 매칭해 원작사·정산규칙을 확정하고, 못 맞춘 건 ORPHAN으로,
      매출의 원작주체와 마스터 원작사가 다른 건 Mismatch로 리포트한다.

판정 원칙 (이전 대화 확정)
  · 지급대상 = 항목 "원작료" AND 마스터 상태 ≠ 해지/종료
  · 매칭은 작품Alias마스터가 정답. 매출의 '웹소설본부 원작여부' 컬럼은 참고/검증·
    자체(웹소설본부) 라우팅 용도로만 사용.
  · 흐름: ①작품명→마스터 매칭 ②성공→마스터 원작사·규칙 ③실패→ORPHAN
          ④매출 원작주체≠마스터 원작사 → Mismatch
  · 원작사 변경·통합계약·공동저자·Alias 추가 등은 전부 마스터에서 제어(코드 불변).
"""
import re

try:
    from gui.master import master_io as M
except Exception:                                    # 평탄 import 대비
    import master_io as M


_SELF = "웹소설본부"                                 # 자체 정산(외부 발송 X)
_BLANK_HINT = ("", "-", "0", "none", "nan")


def _norm(s):
    """작품명 정규화 키: 소문자 + 공백·구두점 제거(괄호 안 글자는 보존)."""
    if s is None:
        return ""
    t = str(s).lower()
    t = re.sub(r"[\s()\[\]<>·,.!?:;'\"~\-_/\\|]", "", t)
    return t


# ── 마스터 인덱스 ───────────────────────────────────────────
def build_master_index(master_path):
    """작품Alias마스터 → {정규화키: [작품레코드]} + 업체ID→정보 + 중복키 경고."""
    works, wrows = M.read_sheet(master_path, "작품Alias마스터")
    vrows = M.read_sheet(master_path, "업체마스터")[1]
    vendor = {}
    for v in vrows:
        vid = str(v.get("업체ID") or "").strip()
        if vid:
            vendor[vid] = {"업체명": v.get("업체명"), "회사": v.get("정산주체(회사)"),
                           "유형": v.get("유형"), "소속": v.get("소속")}
    index, dup = {}, []
    for w in wrows:
        title = w.get("표준작품명")
        if not title:
            continue
        vid = str(w.get("원작사_업체ID") or "").strip()
        rec = {
            "표준작품명": title,
            "원작사_업체ID": vid,
            "원작사명": vendor.get(vid, {}).get("업체명"),
            "회사": vendor.get(vid, {}).get("회사"),
            "유형": vendor.get(vid, {}).get("유형"),
            "소속": vendor.get(vid, {}).get("소속"),
            "RS율": w.get("원작사RS율"),
            "항목분류": w.get("항목분류"),
            "상태": str(w.get("상태") or "").strip(),
            "증빙_국내": w.get("증빙구분_국내"),
            "증빙_해외": w.get("증빙구분_해외"),
            "통화": w.get("통화"),
        }
        keys = [title] + [a for a in str(w.get("Alias명(여러개;구분)") or "").split(";") if a.strip()]
        for k in keys:
            nk = _norm(k)
            if not nk:
                continue
            index.setdefault(nk, [])
            if rec not in index[nk]:
                index[nk].append(rec)
            if len(index[nk]) > 1:
                dup.append((k, [r["표준작품명"] for r in index[nk]]))
    return index, vendor, dup


def _is_terminated(status):
    return any(k in status for k in ("해지", "종료"))


def _load_orphan_pending(master_path):
    """ORPHAN_등록대기 시트의 작품(이미 식별·보류된 건) 정규화키 집합."""
    try:
        _, rows = M.read_sheet(master_path, "ORPHAN_등록대기")
    except Exception:
        return set()
    keys = set()
    for r in rows:
        t = r.get("표준작품명")
        if t:
            keys.add(_norm(t))
    return keys


# ── 매칭 ────────────────────────────────────────────────────
def match_sales(sales_records, master_path, company):
    """company 매출(항목=원작료) 작품을 마스터에 매칭 → 분류 버킷 + 리포트.
       판정은 작품Alias마스터 단독. 매출 '웹소설본부' 표기는 의미 없으므로 무시."""
    index, vendor, dupkeys = build_master_index(master_path)
    orphan_pending = _load_orphan_pending(master_path)

    # 항목=원작료 행만 → 작품명 단위로 묶기 (실제 원작사명 힌트만 따로 보관)
    works = {}
    for r in sales_records:
        if r["회사"] != company:
            continue
        if "원작료" not in str(r.get("항목") or ""):
            continue
        title = r.get("작품명")
        if not title:
            continue
        hint = str(r.get("웹소설본부") or "").strip()
        w = works.setdefault(title, {"작품명": title, "행수": 0,
                                     "원작사힌트": set(), "순매출합": 0.0})
        w["행수"] += 1
        # '웹소설본부'·빈값은 의미 없는 표시 → 참고 힌트에서 제외, 실제 원작사명만 보관
        if hint and hint.lower() not in _BLANK_HINT and hint != _SELF:
            w["원작사힌트"].add(hint)
        w["순매출합"] += (r.get("정산순매출") or 0)

    지급대상, 제외, 신규orphan, 기존대기, mismatch = [], [], [], [], []
    for title, w in works.items():
        hints = sorted(w["원작사힌트"])
        hit = index.get(_norm(title))
        if hit:                                       # ── 마스터 매칭 성공
            m = hit[0]
            row = {**w, "원작사힌트": hints, "원작사_업체ID": m["원작사_업체ID"],
                   "원작사명": m["원작사명"], "회사": m["회사"], "유형": m["유형"],
                   "소속": m["소속"], "RS율": m["RS율"], "상태": m["상태"],
                   "마스터작품명": m["표준작품명"], "중복매칭": len(hit) > 1}
            (제외 if _is_terminated(m["상태"]) else 지급대상).append(row)
            # Mismatch: 실제 원작사명 힌트가 있고 마스터 원작사와 다를 때만
            if hints and m["원작사명"] and all(_norm(h) != _norm(m["원작사명"]) for h in hints):
                mismatch.append({"작품명": title, "매출_원작사힌트": hints,
                                 "마스터_원작사": m["원작사명"], "업체ID": m["원작사_업체ID"]})
        else:                                         # ── 미매칭
            row = {**w, "원작사힌트": hints}
            (기존대기 if _norm(title) in orphan_pending else 신규orphan).append(row)

    return {
        "회사": company,
        "지급대상": 지급대상,
        "제외_해지종료": 제외,
        "ORPHAN_신규": 신규orphan,
        "ORPHAN_기존대기": 기존대기,
        "Mismatch": mismatch,
        "마스터중복키": dupkeys,
        "요약": {"원작료작품수": len(works),
                "지급대상": len(지급대상), "제외(해지/종료)": len(제외),
                "신규ORPHAN(검토필요)": len(신규orphan),
                "기존대기(이미보류)": len(기존대기), "Mismatch": len(mismatch)},
    }
