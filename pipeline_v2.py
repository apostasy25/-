# -*- coding: utf-8 -*-
"""
원작료 정산 파이프라인 (검수 게이트 + 로그).

흐름:
  RAW 선택 → [settle] M0→M1→M2→M3 (자동) → 검수 리포트
           → (사용자 승인) → [to_pdf] 정산시트 PDF
           → (사용자 승인) → [secure_and_mail] PDF 비밀번호 + 메일 .eml 초안

PDF·메일 단계는 호출측(GUI)이 승인 후에만 호출한다. 무발송: .eml 까지만.

실행 로그(검수용 지표): 처리 작품 수·MG 작품 수·해외 정산 수·이월 대상 수·오류 건수.
"""
import os
import sys
import json
import datetime
import subprocess

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_HERE, "engine"))
sys.path.insert(0, _HERE)

import m3_run                              # noqa: E402
import securepdf                           # noqa: E402
import maildraft                           # noqa: E402

RECALC = os.environ.get("RECALC_PATH", "/mnt/skills/public/xlsx/scripts/recalc.py")


# ─────────────────────────────────────────────
# STAGE 1 — settle (자동): M0→M1→M2→M3 + 검수 리포트
# ─────────────────────────────────────────────
def settle(master, sales, etc, company, 정산서월, out_root="output", 이월_prior=None,
           groups=("네이버", "그외"), recalc=True):
    produced, report = m3_run.run_settlement(
        master, sales, etc, company, 정산서월, out_root, groups, 이월_prior=이월_prior)
    # 수식 오류 점검(LibreOffice 재계산)
    errs, detail = 0, {}
    if recalc and os.path.exists(RECALC):
        for f in report["files"]:
            e, summ = _recalc_errors(f["path"])
            if e and e > 0:
                errs += e
                detail[os.path.basename(f["path"])] = summ
    report["오류건수"] = errs
    report["오류상세"] = detail
    report["로그파일"] = _write_log(report, out_root, 정산서월, company)
    report["승인_정산"] = True              # 자동 통과
    report["승인_PDF"] = False
    report["승인_메일"] = False
    return report


def _recalc_errors(xlsx):
    try:
        r = subprocess.run([sys.executable, RECALC, xlsx, "90"],
                           capture_output=True, text=True, timeout=160)
        d = json.loads(r.stdout)
        return d.get("total_errors", 0), d.get("error_summary", {})
    except Exception as e:                  # noqa: BLE001
        return -1, {"예외": str(e)}


def _write_log(report, out_root, 정산서월, company):
    d = os.path.join(out_root, 정산서월, company)
    os.makedirs(d, exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    L = [
        f"========== 정산 실행 로그 ==========",
        f"회사       : {company}",
        f"정산서월   : {정산서월}",
        f"실행시각   : {ts}",
        f"------------------------------------",
        f"처리 업체 수 : {report['업체수']}",
        f"처리 작품 수 : {report['처리작품수']}",
        f"MG 작품 수   : {report['MG작품수']}",
        f"해외 정산 수 : {report['해외정산수']}",
        f"이월 대상 수 : {report['이월대상수']}",
        f"오류 건수    : {report['오류건수']}",
        f"------------------------------------",
        "생성 파일:",
    ]
    L += [f"  - [{f['group']}] {f['업체']} : {f['path']}" for f in report["files"]]
    if report.get("미구현그룹"):
        L += ["", "전용양식 대기(미생성):"] + [f"  - {g}" for g in report["미구현그룹"]]
    if report.get("경고"):
        L += ["", "경고:"] + [f"  - {w}" for w in report["경고"]]
    if report.get("오류상세"):
        L += ["", "오류상세:"] + [f"  - {k}: {v}" for k, v in report["오류상세"].items()]
    path = os.path.join(d, f"실행로그_{datetime.datetime.now():%Y%m%d_%H%M%S}.txt")
    with open(path, "w", encoding="utf-8") as fp:
        fp.write("\n".join(L))
    return path


def format_review(report):
    """검수 화면용 요약 문자열."""
    return (
        f"[검수 — {report['company']} {report['정산서월']}]\n"
        f"  처리 업체 {report['업체수']} · 작품 {report['처리작품수']} · "
        f"MG {report['MG작품수']} · 해외 {report['해외정산수']} · "
        f"이월 {report['이월대상수']} · 오류 {report['오류건수']}\n"
        + ("\n".join(f"  - [{f['group']}] {f['업체']}" for f in report["files"]))
    )


# ─────────────────────────────────────────────
# STAGE 2 — to_pdf (정산 승인 후): 정산시트 PDF (비밀번호 없음, 검토용)
# ─────────────────────────────────────────────
def to_pdf(report, out_root="output"):
    if not report.get("승인_정산"):
        raise RuntimeError("정산 검수 미승인 — PDF 생성 불가")
    base = os.path.join(out_root, report["정산서월"], report["company"])
    pdfdir = os.path.join(base, "PDF")
    tmp = os.path.join(base, "_tmp")
    os.makedirs(pdfdir, exist_ok=True)
    os.makedirs(tmp, exist_ok=True)
    for f in report["files"]:
        try:
            f["pdf"] = securepdf.settlement_only_pdf(f["path"], tmp, pdfdir)
        except Exception as e:              # noqa: BLE001
            f["pdf_error"] = str(e)
    report["승인_PDF_대기"] = True
    return [f.get("pdf") for f in report["files"] if f.get("pdf")]


# ─────────────────────────────────────────────
# STAGE 3 — secure_and_mail (PDF 승인 후): 비밀번호 + 메일 .eml
# ─────────────────────────────────────────────
def secure_and_mail(report, master, out_root="output"):
    if not report.get("승인_PDF"):
        raise RuntimeError("PDF 검수 미승인 — 암호화·메일 생성 불가")
    pwmap = securepdf.build_pwmap(master)
    emaster, _nc2id = maildraft.load_email_master(master)
    vid2email = {}
    for (vid, _key), v in emaster.items():
        vid2email.setdefault(str(vid), v)
    base = os.path.join(out_root, report["정산서월"], report["company"])
    maildir = os.path.join(base, "메일Draft")
    os.makedirs(maildir, exist_ok=True)
    results = []
    for f in report["files"]:
        pdf = f.get("pdf")
        if not pdf:
            continue
        pw = pwmap.get(securepdf._nmf(str(f["업체"])), (None, None))[0]
        if pw:
            try:
                securepdf.encrypt(pdf, pw)
                f["암호화"] = True
            except Exception as e:          # noqa: BLE001
                f["암호화오류"] = str(e)
        else:
            f["경고_pw"] = "비밀번호 없음(사업자/주민번호 미등록)"
        em = vid2email.get(str(f["vid"]))
        if em:
            f["eml"] = _make_mail(report, f, pdf, em, maildir)
        else:
            f["경고_mail"] = "이메일마스터 미등록"
        results.append((f["업체"], bool(pw), f.get("eml")))
    return results


def _split_addr(s):
    return [x.strip() for x in str(s or "").replace(";", ",").split(",") if x.strip()]


def _make_mail(report, f, pdf, em, maildir):
    mo = report["정산서월"].split("-")[1]
    body = f"{report['company']} {mo}월 원작료 정산서를 첨부드립니다. (검토용 초안)"
    subject = f"[{report['company']}] {report['정산서월']} 원작료 정산서 — {f['업체']}"
    out = os.path.join(maildir, f"{m3_run._safe_name(f['업체'])}_{report['정산서월']}.eml")
    try:
        maildraft.make_eml(out, em.get("frm"),
                           _split_addr(em.get("to")), _split_addr(em.get("cc")),
                           _split_addr(em.get("bcc")), subject, body, "", [pdf])
        return out
    except Exception as e:                  # noqa: BLE001
        f["메일오류"] = str(e)
        return None


# ─────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--master", required=True)
    ap.add_argument("--sales", required=True)
    ap.add_argument("--etc", required=True)
    ap.add_argument("--company", default="테라핀")
    ap.add_argument("--month", required=True, help="정산서월 YYYY-MM")
    ap.add_argument("--out", default="output")
    a = ap.parse_args()
    rep = settle(a.master, a.sales, a.etc, a.company, a.month, a.out)
    print(format_review(rep))
    print(f"\n로그: {rep['로그파일']}")
    print("PDF·메일은 검수 승인 후 to_pdf()/secure_and_mail() 호출")
