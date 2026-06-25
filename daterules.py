# -*- coding: utf-8 -*-
"""
날짜 규칙 — 작성일자 / (세금)계산서 발행 마감일 / 지급일

확정 규칙 (실거래 검증):
 ① 네이버   : 작성=대상월 말일,  마감=익월 8일,   지급=익월 15일
 ② 리디(분기): 작성=분기말 익월25, 마감=분기말 익월말일, 지급=분기말 익익월10
 ③ 기타 사업자: 작성=익월 5일,    마감=익월 10일,  지급=익월 15일
 ④ 개인      : 지급=익월 15일 (작성/마감은 원본 양식 기준 — 개인 양식엔 날짜칸 없음)
 ⑤ 카카오    : 작성/마감/지급 모두 "-" (카카오 직접 정산 구조)

모든 날짜는 YYYY-MM-DD. 미국식(mm/dd/yy) 금지.
"""
import datetime, calendar


def _last(y, mo):
    return datetime.datetime(y, mo, calendar.monthrange(y, mo)[1])


def _d(y, mo, d):
    ny, nm = y, mo
    while nm > 12:
        nm -= 12
        ny += 1
    return datetime.datetime(ny, nm, d)


def _nextmo(y, mo, d):
    ny, nm = (y, mo + 1) if mo < 12 else (y + 1, 1)
    return datetime.datetime(ny, nm, d)


def _lastnextmo(y, mo):
    ny, nm = (y, mo + 1) if mo < 12 else (y + 1, 1)
    return _last(ny, nm)


def settle_dates(vtype, y, mo):
    """월간 정산 날짜. vtype: naver | biz | personal | kakao."""
    if vtype == "kakao":
        return dict(작성=None, 마감=None, 지급=None, dash=True)
    if vtype == "naver":
        return dict(작성=_last(y, mo), 마감=_nextmo(y, mo, 8), 지급=_nextmo(y, mo, 15))
    if vtype == "personal":
        return dict(작성=None, 마감=None, 지급=_nextmo(y, mo, 15))  # 작성/마감 원본 유지
    return dict(작성=_nextmo(y, mo, 5), 마감=_nextmo(y, mo, 10), 지급=_nextmo(y, mo, 15))


def settle_dates_quarter(y, q):
    """리디 분기 정산. q=1..4.
       작성=분기말 익월25, 마감=분기말 익월말일, 지급=분기말 익익월10."""
    qend = q * 3
    return dict(작성=_d(y, qend + 1, 25),
                마감=_lastnextmo(y, qend),
                지급=_d(y, qend + 2, 10))


def vendor_type(name):
    """업체명으로 날짜규칙 유형 판정."""
    if "네이버" in name:
        return "naver"
    if any(k in name for k in ["카카오", "타파스", "픽코마"]):
        return "kakao"
    if name == "리디" or name.startswith("리디"):
        return "ridi"
    return "biz"
