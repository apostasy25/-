# -*- coding: utf-8 -*-
"""마스터·템플릿·첨부 폴더의 고정 경로를 한 곳에서 결정(런처·정산·마스터 공용).

프로그램 폴더 기준:
  · 빌드(EXE)  → 실행 파일이 있는 폴더
  · 소스 실행  → 이 파일이 있는 프로젝트 루트
환경변수(MASTER_PATH / TEMPLATE_PATH)로 개별 덮어쓰기 가능.
"""
import os
import sys


def program_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def master_path():
    return os.environ.get(
        "MASTER_PATH", os.path.join(program_dir(), "원작료정산_마스터.xlsx"))


def template_path():
    return os.environ.get(
        "TEMPLATE_PATH", os.path.join(program_dir(), "정산서_양식_기초.xlsx"))


def biz_cert_dir():
    """사업자등록증 보관 폴더(없으면 생성)."""
    d = os.path.join(program_dir(), "사업자등록증")
    os.makedirs(d, exist_ok=True)
    return d
