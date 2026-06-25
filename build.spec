# -*- mode: python ; coding: utf-8 -*-
# PyInstaller onedir 빌드 사양 — 정산자동화 (단일 진입점 런처)
#   · 진입점: launcher.py  → 시작 화면에서 '정산 실행' / '마스터 관리' 선택
#   · 런처가 선택 도구를 같은 EXE의 별도 프로세스(--mode)로 재실행 → Qt+Tk 프로세스 격리
#   · onedir: 폴더 통째 배포(빠른 시작, OneDrive 안전). onefile 아님.
#   · 마스터/양식 xlsx 는 번들하지 않고 exe 옆에 둠(편집 가능한 영구 DB). MASTER_PATH 로도 지정 가능.
from PyInstaller.utils.hooks import collect_all

# PySide6(Qt 플러그인 포함) 전체 수집
_qt_datas, _qt_bins, _qt_hidden = collect_all('PySide6')

# (선택) tkinterdnd2 — drag&drop. 미설치 시 빈 값(파일 선택만 동작)
try:
    _dnd_datas, _dnd_bins, _dnd_hidden = collect_all('tkinterdnd2')
except Exception:
    _dnd_datas, _dnd_bins, _dnd_hidden = [], [], []

# 동적 import 되는 로컬 모듈(런처가 import) — 명시적으로 번들
_local = [
    'pipeline_app', 'app', 'master_io',                       # GUI
    'pipeline_v2', 'securepdf', 'maildraft', 'daterules',     # 파이프라인
    'm0_loader', 'm1_matcher', 'm2_calc', 'ledger',           # 엔진
    'm3_builder', 'm3_sheet', 'm3_run', 'config',
    'template_fill', 'paths',                                 # 함수 내 동적 import(누락 방지)
]

a = Analysis(
    ['launcher.py'],
    pathex=['.', 'engine', 'gui', 'gui/master'],
    binaries=_qt_bins + _dnd_bins,
    datas=_qt_datas + _dnd_datas,
    hiddenimports=['openpyxl', 'pypdf', 'tkinter'] + _local + _qt_hidden + _dnd_hidden,
    excludes=['matplotlib', 'numpy', 'pandas', 'PIL', 'scipy'],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name='정산자동화',
    console=False,                # GUI(콘솔 창 숨김)
)
coll = COLLECT(
    exe, a.binaries, a.datas,
    name='정산자동화',            # 결과: dist/정산자동화/정산자동화.exe
)
