# -*- coding: utf-8 -*-
"""
정산자동화.exe 단일 진입점(런처).

사용자는 이 하나의 프로그램만 실행한다. 시작 화면(런처)에서
  · 정산 실행   → pipeline_app (tkinter)
  · 마스터 관리 → gui/app      (PySide6)
중 하나를 선택한다.

[안정성 설계] PySide6(Qt)와 tkinter를 한 프로세스에서 동시에 띄우면 이벤트 루프가
충돌한다. 그래서 런처는 선택한 도구를 **같은 실행 파일의 별도 프로세스**로 재실행한다
(`정산자동화.exe --mode=pipeline|master`). 결과적으로
  · 런처 프로세스      → tkinter 만 로드(가벼운 선택 창)
  · 정산 실행 프로세스 → tkinter 만 로드
  · 마스터 관리 프로세스→ PySide6 만 로드
어떤 프로세스도 두 프레임워크를 동시에 로드하지 않는다. 패키징은 PyInstaller 가
두 프레임워크를 모두 번들하되, 각 프로세스는 런타임에 자기 것만 import 한다.
"""
import os
import sys
import subprocess


def _bundle_dir():
    """소스/번들 모듈을 찾을 기준 경로."""
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _setup_path():
    base = _bundle_dir()
    for sub in ("", "engine", "gui", os.path.join("gui", "master")):
        p = os.path.join(base, sub) if sub else base
        if os.path.isdir(p) and p not in sys.path:
            sys.path.insert(0, p)


def _self_cmd(mode):
    """자기 자신을 --mode 로 재실행하는 커맨드."""
    if getattr(sys, "frozen", False):
        return [sys.executable, f"--mode={mode}"]
    return [sys.executable, os.path.abspath(__file__), f"--mode={mode}"]


def run_pipeline():
    """정산 실행(tkinter) — 이 프로세스에서 실행."""
    _setup_path()
    import pipeline_app
    pipeline_app.main()


def run_master():
    """마스터 관리(PySide6) — 이 프로세스에서 실행."""
    _setup_path()
    import app as master_app          # gui/app.py
    master_app.main()


def show_launcher():
    """시작 화면(tkinter). 선택 도구를 별도 프로세스로 재실행."""
    import tkinter as tk
    from tkinter import ttk, messagebox

    root = tk.Tk()
    root.title("원작료 정산 자동화")
    root.geometry("380x280")
    root.resizable(False, False)

    ttk.Label(root, text="원작료 정산 자동화",
              font=("맑은 고딕", 16, "bold")).pack(pady=(28, 4))
    ttk.Label(root, text="실행할 작업을 선택하세요",
              font=("맑은 고딕", 10)).pack(pady=(0, 20))

    def launch(mode, label):
        try:
            subprocess.Popen(_self_cmd(mode))
        except Exception as e:           # noqa: BLE001
            messagebox.showerror("실행 오류", f"{label} 실행 실패:\n{e}")

    ttk.Button(root, text="\U0001F4C4   정산 실행", width=26,
               command=lambda: launch("pipeline", "정산 실행")).pack(pady=7)
    ttk.Button(root, text="\U0001F5C2   마스터 관리", width=26,
               command=lambda: launch("master", "마스터 관리")).pack(pady=7)

    tk.Label(root, text="두 작업을 모두 열 수 있습니다. 이 창은 닫아도 됩니다.",
             fg="#666", font=("맑은 고딕", 9)).pack(side="bottom", pady=14)
    root.mainloop()


def main():
    mode = None
    for a in sys.argv[1:]:
        if a.startswith("--mode="):
            mode = a.split("=", 1)[1]
    if mode == "pipeline":
        run_pipeline()
    elif mode == "master":
        run_master()
    else:
        show_launcher()


if __name__ == "__main__":
    main()
