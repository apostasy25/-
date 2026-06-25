# -*- coding: utf-8 -*-
"""
원작료 정산 파이프라인 GUI (검수 게이트 방식 a) — tkinter.

흐름:
  RAW 폴더 선택 → 정산 실행 → 정산서 생성 → [검수 리포트 팝업·게이트1] → 확인
  → PDF 생성 → [최종확인 게이트2] → 확인 → PDF 비밀번호 설정 → 메일 Draft 생성

오류 건수 > 0 이면 게이트1에서 강한 경고를 띄우고, 사용자가 명시적으로 강행하지
않는 한 PDF·메일로 진행하지 않는다(대량 오발송 방지). 메일은 .eml 초안만 생성(무발송).
"""
import os
import sys
import glob
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_HERE)
sys.path.insert(0, os.path.join(_ROOT, "engine"))
sys.path.insert(0, _ROOT)

import pipeline_v2 as P                       # noqa: E402
import paths as _paths                         # noqa: E402

try:                                          # 드래그&드롭(있으면 사용, 없으면 파일선택만)
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _DND = True
except Exception:                             # noqa: BLE001
    TkinterDnD = None
    DND_FILES = None
    _DND = False


def _find(folder, *patterns):
    for pat in patterns:
        hits = glob.glob(os.path.join(folder, "**", pat), recursive=True)
        if hits:
            return hits[0]
    return None


class App:
    def __init__(self, root):
        self.root = root
        self.report = None
        self.sales_path = tk.StringVar()
        self.etc_path = tk.StringVar()
        self.master = _paths.master_path()        # 고정 경로(첨부 불필요)
        self.company = tk.StringVar(value="테라핀")
        self.month = tk.StringVar()
        self.out_root = tk.StringVar(value=os.path.join(_paths.program_dir(), "output"))
        root.title("원작료 정산 — 검수형 파이프라인")
        root.geometry("820x660")
        self._build()

    def _build(self):
        pad = {"padx": 8, "pady": 4}
        hint = "파일 선택 버튼을 누르거나, 파일을 칸에 끌어다 놓으세요."
        if not _DND:
            hint += "  (끌어다 놓기를 쓰려면 tkinterdnd2 설치 필요 — 파일 선택은 항상 가능)"
        frm = ttk.LabelFrame(self.root, text="1. 입력 파일")
        frm.pack(fill="x", **pad)
        ttk.Label(frm, text=hint, foreground="#666").grid(
            row=0, column=0, columnspan=3, sticky="w", padx=6, pady=(4, 8))
        self._file_row(frm, 1, "매출리스트 파일 *", self.sales_path, self.pick_sales,
                       "플랫폼 순매출 RAW (.xlsx/.xls) — 예: …테라핀…원작료…정산서가 아니라 '매출리스트' 파일")
        self._file_row(frm, 2, "기타수익 파일", self.etc_path, self.pick_etc,
                       "기타수익_누적 (.xlsx/.xls) — 해당 없으면 비워 두세요")
        # 마스터는 프로그램 폴더 고정 경로(첨부 불필요) — 읽기 전용 안내
        ttk.Label(frm, text="마스터 파일", width=14, anchor="w").grid(
            row=6, column=0, sticky="w", padx=6, pady=(8, 0))
        _ok = os.path.exists(self.master)
        ttk.Label(frm, text=("자동 사용: " + self.master) if _ok else
                  ("⚠ 없음 — 프로그램 폴더에 원작료정산_마스터.xlsx 필요: " + self.master),
                  foreground=("#0a7" if _ok else "#c00"),
                  font=("맑은 고딕", 9)).grid(row=6, column=1, columnspan=2, sticky="w", padx=4, pady=(8, 0))
        ttk.Label(frm, text="(마스터 편집은 시작 화면의 '마스터 관리'에서 — 같은 파일을 공유합니다)",
                  foreground="#888", font=("맑은 고딕", 8)).grid(
            row=7, column=1, sticky="w", padx=4, pady=(0, 4))

        opt = ttk.Frame(self.root)
        opt.pack(fill="x", **pad)
        ttk.Label(opt, text="회사").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Combobox(opt, textvariable=self.company, values=["테라핀", "수성웹툰"],
                     width=14, state="readonly").grid(row=0, column=1, sticky="w", padx=4)
        ttk.Label(opt, text="정산서월 (YYYY-MM)").grid(row=0, column=2, sticky="w", padx=16, pady=4)
        ttk.Entry(opt, textvariable=self.month, width=14).grid(row=0, column=3, sticky="w", padx=4)

        btnf = ttk.Frame(self.root)
        btnf.pack(fill="x", **pad)
        self.run_btn = ttk.Button(btnf, text="▶ 정산 실행", command=self.on_run)
        self.run_btn.pack(side="left", padx=6)
        ttk.Button(btnf, text="결과 폴더 열기", command=self.open_out).pack(side="left", padx=6)

        logf = ttk.LabelFrame(self.root, text="진행 로그")
        logf.pack(fill="both", expand=True, **pad)
        self.log = scrolledtext.ScrolledText(logf, height=18, wrap="word")
        self.log.pack(fill="both", expand=True, padx=4, pady=4)

    def _file_row(self, frm, row, label, var, picker, desc):
        ttk.Label(frm, text=label, width=14, anchor="w").grid(
            row=row * 2, column=0, sticky="w", padx=6, pady=(4, 0))
        ent = ttk.Entry(frm, textvariable=var, width=72)
        ent.grid(row=row * 2, column=1, padx=4, pady=(4, 0))
        ttk.Button(frm, text="파일 선택", command=picker).grid(
            row=row * 2, column=2, padx=4, pady=(4, 0))
        ttk.Label(frm, text=desc, foreground="#888", font=("맑은 고딕", 8)).grid(
            row=row * 2 + 1, column=1, sticky="w", padx=4, pady=(0, 4))
        self._enable_dnd(ent, var)

    def _enable_dnd(self, entry, var):
        """tkinterdnd2 가 있으면 해당 입력칸에 파일 드롭을 활성화."""
        if not _DND:
            return
        try:
            entry.drop_target_register(DND_FILES)
            entry.dnd_bind("<<Drop>>",
                           lambda e, v=var: v.set(e.data.strip().strip("{}")))
        except Exception:                        # noqa: BLE001
            pass

    _XLS = [("Excel 파일", "*.xlsx *.xls"), ("모든 파일", "*.*")]

    def pick_sales(self):
        _ib = os.path.join(_paths.program_dir(), "inbox", "매출리스트")
        f = filedialog.askopenfilename(title="매출리스트 파일 선택", filetypes=self._XLS,
                                       initialdir=_ib if os.path.isdir(_ib) else _paths.program_dir())
        if f:
            self.sales_path.set(f)

    def pick_etc(self):
        _ib = os.path.join(_paths.program_dir(), "inbox", "기타수익")
        f = filedialog.askopenfilename(title="기타수익 파일 선택", filetypes=self._XLS,
                                       initialdir=_ib if os.path.isdir(_ib) else _paths.program_dir())
        if f:
            self.etc_path.set(f)


    def open_out(self):
        d = os.path.join(self.out_root.get(), self.month.get(), self.company.get())
        if os.path.isdir(d) and sys.platform.startswith("win"):
            try:
                os.startfile(d)                  # noqa
                return
            except Exception:
                pass
        self._log(f"결과 폴더: {d}")

    def _log(self, msg):
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.root.update_idletasks()

    def _busy(self, on):
        self.run_btn.config(state="disabled" if on else "normal")

    def _dispatch(self, fn, callback, errtitle):
        """장시간 작업을 백그라운드 스레드로 실행하고 결과를 메인 스레드 콜백으로 전달."""
        def work():
            try:
                result = fn()
                self.root.after(0, lambda r=result: callback(r))
            except Exception:                    # noqa: BLE001
                err = traceback.format_exc()
                self.root.after(0, lambda e=err: self._fail(errtitle, e))
        threading.Thread(target=work, daemon=True).start()

    # ── STAGE 1: 정산 실행 → 검수 리포트(게이트1) ─────────
    def on_run(self):
        sales = self.sales_path.get().strip()
        etc = self.etc_path.get().strip()
        master = self.master
        month = self.month.get().strip()
        if not (sales and master and month):
            messagebox.showwarning("입력 필요",
                                   "매출리스트 파일·마스터 파일·정산서월은 필수입니다.")
            return
        if not os.path.exists(sales):
            messagebox.showerror("파일 없음", f"매출리스트 파일을 찾을 수 없습니다:\n{sales}")
            return
        if not os.path.exists(master):
            messagebox.showerror("파일 없음", f"마스터 파일을 찾을 수 없습니다:\n{master}")
            return
        if etc and not os.path.exists(etc):
            messagebox.showerror("파일 없음", f"기타수익 파일을 찾을 수 없습니다:\n{etc}")
            return
        self._busy(True)
        company = self.company.get()
        groups = ("수성",) if company == "수성웹툰" else \
                 ("네이버", "그외", "개인", "해외", "리디분기")
        self._log("\n=== 정산 실행 시작 ===")
        self._log(f"  매출리스트: {os.path.basename(sales)}")
        self._log(f"  기타수익  : {os.path.basename(etc) if etc else '(없음)'}")
        self._log(f"  마스터    : {os.path.basename(master)}")
        self._log(f"  대상 그룹 : {', '.join(groups)}")
        self._dispatch(
            lambda: P.settle(master, sales, etc or "", company, month,
                             out_root=self.out_root.get(), groups=groups, recalc=True),
            self._after_settle, "정산 실행 오류")

    def _after_settle(self, rep):
        self.report = rep
        self._busy(False)
        self._log("정산서 생성 완료. 검수 리포트를 확인하세요.")
        self.show_review_gate(rep)

    def show_review_gate(self, rep):
        """게이트1 — 검수 리포트 팝업. 확인 시에만 PDF로 진행."""
        win = tk.Toplevel(self.root)
        win.title("검수 리포트 (PDF 생성 전 확인)")
        win.geometry("440x460")
        win.transient(self.root)
        win.grab_set()
        오류 = rep.get("오류건수", 0)
        경고 = len(rep.get("경고", []))
        items = [
            ("정산 대상 업체 수", rep.get("업체수", 0)),
            ("생성 정산서 수", len(rep.get("files", []))),
            ("MG 작품 수", rep.get("MG작품수", 0)),
            ("해외 정산 수", rep.get("해외정산수", 0)),
            ("분기/반기 정산 수", rep.get("분기반기수", 0)),
            ("이월 업체 수", rep.get("이월대상수", 0)),
            ("오류 건수", 오류),
            ("경고 건수", 경고),
        ]
        ttk.Label(win, text=f"[{rep.get('company')}] {rep.get('정산서월')} 검수 리포트",
                  font=("맑은 고딕", 12, "bold")).pack(pady=10)
        body = ttk.Frame(win)
        body.pack(fill="x", padx=24)
        for i, (k, v) in enumerate(items):
            ttk.Label(body, text=k, width=20, anchor="w").grid(row=i, column=0, sticky="w", pady=2)
            color = "red" if (k == "오류 건수" and v) else ("#b36b00" if (k == "경고 건수" and v) else "black")
            tk.Label(body, text=str(v), anchor="e", width=10, fg=color,
                     font=("맑은 고딕", 10, "bold")).grid(row=i, column=1, sticky="e", pady=2)
        if rep.get("경고"):
            wf = ttk.LabelFrame(win, text="경고 내역")
            wf.pack(fill="both", expand=True, padx=16, pady=6)
            t = scrolledtext.ScrolledText(wf, height=5, wrap="word")
            t.pack(fill="both", expand=True)
            t.insert("end", "\n".join("· " + w for w in rep["경고"]))
            t.config(state="disabled")

        btnf = ttk.Frame(win)
        btnf.pack(pady=10)

        def proceed():
            if 오류:
                if not messagebox.askyesno(
                        "오류 있음 — 강행?",
                        f"수식 오류 {오류}건이 있습니다. 그래도 PDF 생성을 진행할까요?\n"
                        "(권장: 취소 후 오류를 먼저 해결)", icon="warning", default="no", parent=win):
                    return
            win.destroy()
            self.run_pdf()

        ttk.Button(btnf, text="확인 — PDF 생성", command=proceed).pack(side="left", padx=8)
        ttk.Button(btnf, text="취소", command=win.destroy).pack(side="left", padx=8)

    # ── STAGE 2: PDF 생성 → 최종확인(게이트2) ─────────────
    def run_pdf(self):
        self._busy(True)
        self._log("\n=== PDF 생성 시작 ===")
        self._dispatch(lambda: P.to_pdf(self.report, out_root=self.out_root.get()),
                       self._after_pdf, "PDF 생성 오류")

    def _after_pdf(self, pdfs):
        self._busy(False)
        n = sum(1 for p in pdfs if p)
        self._log(f"PDF 생성 완료: {n}건")
        go = messagebox.askyesno(
            "최종 확인 — 메일 Draft 생성",
            f"PDF {n}건이 생성되었습니다.\n\n"
            "이제 PDF 비밀번호 설정 + 메일 Draft(.eml)를 생성합니다.\n"
            "메일은 '초안(.eml)'만 만들며 자동 발송되지 않습니다.\n\n"
            "진행할까요?", icon="question", default="no")
        if not go:
            self._log("메일 Draft 생성 취소(게이트2). PDF까지만 완료.")
            return
        self.run_mail()

    # ── STAGE 3: 비밀번호 + 메일 Draft ───────────────────
    def run_mail(self):
        self._busy(True)
        self._log("\n=== 비밀번호 설정 + 메일 Draft 생성 ===")
        self.report["승인_PDF"] = True            # 게이트2 통과 표시
        self._dispatch(
            lambda: P.secure_and_mail(self.report, self.master,
                                      out_root=self.out_root.get()),
            self._after_mail, "메일 Draft 오류")

    def _after_mail(self, res):
        self._busy(False)
        eml = sum(1 for *_, e in res if e)
        nopw = [n for n, pw, _e in res if not pw]
        self._log(f"메일 Draft 생성 완료: {eml}건")
        if nopw:
            self._log(f"  ⚠ 비밀번호 미설정(증빙번호 미등록): {', '.join(nopw)}")
        messagebox.showinfo("완료",
                            f"완료되었습니다.\n메일 Draft(.eml) {eml}건 생성.\n"
                            "메일은 자동 발송되지 않았습니다(초안만).")

    def _fail(self, title, err):
        self._busy(False)
        self._log("‼ " + title + "\n" + err)
        messagebox.showerror(title, (err.strip().splitlines() or [title])[-1])


def main():
    root = TkinterDnD.Tk() if _DND else tk.Tk()    # DnD 가능 시 DnD 루트
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
