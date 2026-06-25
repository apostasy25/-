# -*- coding: utf-8 -*-
"""
원작료 정산 자동화 — 마스터 관리 GUI (PySide6).

이 단계의 범위(토대):
  · 메인 창 + 좌측 마스터 네비게이션 + 우측 편집 영역
  · 동시편집 잠금 배너(편집 모드 / 읽기 전용) — master_io 연동
  · '업체 마스터' 탭을 완전 동작(불러오기·추가·삭제·검색·저장)으로 구현
  · 나머지 탭은 동일한 표 편집 컴포넌트(MasterTab)로 재사용
진입점: 추후 exe 빌드 시 이 파일(app.py)을 PyInstaller 진입점으로 지정.
"""
import os
import sys

# master_io 임포트 (gui/master/master_io.py)
_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_HERE, "master"))
sys.path.insert(0, os.path.dirname(_HERE))        # 프로젝트 루트(paths 등)
import master_io as mio

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QListWidget, QListWidgetItem, QStackedWidget, QTableWidget,
    QTableWidgetItem, QPushButton, QLineEdit, QLabel, QMessageBox,
    QAbstractItemView, QHeaderView,
    QDialog, QComboBox, QDoubleSpinBox, QDialogButtonBox, QFormLayout,
    QFileDialog,
)

# 마스터 파일은 프로그램 폴더 기준 고정 경로(정산 실행과 동일 파일 참조)
import paths as _paths
MASTER = _paths.master_path()

# 좌측 네비게이션: (표시명, 시트명 또는 None)
NAV = [
    ("업체 마스터", "업체마스터"),
    ("작품 마스터", "작품Alias마스터"),
    ("이메일 마스터", "이메일마스터"),
    ("환율 마스터", "환율마스터"),
    ("선인세 MG", "선인세MG마스터"),
    ("이월 마스터", "이월마스터"),
    ("ORPHAN 등록대기", "ORPHAN_등록대기"),
]


class MasterTab(QWidget):
    """시트 1개를 표로 불러와 편집·저장하는 범용 컴포넌트."""

    def __init__(self, sheet, readonly=False):
        super().__init__()
        self.sheet = sheet
        self.readonly = readonly
        self.headers = []
        v = QVBoxLayout(self)

        # 툴바
        self.bar = bar = QHBoxLayout()
        self.btn_add = QPushButton("＋ 행 추가")
        self.btn_del = QPushButton("🗑 선택 삭제")
        self.btn_save = QPushButton("💾 저장")
        self.search = QLineEdit()
        self.search.setPlaceholderText("검색…")
        self.search.textChanged.connect(self._filter)
        self.btn_add.clicked.connect(self._add_row)
        self.btn_del.clicked.connect(self._del_row)
        self.btn_save.clicked.connect(self._save)
        for w in (self.btn_add, self.btn_del, self.btn_save):
            bar.addWidget(w)
        if sheet == "업체마스터" and not readonly:    # #4 사업자등록증 첨부
            self.btn_cert = QPushButton("📎 사업자등록증 첨부")
            self.btn_cert.clicked.connect(self._upload_cert)
            bar.addWidget(self.btn_cert)
        bar.addStretch(1)
        bar.addWidget(self.search)
        v.addLayout(bar)

        # 표
        self.table = QTableWidget()
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        v.addWidget(self.table)

        # 상태줄
        self.status = QLabel("")
        self.status.setStyleSheet("color:#6b6b6b;font-size:11px")
        v.addWidget(self.status)

        if readonly:                       # 읽기 전용: 편집·버튼 잠금
            for b in (self.btn_add, self.btn_del, self.btn_save):
                b.setEnabled(False)
            self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        self.load()

    def load(self):
        try:
            self.headers, rows = mio.read_sheet(MASTER, self.sheet, drop_example=False)
        except Exception as e:
            self.headers, rows = [], []
            self.status.setText(f"불러오기 실패: {e}")
            return
        self.table.setColumnCount(len(self.headers))
        self.table.setHorizontalHeaderLabels(self.headers)
        self.table.setRowCount(len(rows))
        for r, row in enumerate(rows):
            for c, h in enumerate(self.headers):
                val = row.get(h)
                self.table.setItem(r, c, QTableWidgetItem("" if val is None else str(val)))
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.status.setText(f"총 {len(rows)}행")

    def _add_row(self):
        self.table.insertRow(self.table.rowCount())

    def _upload_cert(self):
        """선택 업체 행에 사업자등록증 파일 첨부 → 사업자등록증/ 폴더 복사 + 경로 기록."""
        r = self.table.currentRow()
        if r < 0:
            QMessageBox.information(self, "사업자등록증", "업체 행을 먼저 선택하세요.")
            return
        hidx = {h: i for i, h in enumerate(self.headers)}
        pathcol = hidx.get("사업자등록증_경로")
        if pathcol is None:
            QMessageBox.information(self, "사업자등록증", "사업자등록증_경로 열이 없습니다.")
            return
        fn, _ = QFileDialog.getOpenFileName(
            self, "사업자등록증 선택", "", "문서/이미지 (*.pdf *.png *.jpg *.jpeg)")
        if not fn:
            return
        try:
            import shutil
            idcol, namecol = hidx.get("업체ID", 0), hidx.get("업체명", 1)
            vid = (self.table.item(r, idcol).text() if self.table.item(r, idcol) else "").strip()
            name = (self.table.item(r, namecol).text() if self.table.item(r, namecol) else "").strip()
            for ch in r'[]:*?/\\<>|"':
                name = name.replace(ch, "")
            ext = os.path.splitext(fn)[1].lower()
            dest = os.path.join(_paths.biz_cert_dir(), f"{vid}_{name}{ext}".strip("_"))
            shutil.copy2(fn, dest)
            self.table.setItem(r, pathcol, QTableWidgetItem(dest))
            self.status.setText(
                f"사업자등록증 첨부: {os.path.basename(dest)} — 💾 저장으로 경로를 마스터에 반영하세요")
        except Exception as e:
            QMessageBox.warning(self, "첨부 실패", str(e))

    def _del_row(self):
        for idx in sorted({i.row() for i in self.table.selectedIndexes()}, reverse=True):
            self.table.removeRow(idx)

    def _filter(self, text):
        text = text.strip().lower()
        for r in range(self.table.rowCount()):
            hit = not text or any(
                (self.table.item(r, c) and text in self.table.item(r, c).text().lower())
                for c in range(self.table.columnCount()))
            self.table.setRowHidden(r, not hit)

    def _collect(self):
        rows = []
        for r in range(self.table.rowCount()):
            row = {}
            empty = True
            for c, h in enumerate(self.headers):
                it = self.table.item(r, c)
                txt = it.text() if it else ""
                row[h] = txt if txt != "" else None
                if txt != "":
                    empty = False
            if not empty:
                rows.append(row)
        return rows

    def _save(self):
        rows = self._collect()
        errors, warnings = validate(self.sheet, self.headers, rows)
        if errors:
            QMessageBox.warning(self, "검증 오류 — 저장 취소", "\n".join(errors[:15]))
            return
        if warnings:
            ret = QMessageBox.question(self, "경고 (저장은 가능)",
                                       "\n".join(warnings[:15]) + "\n\n그래도 저장할까요?")
            if ret != QMessageBox.StandardButton.Yes:
                return
        try:
            n = mio.save_sheet(MASTER, self.sheet, self.headers, rows)
            self.status.setText(f"저장 완료: {n}행 (자동 백업됨)")
        except Exception as e:
            QMessageBox.warning(self, "저장 실패", str(e))


class Placeholder(QWidget):
    def __init__(self, name):
        super().__init__()
        v = QVBoxLayout(self)
        lab = QLabel(f"［{name}］ 준비 중")
        lab.setAlignment(Qt.AlignCenter)
        lab.setStyleSheet("color:#9b9b9b;font-size:14px")
        v.addWidget(lab)


def _vendor_ids():
    """업체마스터의 업체ID → 업체명 (참조 무결성 검증·편입 선택용)."""
    try:
        _, rows = mio.read_sheet(MASTER, "업체마스터")
    except Exception:
        return {}
    return {str(r.get("업체ID")).strip(): r.get("업체명")
            for r in rows if r.get("업체ID")}


def validate(sheet, headers, rows):
    """시트별 입력 검증. (errors, warnings). errors가 있으면 저장 차단."""
    errors, warnings = [], []
    g = lambda r, k: str(r.get(k) or "").strip()
    if sheet == "업체마스터":
        seen = set()
        for i, r in enumerate(rows, 1):
            vid = g(r, "업체ID")
            if not vid:
                errors.append(f"{i}행: 업체ID 비어있음")
            elif vid in seen:
                errors.append(f"{i}행: 업체ID 중복({vid})")
            else:
                seen.add(vid)
            t = g(r, "유형")
            if t and t not in ("사업자", "개인"):
                errors.append(f"{i}행: 유형은 사업자/개인 ({t})")
            if t == "사업자" and not g(r, "사업자번호"):
                warnings.append(f"{i}행({vid}): 사업자번호 비어있음")
            if t == "개인" and not g(r, "주민번호앞6자리(개인)"):
                warnings.append(f"{i}행({vid}): 주민번호 앞6자리 비어있음")
    elif sheet == "작품Alias마스터":
        ids = set(_vendor_ids())
        for i, r in enumerate(rows, 1):
            if not g(r, "표준작품명"):
                errors.append(f"{i}행: 표준작품명 비어있음")
            vid = g(r, "원작사_업체ID")
            if not vid:
                errors.append(f"{i}행: 원작사_업체ID 비어있음")
            elif vid not in ids:                       # 참조 무결성
                errors.append(f"{i}행: 원작사_업체ID '{vid}' 가 업체마스터에 없음")
            rs = g(r, "원작사RS율")
            if rs:
                try:
                    f = float(rs)
                    if not (0 < f <= 1):
                        warnings.append(f"{i}행: RS율 {f} 범위 의심(0~1)")
                except ValueError:
                    errors.append(f"{i}행: RS율 숫자 아님({rs})")
            st = g(r, "상태")
            if st and not any(k in st for k in ("정산중", "해지", "종료")):
                warnings.append(f"{i}행: 상태값 확인({st})")
    return errors, warnings


class PromoteDialog(QDialog):
    """ORPHAN → 작품마스터 편입: 원작사·RS·상태 선택."""

    def __init__(self, title, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"마스터 편입 — {title}")
        form = QFormLayout(self)
        self.cmb = QComboBox()
        for vid, name in sorted(_vendor_ids().items()):
            self.cmb.addItem(f"{vid} — {name}", vid)
        self.rs = QDoubleSpinBox()
        self.rs.setRange(0, 1)
        self.rs.setSingleStep(0.01)
        self.rs.setDecimals(2)
        self.rs.setValue(0.10)
        self.st = QComboBox()
        self.st.addItems(["정산중", "해지", "종료"])
        form.addRow("원작사", self.cmb)
        form.addRow("RS율", self.rs)
        form.addRow("상태", self.st)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        form.addRow(bb)

    def values(self):
        return self.cmb.currentData(), round(self.rs.value(), 2), self.st.currentText()


class OrphanTab(MasterTab):
    """ORPHAN_등록대기 + '마스터로 편입' 버튼."""

    def __init__(self, readonly=False):
        super().__init__("ORPHAN_등록대기", readonly)
        self.btn_promote = QPushButton("➡ 마스터로 편입")
        self.btn_promote.clicked.connect(self._promote)
        self.bar.insertWidget(3, self.btn_promote)
        self.btn_promote_all = QPushButton("⏩ 일괄 편입(업체ID 채운 행)")
        self.btn_promote_all.clicked.connect(self._promote_all)
        self.bar.insertWidget(4, self.btn_promote_all)
        if readonly:
            self.btn_promote.setEnabled(False)
            self.btn_promote_all.setEnabled(False)

    def _promote_all(self):
        """원작사_업체ID가 채워진(유효) 모든 행을 한 번에 편입 + 대기열에서 제거."""
        hidx = {h: i for i, h in enumerate(self.headers)}
        vc, tc = hidx.get("원작사_업체ID"), hidx.get("표준작품명", 0)
        if vc is None:
            QMessageBox.information(self, "일괄 편입", "원작사_업체ID 열이 없습니다.")
            return
        vids = _vendor_ids()
        targets = []
        for r in range(self.table.rowCount()):
            vit, tit = self.table.item(r, vc), self.table.item(r, tc)
            vid = vit.text().strip() if vit else ""
            title = tit.text().strip() if tit else ""
            if vid and title and vid in vids:
                rs = 0.10
                if "RS율" in hidx and self.table.item(r, hidx["RS율"]):
                    try:
                        rs = float(self.table.item(r, hidx["RS율"]).text() or 0.10)
                    except ValueError:
                        rs = 0.10
                targets.append((r, title, vid, rs))
        if not targets:
            QMessageBox.information(self, "일괄 편입",
                                    "원작사_업체ID가 유효하게 채워진 행이 없습니다.")
            return
        if QMessageBox.question(self, "일괄 편입",
                                f"{len(targets)}개 작품을 작품마스터로 편입하고 대기열에서 제거할까요?") \
                != QMessageBox.Yes:
            return
        new = [{"표준작품명": t, "원작사_업체ID": v, "원작사RS율": rs, "항목분류": "원작료",
                "상태": "정산중", "통화": "KRW", "비고": "ORPHAN 일괄편입"}
               for (_, t, v, rs) in targets]
        try:
            mio.append_rows(MASTER, "작품Alias마스터", new)
            for r in sorted((x[0] for x in targets), reverse=True):
                self.table.removeRow(r)
            mio.save_sheet(MASTER, "ORPHAN_등록대기", self.headers, self._collect())
            self.status.setText(
                f"일괄 편입 완료: {len(new)}개 → 작품마스터 등록 · 대기열에서 제거")
        except Exception as e:
            QMessageBox.warning(self, "일괄 편입 실패", str(e))

    def _promote(self):
        r = self.table.currentRow()
        if r < 0:
            QMessageBox.information(self, "편입", "편입할 작품 행을 먼저 선택하세요.")
            return
        tc = self.headers.index("표준작품명") if "표준작품명" in self.headers else 0
        it = self.table.item(r, tc)
        title = it.text().strip() if it else ""
        if not title:
            QMessageBox.information(self, "편입", "표준작품명이 비어있습니다.")
            return
        dlg = PromoteDialog(title, self)
        if dlg.exec() != QDialog.Accepted:
            return
        vid, rs, status = dlg.values()
        if vid not in _vendor_ids():                   # 참조 무결성
            QMessageBox.warning(self, "편입 실패", f"업체ID '{vid}' 가 업체마스터에 없습니다.")
            return
        try:
            mio.append_rows(MASTER, "작품Alias마스터", [{
                "표준작품명": title, "원작사_업체ID": vid, "원작사RS율": rs,
                "항목분류": "원작료", "상태": status, "통화": "KRW", "비고": "ORPHAN 편입"}])
            self.table.removeRow(r)
            mio.save_sheet(MASTER, "ORPHAN_등록대기", self.headers, self._collect())
            self.status.setText(f"편입 완료: {title} → {vid} (작품마스터 등록 · 대기열에서 제거)")
        except Exception as e:
            QMessageBox.warning(self, "편입 실패", str(e))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("원작료 정산 자동화 — 기초 세팅")
        self.resize(1000, 640)

        # 잠금 시도
        self.acquired, holder = mio.acquire_lock(MASTER)
        self.readonly = not self.acquired

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)

        # 잠금 배너
        banner = QLabel()
        if self.acquired:
            who = f"{holder.get('user')}·{holder.get('pc')}" if holder else "본인"
            banner.setText(f"🔓 편집 모드 · 잠금 보유: {who}    |    저장 시 자동 백업(최근 30개)")
            banner.setStyleSheet("background:#e1f5ee;color:#0f6e56;padding:6px 12px;font-size:12px")
        else:
            who = f"{holder.get('user')}·{holder.get('pc')}" if holder else "다른 사용자"
            banner.setText(f"🔒 읽기 전용 — 현재 {who} 님이 편집 중입니다. (편집·저장 불가)")
            banner.setStyleSheet("background:#fbeaf0;color:#72243e;padding:6px 12px;font-size:12px")
        root.addWidget(banner)

        # 툴바: 마스터 엑셀 내보내기 / 가져오기(병합·덮어쓰기)
        tb = QHBoxLayout()
        self.btn_export = QPushButton("📥 마스터 내보내기(엑셀)")
        self.btn_import = QPushButton("📤 마스터 가져오기(엑셀)")
        self.btn_export.clicked.connect(self._export_master)
        self.btn_import.clicked.connect(self._import_master)
        self.btn_import.setEnabled(self.acquired)        # 읽기전용이면 반영 불가
        tb.addWidget(self.btn_export)
        tb.addWidget(self.btn_import)
        tb.addStretch(1)
        root.addLayout(tb)

        # 본문: 좌 네비 + 우 스택
        body = QHBoxLayout()
        root.addLayout(body, 1)
        self.nav = QListWidget()
        self.nav.setFixedWidth(190)
        self.nav.setStyleSheet("font-size:13px")
        self.stack = QStackedWidget()
        body.addWidget(self.nav)
        body.addWidget(self.stack, 1)

        for label, sheet in NAV:
            QListWidgetItem(label, self.nav)
            if sheet == "ORPHAN_등록대기":
                self.stack.addWidget(OrphanTab(readonly=self.readonly))
            elif sheet:
                self.stack.addWidget(MasterTab(sheet, readonly=self.readonly))
            else:
                self.stack.addWidget(Placeholder(label))
        QListWidgetItem("환경 설정", self.nav)
        self.stack.addWidget(Placeholder("환경 설정"))

        self.nav.currentRowChanged.connect(self.stack.setCurrentIndex)
        self.nav.setCurrentRow(0)

    def _reload_all(self):
        for i in range(self.stack.count()):
            w = self.stack.widget(i)
            if hasattr(w, "load"):
                try:
                    w.load()
                except Exception:                    # noqa: BLE001
                    pass

    def _export_master(self):
        ts = mio._ts()
        dest, _ = QFileDialog.getSaveFileName(
            self, "마스터 내보내기", f"원작료정산_마스터_{ts}.xlsx", "Excel (*.xlsx)")
        if not dest:
            return
        try:
            mio.export_master(MASTER, dest)
            QMessageBox.information(self, "내보내기 완료",
                                    f"마스터를 내보냈습니다:\n{dest}\n\n"
                                    "이 파일을 수정한 뒤 '가져오기'로 반영하세요.")
        except Exception as e:                       # noqa: BLE001
            QMessageBox.warning(self, "내보내기 실패", str(e))

    def _import_master(self):
        src, _ = QFileDialog.getOpenFileName(
            self, "가져올 마스터(수정본) 선택", "", "Excel (*.xlsx)")
        if not src:
            return
        box = QMessageBox(self)
        box.setWindowTitle("반영 방식 선택")
        box.setIcon(QMessageBox.Question)
        box.setText("업로드한 파일을 어떻게 반영할까요?")
        box.setInformativeText(
            "· 병합: 키(업체ID·표준작품명 등) 기준으로 기존 행은 갱신, 새 행은 추가\n"
            "· 덮어쓰기: 공통 시트를 업로드 파일 내용으로 통째 교체\n"
            "(반영 전 자동 백업됩니다)")
        b_merge = box.addButton("병합", QMessageBox.AcceptRole)
        b_over = box.addButton("덮어쓰기", QMessageBox.DestructiveRole)
        box.addButton("취소", QMessageBox.RejectRole)
        box.exec()
        clicked = box.clickedButton()
        if clicked not in (b_merge, b_over):
            return
        mode = "merge" if clicked is b_merge else "overwrite"
        try:
            res = mio.import_master(MASTER, src, mode=mode)
            lines = [f"· {s}: {d['mode']} (추가 {d['added']} · 갱신 {d['updated']} · 총 {d['after']}행)"
                     for s, d in res.items()]
            self._reload_all()
            QMessageBox.information(self, "가져오기 완료",
                                    f"[{mode}] 반영 완료 (자동 백업됨)\n\n" + "\n".join(lines))
        except Exception as e:                       # noqa: BLE001
            QMessageBox.warning(self, "가져오기 실패", str(e))

    def closeEvent(self, e):
        if self.acquired:
            mio.release_lock(MASTER)
        super().closeEvent(e)


def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
