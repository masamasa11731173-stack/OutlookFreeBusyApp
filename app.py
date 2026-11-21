"""Outlook FreeBusy GUI: editable participant list (classic Outlook, Windows)."""
import json
import sys
from datetime import datetime, timedelta
from pathlib import Path

from PySide6.QtCore import QDate, QTime, Qt, QThread, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDateEdit, QGridLayout, QGroupBox,
    QHBoxLayout, QInputDialog, QLabel, QLineEdit, QListWidget, QMainWindow, QMessageBox,
    QTextEdit, QStackedWidget,
    QPushButton, QScrollArea, QTableWidget, QTableWidgetItem,
    QTimeEdit, QVBoxLayout, QWidget,
)

DEFAULT_PEOPLE = {
    "山田 太郎": "Yamada, Taro",
    "佐藤 花子": "Sato, Hanako",
    "鈴木 一郎": "鈴木 一郎",
    "高橋 美咲": "Takahashi, Misaki",
    "田中 健太": "田中 健太",
    "伊藤 彩": "Ito, Aya",
}
CONFIG_PATH = Path.home() / ".outlook_freebusy_people.json"
STATUS = {"0": "空き", "1": "仮予定", "2": "予定あり", "3": "外出中", "4": "他の場所で作業"}
COLORS = {
    "0": QColor("#dff4e8"), "1": QColor("#fff0bb"),
    "2": QColor("#f9d9d9"), "3": QColor("#f2c7df"),
    "4": QColor("#dce9f8"), "?": QColor("#e7e7e7"),
}


def load_people():
    if not CONFIG_PATH.exists():
        return dict(DEFAULT_PEOPLE)
    try:
        obj = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        if not isinstance(obj, dict) or not all(
            isinstance(k, str) and k.strip() and isinstance(v, str) and v.strip()
            for k, v in obj.items()
        ):
            raise ValueError("Invalid participant data")
        return obj
    except (OSError, ValueError) as exc:
        raise RuntimeError(f"参加者設定の読み込みに失敗しました: {CONFIG_PATH}\n{exc}") from exc


def save_people(people):
    temp = CONFIG_PATH.with_suffix(".tmp")
    temp.write_text(json.dumps(people, ensure_ascii=False, indent=2), encoding="utf-8")
    temp.replace(CONFIG_PATH)


class SearchWorker(QThread):
    finished_result = Signal(object)
    failed = Signal(str)

    def __init__(self, people, date_value, start_time, end_time, interval=15):
        super().__init__()
        self.people = people
        self.date_value = date_value
        self.start_time = start_time
        self.end_time = end_time
        self.interval = interval

    def run(self):
        import pythoncom
        pythoncom.CoInitialize()
        try:
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            day = datetime.combine(self.date_value, datetime.min.time())
            # Workaround observed in the user's environment; not a general guarantee.
            request_start = day - timedelta(hours=9)
            base = request_start.replace(hour=0, minute=0, second=0, microsecond=0)
            result = {}
            for label, lookup in self.people:
                recipient = namespace.CreateRecipient(lookup)
                if not recipient.Resolve():
                    raise RuntimeError(f"名前を解決できません: {label} ({lookup})")
                data = recipient.FreeBusy(request_start, self.interval, True)
                if not data:
                    raise RuntimeError(f"FreeBusy情報がありません: {label}")
                result[label] = data
            slots = []
            cursor = datetime.combine(self.date_value, self.start_time)
            finish = datetime.combine(self.date_value, self.end_time)
            while cursor + timedelta(minutes=self.interval) <= finish:
                idx = int((cursor - base).total_seconds() // 60) // self.interval
                codes = {
                    label: data[idx] if 0 <= idx < len(data) and data[idx] in STATUS else "?"
                    for label, data in result.items()
                }
                slots.append((cursor, codes))
                cursor += timedelta(minutes=self.interval)
            self.finished_result.emit(slots)
        except Exception as exc:
            self.failed.emit(f"{type(exc).__name__}: {exc}")
        finally:
            pythoncom.CoUninitialize()


DEFAULT_BODY = """各位

お疲れ様です。
下記のとおり打合せを開催いたします。

【日時】
{date} {start}～{end}

【場所】
{location}

【議題】
・

ご多忙のところ恐れ入りますが、ご出席のほどよろしくお願いいたします。
"""


class MeetingPage(QWidget):
    def __init__(self, parent, candidates, people):
        super().__init__(parent)
        self.candidates = list(candidates)
        self.people = list(people)
        layout = QVBoxLayout(self)
        heading = QLabel("会議案内の作成")
        heading.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(heading)
        note = QLabel("候補日時を選び、件名・場所・本文を編集してください。Outlook画面を開くだけで自動送信はしません。")
        note.setWordWrap(True)
        layout.addWidget(note)
        layout.addWidget(QLabel("共通空き時間の候補（選択すると本文の日時も更新されます）"))
        self.candidate_list = QListWidget()
        self.candidate_list.setMinimumHeight(125)
        for a, b in self.candidates:
            self.candidate_list.addItem(f"{a:%Y/%m/%d %H:%M} ～ {b:%H:%M}")
        layout.addWidget(self.candidate_list)
        form = QGridLayout()
        self.subject_edit = QLineEdit()
        self.subject_edit.setPlaceholderText("例：定例打合せ")
        self.location_edit = QLineEdit()
        self.location_edit.setPlaceholderText("例：会議室A")
        form.addWidget(QLabel("会議件名"), 0, 0)
        form.addWidget(self.subject_edit, 0, 1)
        form.addWidget(QLabel("場所"), 1, 0)
        form.addWidget(self.location_edit, 1, 1)
        layout.addLayout(form)
        layout.addWidget(QLabel("出席者：" + "、".join(label for label, _ in self.people)))
        row = QHBoxLayout()
        row.addWidget(QLabel("本文の定型文"))
        self.template_combo = QComboBox()
        self.template_combo.addItem("標準の会議案内")
        self.template_combo.addItem("簡潔な会議案内")
        row.addWidget(self.template_combo)
        self.template_button = QPushButton("定型文を挿入し直す")
        row.addWidget(self.template_button)
        layout.addLayout(row)
        self.body_edit = QTextEdit()
        self.body_edit.setPlaceholderText("会議案内の本文を入力してください")
        self.body_edit.setMinimumHeight(210)
        layout.addWidget(self.body_edit, 1)
        warning = QLabel("※ FreeBusyの空き判定は検証中です。送信前にOutlookのスケジュールアシスタントで確認してください。")
        warning.setWordWrap(True)
        warning.setStyleSheet("color: #9a5200;")
        layout.addWidget(warning)
        buttons = QHBoxLayout()
        self.create_button = QPushButton("Outlook会議作成画面を開く（送信しない）")
        self.create_button.setMinimumHeight(40)
        self.cancel_button = QPushButton("← 空き時間検索へ戻る")
        buttons.addWidget(self.create_button, 1)
        buttons.addWidget(self.cancel_button)
        layout.addLayout(buttons)
        self.candidate_list.currentRowChanged.connect(self.on_candidate_changed)
        self.template_button.clicked.connect(self.insert_template)
        self.create_button.clicked.connect(self.create_meeting)
        self.cancel_button.clicked.connect(parent.show_search_page)
        self.location_edit.textChanged.connect(self.update_location_in_template)
        self.open_drafts = []
        self._previous_location = ""
        if self.candidates:
            self.candidate_list.setCurrentRow(0)
        else:
            self.insert_template()

    def template_text(self):
        i = self.candidate_list.currentRow()
        if i < 0:
            return ""
        a, b = self.candidates[i]
        loc = self.location_edit.text().strip() or "（未定）"
        if self.template_combo.currentIndex() == 1:
            return (f"各位\n\nお疲れ様です。\n下記日程で打合せを行います。\n\n"
                    f"日時：{a:%Y/%m/%d} {a:%H:%M}～{b:%H:%M}\n"
                    f"場所：{loc}\n\nよろしくお願いいたします。")
        return DEFAULT_BODY.format(date=a.strftime("%Y/%m/%d"), start=a.strftime("%H:%M"), end=b.strftime("%H:%M"), location=loc)

    def insert_template(self):
        if self.body_edit.toPlainText().strip():
            reply = QMessageBox.question(self, "本文の上書き", "現在の本文を定型文で置き換えますか？", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply != QMessageBox.Yes:
                return
        self.body_edit.setPlainText(self.template_text())
        self._previous_location = self.location_edit.text().strip() or "（未定）"

    def on_candidate_changed(self, index):
        # Update only the generated date/time lines; preserve manually edited agenda.
        if index < 0:
            return
        text = self.body_edit.toPlainText()
        if not text.strip():
            self.body_edit.setPlainText(self.template_text())
            return
        import re
        a, b = self.candidates[index]
        date_line = f"{a:%Y/%m/%d} {a:%H:%M}～{b:%H:%M}"
        text = re.sub(r"(?m)^(?:(?:日時：)?\d{4}/\d{2}/\d{2} \d{2}:\d{2}～\d{2}:\d{2})$", lambda m: ("日時：" if m.group(0).startswith("日時：") else "") + date_line, text)
        self.body_edit.setPlainText(text)

    def update_location_in_template(self, new_location):
        if not hasattr(self, "_previous_location"):
            return
        new = new_location.strip() or "（未定）"
        old = self._previous_location
        text = self.body_edit.toPlainText()
        if old and old != new:
            text = text.replace("【場所】\n" + old, "【場所】\n" + new)
            text = text.replace("場所：" + old, "場所：" + new)
            self.body_edit.setPlainText(text)
        self._previous_location = new

    def create_meeting(self):
        index = self.candidate_list.currentRow()
        if not (0 <= index < len(self.candidates)):
            QMessageBox.warning(self, "候補未選択", "会議候補日時を選択してください。")
            return
        subject = self.subject_edit.text().strip()
        if not subject:
            QMessageBox.warning(self, "件名未入力", "会議件名を入力してください。")
            return
        a, b = self.candidates[index]
        answer = QMessageBox.question(
            self, "会議案内の確認",
            f"件名：{subject}\n日時：{a:%Y/%m/%d %H:%M}～{b:%H:%M}\n"
            f"出席者：{', '.join(label for label, _ in self.people)}\n\n"
            "Outlookの会議作成画面を開きます（送信しません）。続けますか？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        import pythoncom
        pythoncom.CoInitialize()
        try:
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application")
            meeting = outlook.CreateItem(1)
            meeting.MeetingStatus = 1
            meeting.Subject = subject
            meeting.Start = a.strftime("%Y-%m-%d %H:%M:%S")
            meeting.End = b.strftime("%Y-%m-%d %H:%M:%S")
            meeting.Location = self.location_edit.text().strip()
            meeting.Body = self.body_edit.toPlainText()
            for label, lookup in self.people:
                attendee = meeting.Recipients.Add(lookup)
                attendee.Type = 1
            if not meeting.Recipients.ResolveAll():
                QMessageBox.warning(self, "出席者の確認", "一部の出席者を名前解決できませんでした。Outlook画面で確認してください。")
            meeting.Display(False)
            self.open_drafts.append(meeting)
        except Exception as exc:
            QMessageBox.critical(self, "会議作成エラー", f"{type(exc).__name__}: {exc}")
        finally:
            pythoncom.CoUninitialize()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Outlook 空き時間検索・会議案内作成")
        self.resize(1100, 780)
        self.people = load_people()
        self.checks = {}
        self.worker = None
        self.search_snapshot = None
        self.candidates = []
        self.open_drafts = []

        self.pages = QStackedWidget()
        self.setCentralWidget(self.pages)
        self.search_page = QWidget()
        self.pages.addWidget(self.search_page)
        layout = QVBoxLayout(self.search_page)
        note = QLabel("※ FreeBusyの前日15:00指定は検証中です。会議設定前にOutlook画面で確認してください。")
        note.setStyleSheet("color: #9a5200; font-weight: bold;")
        note.setWordWrap(True)
        layout.addWidget(note)

        settings = QGroupBox("検索条件")
        settings_grid = QGridLayout(settings)
        self.date_edit = QDateEdit(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("yyyy/MM/dd")
        self.start_edit = QTimeEdit(QTime(8, 0))
        self.start_edit.setDisplayFormat("HH:mm")
        self.end_edit = QTimeEdit(QTime(17, 0))
        self.end_edit.setDisplayFormat("HH:mm")
        self.duration = QComboBox()
        for m in (15, 30, 45, 60, 90, 120):
            self.duration.addItem(f"{m}分", m)
        self.duration.setCurrentIndex(1)
        for col, (label, widget) in enumerate((
            ("日付", self.date_edit), ("開始", self.start_edit),
            ("終了", self.end_edit), ("必要時間", self.duration),
        )):
            settings_grid.addWidget(QLabel(label), 0, col)
            settings_grid.addWidget(widget, 1, col)
        layout.addWidget(settings)

        people_box = QGroupBox("参加者（チェックして検索）")
        people_layout = QVBoxLayout(people_box)
        self.people_area = QScrollArea()
        self.people_area.setWidgetResizable(True)
        self.people_area.setMinimumHeight(145)
        self.people_container = QWidget()
        self.people_grid = QGridLayout(self.people_container)
        self.people_area.setWidget(self.people_container)
        people_layout.addWidget(self.people_area)
        actions = QHBoxLayout()
        self.add_button = QPushButton("＋ 参加者を追加")
        self.remove_button = QPushButton("－ チェックした参加者を削除")
        self.select_button = QPushButton("全員選択")
        self.clear_button = QPushButton("選択解除")
        actions.addWidget(self.add_button)
        actions.addWidget(self.remove_button)
        actions.addWidget(self.select_button)
        actions.addWidget(self.clear_button)
        people_layout.addLayout(actions)
        layout.addWidget(people_box)
        self.add_button.clicked.connect(self.add_person)
        self.remove_button.clicked.connect(self.remove_people)
        self.select_button.clicked.connect(lambda: self.set_checks(True))
        self.clear_button.clicked.connect(lambda: self.set_checks(False))
        self.refresh_people(set(list(self.people)[:3]))

        self.search_button = QPushButton("空き時間を検索")
        self.search_button.setMinimumHeight(42)
        self.search_button.clicked.connect(self.search)
        layout.addWidget(self.search_button)
        self.summary = QLabel("参加者と日付を選んで検索してください。")
        self.summary.setWordWrap(True)
        layout.addWidget(self.summary)
        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        layout.addWidget(self.table, 1)

        self.meeting_button = QPushButton("次へ：会議案内を作成 →")
        self.meeting_button.setMinimumHeight(40)
        self.meeting_button.setEnabled(False)
        self.meeting_button.clicked.connect(self.open_meeting_dialog)
        layout.addWidget(self.meeting_button)
        self.meeting_page = None

        self.date_edit.dateChanged.connect(self.invalidate_candidates)
        self.start_edit.timeChanged.connect(self.invalidate_candidates)
        self.end_edit.timeChanged.connect(self.invalidate_candidates)
        self.duration.currentIndexChanged.connect(self.invalidate_candidates)

    def refresh_people(self, selected=None):
        selected = set(selected or [])
        while self.people_grid.count():
            item = self.people_grid.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
        self.checks = {}
        for i, label in enumerate(self.people):
            cb = QCheckBox(label)
            cb.setToolTip(self.people[label])
            cb.setChecked(label in selected)
            self.people_grid.addWidget(cb, i // 3, i % 3)
            self.checks[label] = cb
            cb.toggled.connect(self.invalidate_candidates)

    def invalidate_candidates(self, *args):
        self.search_snapshot = None
        self.candidates = []
        self.meeting_button.setEnabled(False)
        if self.meeting_page is not None:
            self.pages.removeWidget(self.meeting_page)
            self.open_drafts.extend(self.meeting_page.open_drafts)
            self.meeting_page.deleteLater()
            self.meeting_page = None

    def show_search_page(self):
        self.pages.setCurrentWidget(self.search_page)

    def open_meeting_dialog(self):
        if not self.search_snapshot or not self.candidates:
            QMessageBox.information(self, "候補なし", "先に空き時間を検索してください。")
            return
        # Keep the same editor when navigating back and forth, so edits are preserved.
        if self.meeting_page is None:
            self.meeting_page = MeetingPage(self, self.candidates, self.search_snapshot)
            self.pages.addWidget(self.meeting_page)
        self.pages.setCurrentWidget(self.meeting_page)

    def set_checks(self, value):
        for cb in self.checks.values():
            cb.setChecked(value)

    def add_person(self):
        label, ok = QInputDialog.getText(self, "参加者を追加", "表示名（例：山田 太郎）")
        if not ok:
            return
        label = label.strip()
        if not label:
            QMessageBox.warning(self, "入力不足", "表示名を入力してください。")
            return
        if label in self.people:
            QMessageBox.warning(self, "登録済み", "同じ表示名がすでに登録されています。")
            return
        lookup, ok = QInputDialog.getText(
            self, "Outlook検索名", "Outlookの名前またはメールアドレス", text=label
        )
        if not ok:
            return
        lookup = lookup.strip()
        if not lookup:
            QMessageBox.warning(self, "入力不足", "Outlook検索名を入力してください。")
            return
        selected = {name for name, cb in self.checks.items() if cb.isChecked()}
        new_people = dict(self.people)
        new_people[label] = lookup
        try:
            save_people(new_people)
        except OSError as exc:
            QMessageBox.critical(self, "保存エラー", str(exc))
            return
        self.people = new_people
        selected.add(label)
        self.refresh_people(selected)
        self.summary.setText(f"{label} を追加しました。Outlookでの名前解決は検索時に確認します。")

    def remove_people(self):
        selected = [name for name, cb in self.checks.items() if cb.isChecked()]
        if not selected:
            QMessageBox.information(self, "削除対象なし", "削除する参加者にチェックを入れてください。")
            return
        reply = QMessageBox.question(
            self, "削除の確認", "次の参加者を登録リストから削除しますか？\n" + "、".join(selected),
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return
        new_people = {k: v for k, v in self.people.items() if k not in selected}
        try:
            save_people(new_people)
        except OSError as exc:
            QMessageBox.critical(self, "保存エラー", str(exc))
            return
        self.people = new_people
        self.refresh_people()
        self.summary.setText(f"{len(selected)}人を登録リストから削除しました。")

    def search(self):
        people = [(label, self.people[label]) for label, cb in self.checks.items() if cb.isChecked()]
        if not people:
            QMessageBox.warning(self, "参加者未選択", "参加者を1人以上選択してください。")
            return
        start_time = self.start_edit.time().toPython()
        end_time = self.end_edit.time().toPython()
        if start_time >= end_time:
            QMessageBox.warning(self, "時間設定", "終了時刻は開始時刻より後にしてください。")
            return
        if start_time.minute % 15 or end_time.minute % 15:
            QMessageBox.warning(self, "時間設定", "開始・終了時刻は15分刻みで指定してください。")
            return
        self.invalidate_candidates()
        self.search_button.setEnabled(False)
        self.summary.setText("Outlookから予定情報を取得中です…")
        self.table.setRowCount(0)
        self.worker = SearchWorker(people, self.date_edit.date().toPython(), start_time, end_time)
        labels = [p[0] for p in people]
        self.pending_people = list(people)
        self.worker.finished_result.connect(lambda slots: self.show_result(slots, labels))
        self.worker.failed.connect(self.show_error)
        self.worker.finished.connect(lambda: self.search_button.setEnabled(True))
        self.worker.start()

    def show_error(self, error):
        self.invalidate_candidates()
        self.summary.setText("検索に失敗しました。結果は表示しません。")
        QMessageBox.critical(self, "Outlook取得エラー", error)

    def show_result(self, slots, labels):
        self.table.setColumnCount(len(labels) + 2)
        self.table.setHorizontalHeaderLabels(["時刻", *labels, "全員空き"])
        self.table.setRowCount(len(slots))
        periods = []
        period_start = None
        interval = timedelta(minutes=15)
        for row, (slot_time, codes) in enumerate(slots):
            self.table.setItem(row, 0, QTableWidgetItem(slot_time.strftime("%H:%M")))
            for col, label in enumerate(labels, start=1):
                code = codes[label]
                cell = QTableWidgetItem(STATUS.get(code, "取得不可"))
                cell.setBackground(COLORS.get(code, COLORS["?"]))
                self.table.setItem(row, col, cell)
            all_free = all(code == "0" for code in codes.values())
            result = QTableWidgetItem("○" if all_free else "×")
            result.setTextAlignment(Qt.AlignCenter)
            result.setBackground(COLORS["0"] if all_free else COLORS["?"])
            self.table.setItem(row, len(labels) + 1, result)
            if all_free and period_start is None:
                period_start = slot_time
            elif not all_free and period_start is not None:
                periods.append((period_start, slot_time))
                period_start = None
        if period_start is not None and slots:
            periods.append((period_start, slots[-1][0] + interval))
        self.table.resizeColumnsToContents()
        minimum = self.duration.currentData()
        matches = [(a, b) for a, b in periods if (b - a).total_seconds() / 60 >= minimum]
        self.search_snapshot = list(self.pending_people)
        self.candidates = []
        for period_begin, period_end in matches:
            candidate_start = period_begin
            while candidate_start + timedelta(minutes=minimum) <= period_end:
                candidate_end = candidate_start + timedelta(minutes=minimum)
                self.candidates.append((candidate_start, candidate_end))
                candidate_start += timedelta(minutes=15)
        self.meeting_button.setEnabled(bool(self.candidates))
        if matches:
            summary = " / ".join(
                f"{a:%H:%M}～{b:%H:%M}（{int((b-a).total_seconds()/60)}分）" for a, b in matches
            )
            self.summary.setText(
                f"共通空き時間候補（{minimum}分以上）：{summary}\n"
                "※暫定判定。Outlook画面で最終確認してください。"
            )
        else:
            self.summary.setText(
                f"{minimum}分以上の共通空き時間候補はありません。\n"
                "※暫定判定。Outlook画面で最終確認してください。"
            )


def main():
    app = QApplication(sys.argv)
    try:
        window = MainWindow()
    except RuntimeError as exc:
        QMessageBox.critical(None, "参加者設定エラー", str(exc))
        return 1
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
