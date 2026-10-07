import sys
import os
import json
from PIL import Image, ImageChops

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QDialog, QComboBox, QCheckBox, QLineEdit,
    QSlider, QFormLayout, QGroupBox, QKeySequenceEdit,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QRect
from PyQt6.QtGui import (
    QFont, QGuiApplication, QIcon, QPixmap, QPainter, QColor,
    QRadialGradient, QPen, QKeySequence,
)

try:
    import win32com.client
    import win32gui
    import pythoncom
    WIN32_OK = True
except ImportError:
    WIN32_OK = False

try:
    import keyboard
    KEYBOARD_OK = True
except ImportError:
    KEYBOARD_OK = False

# ─── Version ─────────────────────────────────────────────────────────────────

VERSION = 5
CREATED = "17/04/2026 à 19:30"
APP_TITLE = f"BookBinder — {VERSION} — {CREATED}"

# ─── Settings ────────────────────────────────────────────────────────────────

SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "settings.json")

DEFAULTS = {
    "format": "JPEG",
    "quality": 95,
    "trim_border": True,
    "suffix": "_join",
    "delete_sources": False,
    "hotkey_merge": "ctrl+shift+d",
    "hotkey_merge_reverse": "ctrl+shift+f",
}

FORMAT_EXT = {"JPEG": ".jpg", "JPEG2000": ".jp2", "PNG": ".png", "WebP": ".webp"}
FORMAT_QUALITY_DEFAULTS = {"JPEG": 95, "JPEG2000": 85, "PNG": 6, "WebP": 90}
FORMAT_QUALITY_RANGES = {"JPEG": (1, 100), "JPEG2000": (1, 100), "PNG": (0, 9), "WebP": (1, 100)}
FORMAT_QUALITY_LABELS = {
    "JPEG": "Qualité (1–100)",
    "JPEG2000": "Qualité (1–100)",
    "PNG": "Compression (0–9)",
    "WebP": "Qualité (1–100)",
}


def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r") as f:
                return {**DEFAULTS, **json.load(f)}
        except Exception:
            pass
    return DEFAULTS.copy()


def save_settings(s):
    with open(SETTINGS_FILE, "w") as f:
        json.dump(s, f, indent=2)


def keyboard_to_qt(ks: str) -> str:
    return "+".join(p.capitalize() for p in ks.split("+"))


def qt_to_keyboard(ks: str) -> str:
    return ks.lower()

# ─── App Icon ─────────────────────────────────────────────────────────────────


def create_icon() -> QIcon:
    ico_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "fusion_images.ico")
    if os.path.exists(ico_path):
        icon = QIcon(ico_path)
        if not icon.isNull():
            return icon
    # Repli si le .ico est absent : icône dessinée à la volée
    size = 64
    px = QPixmap(size, size)
    px.fill(Qt.GlobalColor.transparent)
    p = QPainter(px)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    grad = QRadialGradient(32, 28, 34)
    grad.setColorAt(0.0, QColor("#F8B8D8"))
    grad.setColorAt(1.0, QColor("#C44C88"))
    p.setBrush(grad)
    p.setPen(Qt.PenStyle.NoPen)
    p.drawEllipse(1, 1, 62, 62)
    p.setPen(QPen(QColor(255, 255, 255)))
    p.setFont(QFont("Segoe UI", 26, QFont.Weight.Bold))
    p.drawText(QRect(0, 2, 64, 64), Qt.AlignmentFlag.AlignCenter, "FI")
    p.end()
    return QIcon(px)

# ─── Image Processing ─────────────────────────────────────────────────────────


def trim_border(img):
    border_color = img.getpixel((0, 0))
    bg = Image.new(img.mode, img.size, border_color)
    diff = ImageChops.difference(img, bg)
    bbox = diff.getbbox()
    return img.crop(bbox) if bbox else img


def merge_images(path1, path2, settings):
    img1 = Image.open(path1).convert("RGB")
    img2 = Image.open(path2).convert("RGB")

    if settings["trim_border"]:
        img1 = trim_border(img1)
        img2 = trim_border(img2)

    max_h = max(img1.height, img2.height)
    img1 = img1.resize((int(img1.width * max_h / img1.height), max_h), Image.LANCZOS)
    img2 = img2.resize((int(img2.width * max_h / img2.height), max_h), Image.LANCZOS)

    result = Image.new("RGB", (img1.width + img2.width, max_h), (255, 255, 255))
    result.paste(img1, (0, 0))
    result.paste(img2, (img1.width, 0))

    fmt = settings["format"]
    ext = FORMAT_EXT[fmt]
    q = settings["quality"]
    suffix = settings.get("suffix", "_join") or "_join"
    folder = os.path.dirname(path1)
    base = os.path.splitext(os.path.basename(path1))[0] + suffix

    candidate = f"{base}{ext}"
    i = 1
    while os.path.exists(os.path.join(folder, candidate)):
        candidate = f"{base}_{i}{ext}"
        i += 1
    out = os.path.join(folder, candidate)

    if fmt == "JPEG":
        result.save(out, "JPEG", quality=q, optimize=True)
    elif fmt == "JPEG2000":
        rate = max(0.01, q / 100.0)
        result.save(out, "JPEG2000", irreversible=True, quality_mode="rates", quality_layers=[rate])
    elif fmt == "PNG":
        result.save(out, "PNG", compress_level=min(9, max(0, q)))
    elif fmt == "WebP":
        result.save(out, "WEBP", quality=q, method=6)

    if settings.get("delete_sources"):
        try:
            import send2trash
            send2trash.send2trash(path1)
            send2trash.send2trash(path2)
        except ImportError:
            # send2trash absent : on ne supprime PAS (jamais de suppression
            # definitive silencieuse a la place de la corbeille)
            pass

    return out

# ─── Explorer Selection ───────────────────────────────────────────────────────


IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".jp2", ".bmp", ".tiff", ".tif", ".gif"}


def get_explorer_selection(hwnd: int):
    if not WIN32_OK:
        return []
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass
    try:
        shell = win32com.client.Dispatch("Shell.Application")
        windows = shell.Windows()
        n = windows.Count

        # Essai 1 : fenêtre correspondant au HWND capturé
        if hwnd:
            for i in range(n):
                try:
                    w = windows.Item(i)
                    if int(w.HWND) == int(hwnd):
                        sel = w.Document.SelectedItems()
                        files = [sel.Item(j).Path for j in range(sel.Count)]
                        result = sorted(f for f in files if os.path.splitext(f)[1].lower() in IMAGE_EXTS)
                        if result:
                            return result
                except Exception:
                    continue

        # Essai 2 : n'importe quelle fenêtre Explorer avec des images sélectionnées
        for i in range(n):
            try:
                w = windows.Item(i)
                sel = w.Document.SelectedItems()
                files = [sel.Item(j).Path for j in range(sel.Count)]
                result = sorted(f for f in files if os.path.splitext(f)[1].lower() in IMAGE_EXTS)
                if result:
                    return result
            except Exception:
                continue
    except Exception:
        pass
    return []

# ─── Workers ─────────────────────────────────────────────────────────────────


class HotkeyWorker(QThread):
    # Émet (reverse, files) — fichiers lus immédiatement dans le thread du callback clavier,
    # avant qu'Explorer traite le raccourci et potentiellement change la sélection
    merge_requested = pyqtSignal(bool, object)

    def __init__(self, hotkey_merge, hotkey_reverse):
        super().__init__()
        self._running = True
        self.hotkey_merge = hotkey_merge
        self.hotkey_reverse = hotkey_reverse

    def _on_hotkey(self, reverse: bool):
        hwnd = win32gui.GetForegroundWindow() if WIN32_OK else 0
        files = get_explorer_selection(hwnd)
        self.merge_requested.emit(reverse, files)

    def run(self):
        if not KEYBOARD_OK:
            return
        try:
            keyboard.add_hotkey(self.hotkey_merge, lambda: self._on_hotkey(False))
            keyboard.add_hotkey(self.hotkey_reverse, lambda: self._on_hotkey(True))
        except Exception:
            pass
        while self._running:
            self.msleep(100)

    def stop(self):
        self._running = False
        if KEYBOARD_OK:
            try:
                keyboard.remove_hotkey(self.hotkey_merge)
                keyboard.remove_hotkey(self.hotkey_reverse)
            except Exception:
                pass


class MergeWorker(QThread):
    finished = pyqtSignal(bool, str)

    def __init__(self, files, settings, reverse):
        super().__init__()
        self.files = files
        self.settings = settings
        self.reverse = reverse

    def run(self):
        try:
            if len(self.files) < 2:
                n = len(self.files)
                self.finished.emit(False, f"{n} image(s) détectée(s) dans l'explorateur — il en faut 2.")
                return
            f1, f2 = self.files[0], self.files[1]
            if self.reverse:
                f1, f2 = f2, f1
            out = merge_images(f1, f2, self.settings)
            self.finished.emit(True, os.path.basename(out))
        except Exception as e:
            self.finished.emit(False, str(e))

# ─── Toast base ───────────────────────────────────────────────────────────────

_active_toasts: list = []
_TOAST_X = 24
_TOAST_TOP = 24
_TOAST_GAP = 8


def _next_toast_y() -> int:
    y = _TOAST_TOP
    for t in _active_toasts:
        if t.isVisible():
            y += t.height() + _TOAST_GAP
    return y


class ToastBase(QWidget):
    def __init__(self, bg: str, border: str):
        super().__init__(
            None,
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool,
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self._bg = QColor(bg)
        self._border = QColor(border)

    def paintEvent(self, _event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setBrush(self._bg)
        p.setPen(QPen(self._border, 2))
        p.drawRoundedRect(self.rect().adjusted(1, 1, -1, -1), 16, 16)

    def _show_stacked(self):
        self.adjustSize()
        screen = QGuiApplication.primaryScreen().geometry()
        y = _next_toast_y()
        self.move(screen.width() - self.width() - _TOAST_X, y)
        _active_toasts.append(self)
        self.show()

    def closeEvent(self, event):
        if self in _active_toasts:
            _active_toasts.remove(self)
        super().closeEvent(event)

# ─── Toast Progress ───────────────────────────────────────────────────────────


SPINNER_CHARS = ["◐", "◓", "◑", "◒"]


class ToastProgress(ToastBase):
    def __init__(self):
        super().__init__("#FFF0F8", "#F5C6DB")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 14, 20, 14)
        layout.setSpacing(6)

        header = QHBoxLayout()
        lbl_icon = QLabel("⚡🔥")
        lbl_icon.setFont(QFont("Segoe UI Emoji", 18))
        lbl_icon.setStyleSheet("background: transparent; border: none;")
        lbl_title = QLabel("Fusion !")
        lbl_title.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        lbl_title.setStyleSheet("color: #C45C8A; background: transparent; border: none;")
        header.addWidget(lbl_icon)
        header.addWidget(lbl_title)
        header.addStretch()
        layout.addLayout(header)

        spinner_row = QHBoxLayout()
        self._lbl_spinner = QLabel("◐")
        self._lbl_spinner.setFont(QFont("Segoe UI", 13))
        self._lbl_spinner.setStyleSheet("color: #E88ABE; background: transparent; border: none;")
        lbl_txt = QLabel("Fusion en cours...")
        lbl_txt.setFont(QFont("Segoe UI", 11))
        lbl_txt.setStyleSheet("color: #8860A0; background: transparent; border: none;")
        spinner_row.addWidget(self._lbl_spinner)
        spinner_row.addWidget(lbl_txt)
        layout.addLayout(spinner_row)

        self._idx = 0
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)
        self._timer.start(180)
        self._show_stacked()

    def _tick(self):
        self._idx = (self._idx + 1) % len(SPINNER_CHARS)
        self._lbl_spinner.setText(SPINNER_CHARS[self._idx])

    def closeEvent(self, event):
        self._timer.stop()
        super().closeEvent(event)

# ─── Toast Result ─────────────────────────────────────────────────────────────


class ToastResult(ToastBase):
    def __init__(self, message, success=True):
        bg = "#F0FFF4" if success else "#FFF0F0"
        border = "#A8F5C0" if success else "#F5A8A8"
        super().__init__(bg, border)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 16, 22, 16)
        layout.setSpacing(8)

        header = QHBoxLayout()
        lbl_icon = QLabel("✅" if success else "❌")
        lbl_icon.setFont(QFont("Segoe UI Emoji", 20))
        lbl_icon.setStyleSheet("background: transparent; border: none;")
        lbl_title = QLabel("Fusion ok !" if success else "Erreur")
        lbl_title.setFont(QFont("Segoe UI", 15, QFont.Weight.Bold))
        color = "#3A8A3A" if success else "#A03030"
        lbl_title.setStyleSheet(f"color: {color}; background: transparent; border: none;")
        header.addWidget(lbl_icon)
        header.addWidget(lbl_title)
        header.addStretch()
        layout.addLayout(header)

        lbl_msg = QLabel(message)
        lbl_msg.setFont(QFont("Segoe UI", 11))
        lbl_msg.setStyleSheet("color: #5C3D5E; background: transparent; border: none;")
        lbl_msg.setWordWrap(True)
        lbl_msg.setMaximumWidth(340)
        layout.addWidget(lbl_msg)

        self.setMinimumWidth(300)
        self._show_stacked()
        QTimer.singleShot(5000, self.close)

# ─── Info Dialog ──────────────────────────────────────────────────────────────


INFO_STYLE = """
QDialog, QWidget { background-color: #FFF0F8; font-family: 'Segoe UI'; }
QLabel { color: #5C3D5E; background: transparent; }
QPushButton {
    background-color: #F5A8CC; color: white; border: none;
    border-radius: 16px; padding: 9px 28px; font-weight: bold; font-size: 13px;
}
QPushButton:hover { background-color: #E88ABE; }
"""


class InfoDialog(QDialog):
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.setWindowTitle("À propos — BookBinder")
        self.setModal(True)
        self.setMinimumWidth(430)
        self.setStyleSheet(INFO_STYLE)

        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(28, 28, 28, 28)

        title = QLabel("🌸  BookBinder")
        title.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        title.setStyleSheet("color: #C45C8A;")
        layout.addWidget(title)

        ver = QLabel(f"Version {VERSION}  ·  {CREATED}")
        ver.setFont(QFont("Segoe UI", 9))
        ver.setStyleSheet("color: #C0A0D4;")
        layout.addWidget(ver)

        sep = QWidget()
        sep.setFixedHeight(1)
        sep.setStyleSheet("background-color: #F5C6DB;")
        layout.addWidget(sep)

        hk1 = settings.get("hotkey_merge", "ctrl+shift+d").upper()
        hk2 = settings.get("hotkey_merge_reverse", "ctrl+shift+f").upper()

        content = [
            ("🎯  Comment utiliser", True),
            (f"1. Sélectionne <b>2 images</b> dans l'explorateur Windows", False),
            (f"2. Appuie sur <b>{hk1}</b> pour fusionner (gauche → droite, ordre alpha)", False),
            (f"3. Appuie sur <b>{hk2}</b> pour fusionner (ordre inversé)", False),
            ("", False),
            ("🖼️  Ce que fait l'app", True),
            ("• Supprime les bordures unies (cadres de scan)", False),
            ("• Harmonise la hauteur des deux images", False),
            ("• Colle les deux images côte à côte", False),
            ("• Enregistre le résultat dans le même dossier que l'image source", False),
        ]

        for text, bold in content:
            if not text:
                layout.addSpacing(2)
                continue
            lbl = QLabel(text)
            lbl.setTextFormat(Qt.TextFormat.RichText)
            lbl.setFont(QFont("Segoe UI", 11 if bold else 10, QFont.Weight.Bold if bold else QFont.Weight.Normal))
            if bold:
                lbl.setStyleSheet("color: #C45C8A;")
            layout.addWidget(lbl)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn = QPushButton("Fermer")
        btn.clicked.connect(self.accept)
        btn_row.addWidget(btn)
        layout.addLayout(btn_row)

# ─── Options Dialog ───────────────────────────────────────────────────────────


DIALOG_STYLE = """
QDialog, QWidget { background-color: #FFF0F8; font-family: 'Segoe UI'; }
QGroupBox {
    border: 2px solid #F5C6DB; border-radius: 12px; margin-top: 12px;
    font-weight: bold; color: #C45C8A; padding: 10px 8px 8px 8px;
}
QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 6px; }
QLabel { color: #5C3D5E; background: transparent; }
QComboBox {
    border: 2px solid #F5C6DB; border-radius: 8px; padding: 6px 12px;
    background: white; color: #5C3D5E; min-width: 120px;
}
QComboBox:hover { border-color: #E88ABE; }
QComboBox::drop-down { border: none; width: 24px; }
QSlider::groove:horizontal { height: 6px; background: #F5C6DB; border-radius: 3px; }
QSlider::handle:horizontal {
    background: #F5A8CC; border: none; width: 18px; height: 18px;
    margin: -6px 0; border-radius: 9px;
}
QSlider::sub-page:horizontal { background: #E88ABE; border-radius: 3px; }
QCheckBox { color: #5C3D5E; font-size: 13px; spacing: 8px; }
QCheckBox::indicator {
    width: 18px; height: 18px; border-radius: 5px;
    border: 2px solid #F5C6DB; background: white;
}
QCheckBox::indicator:checked { background: #F5A8CC; border-color: #E88ABE; }
QLineEdit, QKeySequenceEdit {
    border: 2px solid #F5C6DB; border-radius: 8px; padding: 6px 12px;
    background: white; color: #5C3D5E;
}
QLineEdit:focus, QKeySequenceEdit:focus { border-color: #E88ABE; }
QPushButton#btnSave {
    background-color: #F5A8CC; color: white; border: none;
    border-radius: 16px; padding: 10px 28px; font-weight: bold; font-size: 13px;
}
QPushButton#btnSave:hover { background-color: #E88ABE; }
QPushButton#btnCancel {
    background-color: #EDE0F5; color: #8860A8; border: none;
    border-radius: 16px; padding: 10px 28px; font-weight: bold; font-size: 13px;
}
QPushButton#btnCancel:hover { background-color: #D8C8EC; }
"""


class OptionsDialog(QDialog):
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.settings = settings.copy()
        self.setWindowTitle("Options — BookBinder")
        self.setModal(True)
        self.setMinimumWidth(430)
        self.setStyleSheet(DIALOG_STYLE)
        self._build()

    def _build(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 24, 24, 24)

        title = QLabel("⚙️  Options")
        title.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        title.setStyleSheet("color: #C45C8A;")
        layout.addWidget(title)

        # Format
        fmt_group = QGroupBox("Format de sortie")
        fmt_layout = QFormLayout(fmt_group)
        fmt_layout.setSpacing(10)

        self.combo_fmt = QComboBox()
        self.combo_fmt.addItems(["JPEG", "JPEG2000", "PNG", "WebP"])
        self.combo_fmt.setCurrentText(self.settings["format"])
        self.combo_fmt.currentTextChanged.connect(self._on_format_changed)
        fmt_layout.addRow("Format :", self.combo_fmt)

        self.lbl_quality = QLabel()
        self.slider_quality = QSlider(Qt.Orientation.Horizontal)
        self._update_quality_ui(self.settings["format"], self.settings["quality"])
        self.slider_quality.valueChanged.connect(
            lambda v: self.lbl_quality.setText(
                f"{FORMAT_QUALITY_LABELS[self.combo_fmt.currentText()]} : {v}"
            )
        )
        fmt_layout.addRow(self.lbl_quality, self.slider_quality)
        layout.addWidget(fmt_group)

        # Options
        opt_group = QGroupBox("Options")
        opt_layout = QVBoxLayout(opt_group)
        opt_layout.setSpacing(10)

        self.check_trim = QCheckBox("Supprimer les bordures unies (cadres de scan)")
        self.check_trim.setChecked(self.settings["trim_border"])
        opt_layout.addWidget(self.check_trim)

        self.check_delete = QCheckBox("Envoyer les images source à la corbeille après fusion")
        self.check_delete.setChecked(self.settings["delete_sources"])
        opt_layout.addWidget(self.check_delete)

        row = QHBoxLayout()
        row.addWidget(QLabel("Suffixe :"))
        self.edit_suffix = QLineEdit(self.settings["suffix"])
        self.edit_suffix.setPlaceholderText("_join")
        row.addWidget(self.edit_suffix)
        opt_layout.addLayout(row)
        layout.addWidget(opt_group)

        # Hotkeys
        hk_group = QGroupBox("Raccourcis clavier  (cliquer puis appuyer sur la combinaison)")
        hk_layout = QFormLayout(hk_group)
        hk_layout.setSpacing(10)

        self.kse_merge = QKeySequenceEdit(
            QKeySequence(keyboard_to_qt(self.settings.get("hotkey_merge", "ctrl+shift+d")))
        )
        self.kse_reverse = QKeySequenceEdit(
            QKeySequence(keyboard_to_qt(self.settings.get("hotkey_merge_reverse", "ctrl+shift+f")))
        )
        hk_layout.addRow("Fusionner :", self.kse_merge)
        hk_layout.addRow("Fusionner (inversé) :", self.kse_reverse)
        layout.addWidget(hk_group)

        # Buttons
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QPushButton("Annuler")
        btn_cancel.setObjectName("btnCancel")
        btn_cancel.clicked.connect(self.reject)
        btn_save = QPushButton("Enregistrer")
        btn_save.setObjectName("btnSave")
        btn_save.clicked.connect(self._save)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_save)
        layout.addLayout(btn_row)

    def _update_quality_ui(self, fmt, value=None):
        lo, hi = FORMAT_QUALITY_RANGES[fmt]
        self.slider_quality.setRange(lo, hi)
        v = value if value is not None else FORMAT_QUALITY_DEFAULTS[fmt]
        self.slider_quality.setValue(max(lo, min(hi, v)))
        self.lbl_quality.setText(f"{FORMAT_QUALITY_LABELS[fmt]} : {self.slider_quality.value()}")

    def _on_format_changed(self, fmt):
        self._update_quality_ui(fmt)

    def _save(self):
        self.settings["format"] = self.combo_fmt.currentText()
        self.settings["quality"] = self.slider_quality.value()
        self.settings["trim_border"] = self.check_trim.isChecked()
        self.settings["delete_sources"] = self.check_delete.isChecked()
        suffix = self.edit_suffix.text().strip()
        self.settings["suffix"] = suffix if suffix else "_join"
        hk1 = qt_to_keyboard(self.kse_merge.keySequence().toString())
        hk2 = qt_to_keyboard(self.kse_reverse.keySequence().toString())
        if hk1:
            self.settings["hotkey_merge"] = hk1
        if hk2:
            self.settings["hotkey_merge_reverse"] = hk2
        self.accept()

    def get_settings(self):
        return self.settings

# ─── Styles ──────────────────────────────────────────────────────────────────


BTN_OPTIONS_STYLE = """
QPushButton {
    background-color: #D9B8E8; color: white; border: none;
    border-radius: 22px; padding: 12px 38px; font-size: 14px;
    font-weight: bold; min-width: 140px;
}
QPushButton:hover { background-color: #C2A0D5; }
QPushButton:pressed { background-color: #A880BF; }
"""

BTN_INFO_STYLE = """
QPushButton {
    background-color: transparent; color: #D4A0D8;
    border: 2px solid #E8C8F0; border-radius: 18px;
    font-size: 16px; font-weight: bold;
    min-width: 36px; max-width: 36px; min-height: 36px; max-height: 36px;
}
QPushButton:hover { background-color: #F5E8FF; color: #C45C8A; border-color: #D4A8E8; }
"""

MAIN_BG = "#FFF0F8"

# ─── Main Window ──────────────────────────────────────────────────────────────


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = load_settings()
        self.hotkey_worker = None
        self._merge_workers = []
        self._progress_toast = None
        self._result_toast = None

        icon = create_icon()
        self.setWindowIcon(icon)
        self.setWindowTitle(APP_TITLE)
        self.setFixedSize(440, 290)
        self.setStyleSheet(
            f"QMainWindow, QWidget {{ background-color: {MAIN_BG}; font-family: 'Segoe UI'; }}"
        )
        self._build_ui()
        self._start_hotkeys()

        if not WIN32_OK:
            self._show_result("pywin32 manquant — pip install pywin32", False)
        if not KEYBOARD_OK:
            self._show_result("keyboard manquant — pip install keyboard", False)

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setSpacing(10)
        root.setContentsMargins(36, 26, 36, 26)

        # Top row: title + info button
        top_row = QHBoxLayout()
        title = QLabel("🌸  BookBinder")
        title.setFont(QFont("Segoe UI", 20, QFont.Weight.Bold))
        title.setStyleSheet("color: #C45C8A; background: transparent;")
        btn_info = QPushButton("ℹ")
        btn_info.setStyleSheet(BTN_INFO_STYLE)
        btn_info.setToolTip("À propos")
        btn_info.clicked.connect(self._open_info)
        top_row.addWidget(title)
        top_row.addStretch()
        top_row.addWidget(btn_info)
        root.addLayout(top_row)

        sub = QLabel("Fusion de doubles pages BD")
        sub.setFont(QFont("Segoe UI", 10))
        sub.setStyleSheet("color: #C0A0D4; background: transparent;")
        root.addWidget(sub)

        root.addSpacing(4)

        # Status indicator
        status_row = QHBoxLayout()
        dot = QLabel("●")
        dot.setFont(QFont("Segoe UI", 10))
        dot.setStyleSheet("color: #60CC80; background: transparent;")
        self.lbl_status = QLabel("Raccourcis actifs")
        self.lbl_status.setFont(QFont("Segoe UI", 10))
        self.lbl_status.setStyleSheet("color: #8860A0; background: transparent;")
        status_row.addWidget(dot)
        status_row.addWidget(self.lbl_status)
        status_row.addStretch()
        root.addLayout(status_row)

        root.addSpacing(2)

        # Hotkey hints (updated when settings change)
        self.lbl_hints = QLabel()
        self.lbl_hints.setFont(QFont("Segoe UI", 9))
        self.lbl_hints.setStyleSheet("color: #D0B0E0; background: transparent;")
        self._refresh_hints()
        root.addWidget(self.lbl_hints)

        root.addStretch()

        # Options button
        btn_options = QPushButton("⚙  Options")
        btn_options.setStyleSheet(BTN_OPTIONS_STYLE)
        btn_options.clicked.connect(self._open_options)
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(btn_options)
        btn_row.addStretch()
        root.addLayout(btn_row)

    def _refresh_hints(self):
        hk1 = self.settings.get("hotkey_merge", "ctrl+shift+d").upper()
        hk2 = self.settings.get("hotkey_merge_reverse", "ctrl+shift+f").upper()
        self.lbl_hints.setText(f"{hk1}  ·  fusionner\n{hk2}  ·  fusionner (ordre inversé)")

    def _start_hotkeys(self):
        if self.hotkey_worker:
            self.hotkey_worker.stop()
            self.hotkey_worker.wait(1000)
        self.hotkey_worker = HotkeyWorker(
            self.settings["hotkey_merge"],
            self.settings["hotkey_merge_reverse"],
        )
        self.hotkey_worker.merge_requested.connect(self._on_merge_requested)
        self.hotkey_worker.start()

    def _open_options(self):
        dlg = OptionsDialog(self.settings, self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self.settings = dlg.get_settings()
            save_settings(self.settings)
            self._refresh_hints()
            self._start_hotkeys()

    def _open_info(self):
        InfoDialog(self.settings, self).exec()

    def _on_merge_requested(self, reverse, files):
        if self._progress_toast:
            self._progress_toast.close()
        self._progress_toast = ToastProgress()

        worker = MergeWorker(files, self.settings, reverse)
        worker.finished.connect(self._on_merge_done)
        self._merge_workers.append(worker)
        worker.start()

    def _on_merge_done(self, success, message):
        if self._progress_toast:
            self._progress_toast.close()
            self._progress_toast = None
        self._show_result(message, success)
        self._merge_workers = [w for w in self._merge_workers if w.isRunning()]

    def _show_result(self, message, success=True):
        if self._result_toast:
            self._result_toast.close()
        self._result_toast = ToastResult(message, success)

    def closeEvent(self, event):
        if self.hotkey_worker:
            self.hotkey_worker.stop()
        event.accept()

# ─── Entry Point ─────────────────────────────────────────────────────────────


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    icon = create_icon()
    app.setWindowIcon(icon)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
