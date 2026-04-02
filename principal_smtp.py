import os
import re
import sys
import pytesseract
import tempfile
import traceback
import shutil
import pandas as pd
from pdf2image import convert_from_path
from pypdf import PdfReader, PdfWriter
import fitz  # PyMuPDF
import win32com.client as win32
import json
import ssl
import smtplib
from pathlib import Path
from email.message import EmailMessage

from PIL import Image, ImageDraw, ImageOps


from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QTextEdit, QPushButton,
    QFileDialog, QMessageBox, QHBoxLayout, QVBoxLayout, QPlainTextEdit,
    QCheckBox, QComboBox, QTabWidget
)
from PySide6.QtCore import Qt

# --- CONFIGURACIÓ OCR ---

REGEX_DNI = r"\b(?:\d{7,8}[A-Z]|[XYZ]\d{7}[A-Z])\b"
REGEX_DNI_NOB = r"(?:\d{7,8}[A-Z]|[XYZ]\d{7}[A-Z])"
REGEX_DNI_CANDIDATE = r"\b(?:\d{7,8}[A-Z0-9]|[XYZ]\d{7}[A-Z0-9])\b"
REGEX_DNI_CANDIDATE_NOB = r"(?:\d{7,8}[A-Z0-9]|[XYZ]\d{7}[A-Z0-9])"
DNI_LETTERS = "TRWAGMYFPDXBNJZSQVHLCKE"
NIE_PREFIX_MAP = {"X": "0", "Y": "1", "Z": "2"}

# ForÃ§a la detecciÃ³ per text. Posa-ho a False per tornar a activar l'OCR.
FORCE_TEXT_CLIP_ONLY = False
FORCE_TEXT_ONLY = FORCE_TEXT_CLIP_ONLY
PRE_NORMALIZE_PDF = False
DETECT_MODE_BOTH = "both"
DETECT_MODE_TEXT = "text_clip"
DETECT_MODE_OCR = "ocr"
# Labels to locate the DNI line in text-based PDFs
DNI_LABELS = ["D.N.I.", "DNI", "D.N.I"]

def _find_dni_box(page: "fitz.Page") -> "fitz.Rect | None":
    rects = []
    for lab in DNI_LABELS:
        rects.extend(page.search_for(lab))
    if not rects:
        return None

    page_rect = page.rect
    target_x = page_rect.x0 + page_rect.width * 0.50
    target_y = page_rect.y0 + page_rect.height * 0.10

    def _score(r: "fitz.Rect") -> float:
        cx = (r.x0 + r.x1) / 2.0
        cy = (r.y0 + r.y1) / 2.0
        dx = abs(cx - target_x) / max(1.0, page_rect.width)
        dy = abs(cy - target_y) / max(1.0, page_rect.height)
        return (dy * 2.0) + dx

    # If there are multiple "DNI" labels, pick the one closest to the expected area.
    label = sorted(rects, key=_score)[0]
    gap = label.height * 0.20
    box_w = max(label.width * 3.5, page_rect.width * 0.10)
    box_h = max(label.height * 2.8, page_rect.height * 0.018)
    x0 = label.x0 - (label.width * 0.45)
    y0 = label.y1 + gap
    x1 = x0 + box_w
    y1 = y0 + box_h

    return fitz.Rect(x0, y0, x1, y1)

# OCR config: línea única + whitelist
TESSERACT_CONFIG_DNI = (
    "--oem 3 --psm 7 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
)

def normalize_dni_nie(candidate: str) -> str | None:
    norm = re.sub(r"[^A-Z0-9]", "", (candidate or "").upper())
    if not re.fullmatch(REGEX_DNI_CANDIDATE_NOB, norm):
        return None

    body = norm[:-1]
    if body[0] in NIE_PREFIX_MAP:
        number = NIE_PREFIX_MAP[body[0]] + body[1:]
    else:
        number = body

    if not number.isdigit():
        return None

    letter = DNI_LETTERS[int(number) % 23]
    return body + letter

def extract_dni_candidates(text: str) -> list[str]:
    up = (text or "").upper()
    raw_matches = re.findall(REGEX_DNI_CANDIDATE, up)
    if not raw_matches:
        compact = re.sub(r"[^A-Z0-9]", "", up)
        raw_matches = re.findall(REGEX_DNI_CANDIDATE_NOB, compact)

    out: list[str] = []
    for token in raw_matches:
        dni = normalize_dni_nie(token)
        if dni and dni not in out:
            out.append(dni)
    return out

def ocr_dni_from_crop(img: Image.Image) -> tuple[str, str, str | None]:
    gray = img.convert("L")
    gray = ImageOps.autocontrast(gray)
    raw = pytesseract.image_to_string(gray, lang="spa", config=TESSERACT_CONFIG_DNI).strip()

    if not raw.strip():
        bw = gray.point(lambda x: 0 if x < 200 else 255)
        raw = pytesseract.image_to_string(bw, lang="spa", config=TESSERACT_CONFIG_DNI).strip()

    raw_up = raw.upper()
    norm = re.sub(r"[^A-Z0-9]", "", raw_up)

    cands = extract_dni_candidates(raw)
    if not cands:
        cands = extract_dni_candidates(norm)
    dni = cands[0] if cands else None
    return raw, norm, dni

def pre_normalize_pdf(input_pdf: str) -> str:
    fd, out_pdf = tempfile.mkstemp(prefix="nomines_prenorm_", suffix=".pdf")
    os.close(fd)
    doc = None
    try:
        doc = fitz.open(input_pdf)
        # Reescribe el PDF para normalizar estructura interna.
        doc.save(out_pdf, garbage=4, clean=True, deflate=True)
        return out_pdf
    except Exception:
        try:
            if os.path.exists(out_pdf):
                os.remove(out_pdf)
        except Exception:
            pass
        raise
    finally:
        try:
            if doc is not None:
                doc.close()
        except Exception:
            pass

def _base_dir() -> str:
    # Si está compilado (PyInstaller), los archivos viven en _MEIPASS
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return sys._MEIPASS  # type: ignore[attr-defined]
    # Si estás en desarrollo, usa la carpeta del .py (no el cwd)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = _base_dir()
TOOLS_DIR = os.path.join(BASE_DIR, "tools")

TESSERACT_EXE = os.path.join(TOOLS_DIR, "tesseract", "tesseract.exe")
TESSDATA_DIR  = os.path.join(TOOLS_DIR, "tesseract", "tessdata")
POPPLER_PATH  = os.path.join(TOOLS_DIR, "poppler", "Library", "bin")



pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE
os.environ["TESSDATA_PREFIX"] = TESSDATA_DIR

# Asegura que pdf2image encuentre los .exe/.dll de Poppler
os.environ["PATH"] = POPPLER_PATH + os.pathsep + os.environ.get("PATH", "")

def _check_tools_or_raise():
    missing = []
    if not os.path.isfile(TESSERACT_EXE):
        missing.append(f"Tesseract no encontrado: {TESSERACT_EXE}")
    if not os.path.isdir(TESSDATA_DIR):
        missing.append(f"tessdata no encontrado: {TESSDATA_DIR}")
    if not os.path.isdir(POPPLER_PATH):
        missing.append(f"Poppler bin no encontrado: {POPPLER_PATH}")
    if missing:
        raise FileNotFoundError("\n".join(missing))

# Llama esto al inicio del __init__ o antes de generar_nomines
# _check_tools_or_raise()

def _appdata_config_path() -> Path:
    appdata = os.environ.get("APPDATA")
    if appdata:
        return Path(appdata) / "NominesApp" / "smtp_config.json"
    return Path("smtp_config.json")

def carregar_smtp_config() -> dict:
    cfg_path = _appdata_config_path()
    if not cfg_path.exists():
        cfg_path.parent.mkdir(parents=True, exist_ok=True)
        # crea plantilla para que el usuario la rellene
        plantilla = {
            "smtp_host": "in-v3.mailjet.com",
            "smtp_port": 25,
            "smtp_user": "TU_API_KEY",
            "smtp_pass": "TU_SECRET_KEY",
            "use_starttls": True,
            "from_email": "nomines@tu-dominio.com",
            "from_name": "Nòmines",
        }
        cfg_path.write_text(json.dumps(plantilla, indent=2, ensure_ascii=False), encoding="utf-8")
        raise FileNotFoundError(
            f"No existe la config SMTP.\nHe creat una plantilla a:\n{cfg_path}\n\n"
            "Omple smtp_user/smtp_pass/from_email i torna-ho a provar."
        )

    with cfg_path.open("r", encoding="utf-8") as f:
        return json.load(f)

def enviar_mail_smtp(
    cfg: dict,
    to_email: str,
    subject: str,
    body_text: str,
    attachments: list[str],
):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = f'{cfg.get("from_name","Nomines")} <{cfg["from_email"]}>'
    msg["To"] = to_email
    msg.set_content(body_text)

    for path in attachments:
        with open(path, "rb") as f:
            data = f.read()
        filename = os.path.basename(path)
        msg.add_attachment(data, maintype="application", subtype="pdf", filename=filename)

    host = cfg["smtp_host"]
    port = int(cfg.get("smtp_port", 587))
    use_starttls = bool(cfg.get("use_starttls", True))

    context = ssl.create_default_context()

    def _try_send(p: int):
        with smtplib.SMTP(host, p, timeout=30) as s:
            s.ehlo()
            if use_starttls and p != 465:
                s.starttls(context=context)
                s.ehlo()
            s.login(cfg["smtp_user"], cfg["smtp_pass"])
            s.send_message(msg)

    try:
        _try_send(port)
    except (TimeoutError, OSError, smtplib.SMTPConnectError):
        # Si el 25 está bloqueado, Mailjet recomienda probar 587 :contentReference[oaicite:3]{index=3}
        if port == 25:
            _try_send(587)
        else:
            raise



class NominesApp(QWidget):
    def __init__(self):
        super().__init__()
        _check_tools_or_raise()
        # Estat intern
        self.PDF_NOMINES = ""
        self.EXCEL_TREBALLADORS = ""
        self.SORTIDA_DIR = ""
        self.mapping = {}
        self.pdf_gen = False
        self.errors = []

        self.setWindowTitle("Enviament automàtic de nòmines")

        self._crear_ui()

    # ---------- UTILITARIS D’UI ----------

    def escriure_log(self, msg: str):
        self.log.appendPlainText(msg)
        sb = self.log.verticalScrollBar()
        sb.setValue(sb.maximum())
        QApplication.processEvents()

    def mostrar_error(self, titol: str, missatge: str):
        QMessageBox.critical(self, titol, missatge)

    def netejar_log(self):
        self.log.clear()

    # ---------- CREACIÓ UI ----------

    def _crear_ui(self):
        layout_principal = QVBoxLayout(self)
        tabs = QTabWidget()
        layout_principal.addWidget(tabs)

        tab_main = QWidget()
        layout_principal = QVBoxLayout(tab_main)

        # PDF nòmines
        label_pdf = QLabel("PDF amb nòmines:")
        layout_principal.addWidget(label_pdf, alignment=Qt.AlignLeft)

        layout_pdf = QHBoxLayout()
        self.entry_pdf = QLineEdit()
        btn_pdf = QPushButton("Selecciona")
        btn_pdf.clicked.connect(self.seleccionar_pdf)
        layout_pdf.addWidget(self.entry_pdf)
        layout_pdf.addWidget(btn_pdf)
        layout_principal.addLayout(layout_pdf)

        # Excel treballadors
        label_excel = QLabel("Excel amb DNIs i emails:")
        layout_principal.addWidget(label_excel, alignment=Qt.AlignLeft)

        layout_excel = QHBoxLayout()
        self.entry_excel = QLineEdit()
        btn_excel = QPushButton("Selecciona")
        btn_excel.clicked.connect(self.seleccionar_excel)
        layout_excel.addWidget(self.entry_excel)
        layout_excel.addWidget(btn_excel)
        layout_principal.addLayout(layout_excel)

        # Carpeta sortida
        label_output = QLabel("Carpeta de sortida:")
        layout_principal.addWidget(label_output, alignment=Qt.AlignLeft)

        layout_output = QHBoxLayout()
        self.entry_output = QLineEdit()
        btn_output = QPushButton("Selecciona")
        btn_output.clicked.connect(self.seleccionar_carpeta)
        layout_output.addWidget(self.entry_output)
        layout_output.addWidget(btn_output)
        layout_principal.addLayout(layout_output)

        # Assumpte
        label_subject = QLabel("Assumpte:")
        layout_principal.addWidget(label_subject, alignment=Qt.AlignLeft)

        self.entry_subject = QLineEdit()
        layout_principal.addWidget(self.entry_subject)

        # Cos del missatge
        label_body = QLabel("Cos del missatge:")
        layout_principal.addWidget(label_body, alignment=Qt.AlignLeft)

        self.text_body = QTextEdit()
        layout_principal.addWidget(self.text_body)

        # Mes (1-12)
        label_mes = QLabel("Mes (1-12):")
        layout_principal.addWidget(label_mes, alignment=Qt.AlignLeft)

        layout_mes = QHBoxLayout()
        self.combo_mes = QComboBox()
        self.combo_mes.addItem("Selecciona el mes")
        for i in range(1, 13):
            self.combo_mes.addItem(str(i))
        layout_mes.addWidget(self.combo_mes)
        layout_mes.addStretch()
        layout_principal.addLayout(layout_mes)

        # Botons accions
        layout_btn = QHBoxLayout()
        btn_generar = QPushButton("Generar PDFs")
        btn_generar.setStyleSheet("background-color: lightyellow;")
        btn_generar.clicked.connect(self.generar_nomines)

        btn_enviar = QPushButton("Enviar correus")
        btn_enviar.setStyleSheet("background-color: lightgreen;")
        btn_enviar.clicked.connect(self.enviar_correus)

        layout_btn.addWidget(btn_generar)
        layout_btn.addWidget(btn_enviar)
        layout_principal.addLayout(layout_btn)

        # Log + botó netejar
        layout_log_header = QHBoxLayout()
        label_log = QLabel("Log:")
        layout_log_header.addWidget(label_log, alignment=Qt.AlignLeft)

        btn_neteja_log = QPushButton("Neteja log")
        btn_neteja_log.clicked.connect(self.netejar_log)
        layout_log_header.addStretch()
        layout_log_header.addWidget(btn_neteja_log)

        layout_principal.addLayout(layout_log_header)

        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        layout_principal.addWidget(self.log)

        tabs.addTab(tab_main, "Principal")

        tab_opcions = QWidget()
        layout_opcions = QVBoxLayout(tab_opcions)

        self.chk_enviar_directament = QCheckBox(
            "Enviar directament (si es desmarca, només es creen esborranys)"
        )
        self.chk_enviar_directament.setChecked(True)
        layout_opcions.addWidget(self.chk_enviar_directament)

        self.chk_debug_crops = QCheckBox(
            "Guardar imatges de debug (_debug_crops)"
        )
        self.chk_debug_crops.setChecked(False)
        layout_opcions.addWidget(self.chk_debug_crops)

        label_detect = QLabel("Mode detecció DNI:")
        layout_opcions.addWidget(label_detect, alignment=Qt.AlignLeft)

        self.combo_detect_mode = QComboBox()
        self.combo_detect_mode.addItem("Ambdós (text-clip primer)", DETECT_MODE_BOTH)
        self.combo_detect_mode.addItem("Només text-clip", DETECT_MODE_TEXT)
        self.combo_detect_mode.addItem("Només OCR", DETECT_MODE_OCR)
        self.combo_detect_mode.setCurrentIndex(1 if FORCE_TEXT_CLIP_ONLY else 0)
        layout_opcions.addWidget(self.combo_detect_mode)
        layout_opcions.addStretch()

        tabs.addTab(tab_opcions, "Opcions")

        tab_instruccions = QWidget()
        layout_instr = QVBoxLayout(tab_instruccions)
        instr = QTextEdit()
        instr.setReadOnly(True)
        instr.setPlainText(
            "Instruccions d'ús:\n"
            "1) Selecciona el PDF amb les nòmines.\n"
            "2) Selecciona l'Excel amb els DNIs/NIEs i els emails.\n"
            "3) Selecciona la carpeta de sortida.\n"
            "4) Escriu l'assumpte i el cos del missatge.\n"
            "5) Selecciona el mes.\n"
            "6) Fes clic a “Generar PDFs” i revisa el log i l'arxiu d'errors.\n"
            "7) Si tot és correcte, fes clic a “Enviar correus”.\n"
            "\n"
            "Notes:\n"
            "- Si no existeix la configuració SMTP, es crearà una plantilla.\n"
            "- Els PDFs es protegeixen amb el DNI/NIE com a contrasenya.\n"
            "- Si hi ha DNIs no trobats, revisa l'Excel i el fitxer d'errors.\n"
            "- Pots activar/desactivar el guardat d'imatges de debug amb el checkbox.\n"
            "- Pots enviar correus sempre i quan tinguis els PDFs generats a la carpeta de sortida."
        )
        layout_instr.addWidget(instr)
        tabs.addTab(tab_instruccions, "Instruccions")

    # ---------- FUNCIONS D’ENTRADA / DIÀLEGS ----------

    def seleccionar_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Selecciona el PDF amb les nòmines",
            "",
            "Fitxers PDF (*.pdf)"
        )
        if file_path:
            self.PDF_NOMINES = file_path
            self.entry_pdf.setText(file_path)

    def seleccionar_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Selecciona l'Excel amb DNIs i emails",
            "",
            "Fitxers Excel (*.xls *.xlsx)"
        )
        if not file_path:
            return

        self.EXCEL_TREBALLADORS = file_path
        self.entry_excel.setText(file_path)

        try:
            df = pd.read_excel(self.EXCEL_TREBALLADORS)
        except Exception as e:
            self.mostrar_error("Error", f"No s'ha pogut llegir l'Excel:\n{e}")
            return

        for col in ("DNI", "Email"):
            if col not in df.columns:
                self.mostrar_error("Error", f"Falta la columna '{col}' a l'Excel.")
                return

        df["DNI"] = df["DNI"].astype(str).str.strip().str.upper()
        df["Email"] = df["Email"].astype(str).str.strip()
        self.mapping = dict(zip(df["DNI"], df["Email"]))

    def seleccionar_carpeta(self):
        dir_path = QFileDialog.getExistingDirectory(
            self,
            "Selecciona la carpeta on guardar les nòmines individuals",
            ""
        )
        if dir_path:
            self.SORTIDA_DIR = dir_path
            self.entry_output.setText(dir_path)

    def obtenir_assumpte_i_cos(self):
        assumpte = self.entry_subject.text().strip()
        cos = self.text_body.toPlainText().strip()
        if not assumpte or not cos:
            self.mostrar_error("Error", "Has d'escriure un assumpte i un cos del correu.")
            return None, None
        return assumpte, cos

    def obtenir_mes(self):
        idx = self.combo_mes.currentIndex()
        if idx <= 0:
            self.mostrar_error("Error", "Has de seleccionar el mes (1-12).")
            return None
        try:
            mes = int(self.combo_mes.currentText())
        except ValueError:
            self.mostrar_error("Error", "El mes seleccionat no és vàlid.")
            return None
        return mes

    # ---------- LÒGICA PRINCIPAL ----------

    def obtenir_mode_deteccio(self) -> str:
        mode = self.combo_detect_mode.currentData()
        if mode in {DETECT_MODE_BOTH, DETECT_MODE_TEXT, DETECT_MODE_OCR}:
            return mode
        return DETECT_MODE_BOTH

    def generar_nomines(self):
        if not self.PDF_NOMINES or not self.EXCEL_TREBALLADORS or not self.SORTIDA_DIR:
            self.mostrar_error("Error", "Has de seleccionar PDF, Excel i carpeta de sortida.")
            return

        assumpte, cos = self.obtenir_assumpte_i_cos()
        if not assumpte:
            return

        self.errors = []
        self.pdf_gen = False
        detect_mode = self.obtenir_mode_deteccio()
        if detect_mode == DETECT_MODE_TEXT:
            self.escriure_log("[Mode] nomes text-clip")
        elif detect_mode == DETECT_MODE_OCR:
            self.escriure_log("[Mode] nomes OCR")
        else:
            self.escriure_log("[Mode] ambdos (text-clip primer)")

        self.escriure_log("➡️ Generant PDFs individuals...")

        os.makedirs(self.SORTIDA_DIR, exist_ok=True)
        pdf_to_process = self.PDF_NOMINES
        pre_norm_pdf = None
        if PRE_NORMALIZE_PDF:
            try:
                pre_norm_pdf = pre_normalize_pdf(self.PDF_NOMINES)
                pdf_to_process = pre_norm_pdf
                self.escriure_log(f"[PDF] Pre-normalitzat: {os.path.basename(pre_norm_pdf)}")
            except Exception as e:
                self.escriure_log(f"[PDF] Pre-normalitzacio fallida ({e}). Es continua amb l'original.")

        try:
            reader = PdfReader(pdf_to_process)
            pages_img = convert_from_path(pdf_to_process, dpi=300, poppler_path=POPPLER_PATH)
            fitz_doc = fitz.open(pdf_to_process)
        except Exception as e:
            if pre_norm_pdf and os.path.exists(pre_norm_pdf):
                try:
                    os.remove(pre_norm_pdf)
                except Exception:
                    pass
            self.mostrar_error("Error", f"No s'ha pogut processar el PDF:\n{e}")
            return

        for i, (page_img, page_pdf) in enumerate(zip(pages_img, reader.pages)):
            w, h = page_img.size
            rect = None
            clip = None

            source_tag = "fixed"
            clip_pdf_str = "fixed_pct(0.53-0.62,0.100-0.110)"
            clip_px_str = ""
            try:
                fitz_page = fitz_doc.load_page(i)
                rect = fitz_page.rect
                margin_w = rect.width * 0.005
                margin_h = rect.height * 0.003
                x0 = rect.x0 + rect.width * 0.53 - margin_w
                y0 = rect.y0 + rect.height * 0.100 - margin_h
                x1 = rect.x0 + rect.width * 0.62 + margin_w
                y1 = rect.y0 + rect.height * 0.110 + margin_h
                clip = fitz.Rect(x0, y0, x1, y1)
                # Clamp to page bounds
                clip = fitz.Rect(
                    max(rect.x0, clip.x0),
                    max(rect.y0, clip.y0),
                    min(rect.x1, clip.x1),
                    min(rect.y1, clip.y1),
                )
                clip_pdf_str = f"({clip.x0:.2f}, {clip.y0:.2f}, {clip.x1:.2f}, {clip.y1:.2f})"
                direct_text = fitz_page.get_text("text", clip=clip) or ""
                print(
                    f"[DEBUG] Pagina {i+1} fixed_clip={clip} text='{direct_text[:80]}'"
                )
                if not direct_text.strip():
                    page_text = fitz_page.get_text("text") or ""
                    print(f"[DEBUG] Pagina {i+1} text_layer_len={len(page_text)}")
            except Exception:
                direct_text = ""
                clip = None
                source_tag = "fixed_px"

            if clip is not None and rect is not None:
                scale_x = w / rect.width
                scale_y = h / rect.height
                px0 = int(clip.x0 * scale_x)
                py0 = int(clip.y0 * scale_y)
                px1 = int(clip.x1 * scale_x)
                py1 = int(clip.y1 * scale_y)
                clip_px_str = f"({px0}, {py0}, {px1}, {py1})"
                print(f"[DEBUG] Pagina {i+1} clip_px={clip_px_str}")
            else:
                px0 = int(w * 0.53)
                py0 = int(h * 0.100)
                px1 = int(w * 0.62)
                py1 = int(h * 0.110)

            px0 = max(0, min(px0, w - 1))
            py0 = max(0, min(py0, h - 1))
            px1 = max(px0 + 1, min(px1, w))
            py1 = max(py0 + 1, min(py1, h))
            dni_zone = page_img.crop((px0, py0, px1, py1))

            if self.chk_debug_crops.isChecked():
                debug_dir = os.path.join(self.SORTIDA_DIR, "_debug_crops")
                os.makedirs(debug_dir, exist_ok=True)
                #self.escriure_log(
                    #f"[DEBUG] Pagina {i+1} source={source_tag} clip_pdf={clip_pdf_str} clip_px={clip_px_str}"
                #)
                dni_zone.save(os.path.join(debug_dir, f"page_{i+1}_dni_zone_{source_tag}.png"))
                debug_page = page_img.copy()
                draw = ImageDraw.Draw(debug_page)
                draw.rectangle([px0, py0, px1, py1], outline="red", width=3)
                debug_page.save(os.path.join(debug_dir, f"page_{i+1}_full_{source_tag}.png"))


            source_label = "text-clip"
            text = direct_text
            dni = None
            if detect_mode in {DETECT_MODE_BOTH, DETECT_MODE_TEXT}:
                direct_candidates = extract_dni_candidates(direct_text or "")
                direct_valid = [d for d in direct_candidates if d in self.mapping]
                if direct_valid:
                    dni = direct_valid[0]
                elif direct_candidates:
                    dni = direct_candidates[0]

            if not dni and detect_mode in {DETECT_MODE_BOTH, DETECT_MODE_OCR}:
                try:
                    cw, ch = dni_zone.size
                    # Primer intent: OCR sobre todo el recorte (debug ya ajustado al DNI).
                    raw, norm, ocr_dni = ocr_dni_from_crop(dni_zone)
                    if not ocr_dni:
                        # Segundo intento: recorte vertical suave para quitar margenes.
                        inner = dni_zone.crop((0, int(ch * 0.10), cw, int(ch * 0.90)))
                        raw, norm, ocr_dni = ocr_dni_from_crop(inner)
                    ocr_candidates = extract_dni_candidates(raw) or extract_dni_candidates(norm)
                    ocr_valid = [d for d in ocr_candidates if d in self.mapping]
                    if ocr_valid:
                        dni = ocr_valid[0]
                    elif ocr_candidates:
                        dni = ocr_candidates[0]
                    else:
                        dni = ocr_dni
                    text = raw
                except Exception as e:
                    self.escriure_log(f"[Pagina {i+1}] Error OCR: {e}")
                    continue

                source_label = "OCR"
                if dni is None:
                    self.escriure_log(f"[DNI OCR] p={i+1} raw={raw!r} norm={norm!r}")

            if not dni:
                self.escriure_log(f"[Pagina {i+1}] No s'ha trobat DNI ({source_label}: {text})")
                continue

            dni = dni.upper()
            self.escriure_log(f"[Pagina {i+1}] ID trobat per {source_label}")
            if dni not in self.mapping:
                self.errors.append((i + 1, dni, text))
                self.escriure_log(f"[Pàgina {i+1}] ⚠️ DNI {dni} no trobat a l'Excel")
                continue

            email = self.mapping[dni]
            fitxer_nomina = os.path.abspath(os.path.join(self.SORTIDA_DIR, f"{dni}.pdf"))

            writer = PdfWriter()
            writer.add_page(page_pdf)

            # 🔐 Protegir el PDF amb contrasenya = DNI
            try:
                writer.encrypt(dni)
            except Exception as e:
                self.escriure_log(f"[Pàgina {i+1}] ⚠️ No s'ha pogut protegir amb contrasenya: {e}")

            with open(fitxer_nomina, "wb") as f:
                writer.write(f)

            self.escriure_log(f"[Pàgina {i+1}] ✅ Nòmina per {email} protegida amb contrasenya ({dni})")

        try:
            fitz_doc.close()
        except Exception:
            pass
        if pre_norm_pdf and os.path.exists(pre_norm_pdf):
            try:
                os.remove(pre_norm_pdf)
            except Exception:
                pass

        if self.errors:
            errors_path = os.path.join(self.SORTIDA_DIR, "errors_nomines.xlsx")
            pd.DataFrame(self.errors, columns=["Pàgina", "DNI detectat", "OCR brut"]).to_excel(errors_path, index=False)
            self.escriure_log(f"⚠️ Hi ha DNIs no trobats, revisa '{errors_path}'")

        self.escriure_log("✅ Generació completada. Ara pots enviar els correus.")
        self.pdf_gen = True

    def enviar_correus(self):
        # --- Validacions bàsiques ---
        if not self.SORTIDA_DIR or not self.mapping:
            self.mostrar_error(
                "Error",
                "Has de seleccionar una carpeta de sortida i un Excel amb DNIs/emails."
            )
            return

        pdfs = [f for f in os.listdir(self.SORTIDA_DIR) if f.lower().endswith(".pdf")]
        if not pdfs:
            self.mostrar_error("Error", "No hi ha cap PDF a la carpeta de sortida.")
            return

        assumpte, cos = self.obtenir_assumpte_i_cos()
        if not assumpte:
            return

        mes = self.obtenir_mes()
        if mes is None:
            return

        # Amb SMTP només té sentit enviar directament
        enviar_directament = self.chk_enviar_directament.isChecked()
        if not enviar_directament:
            self.mostrar_error(
                "Error",
                "En mode SMTP (Mailjet) no es poden crear esborranys.\n"
                "Marca 'Enviar directament'."
            )
            return

        # --- Carregar config SMTP només des de AppData ---
        def smtp_config_path() -> Path:
            appdata = os.environ.get("APPDATA")
            if not appdata:
                raise RuntimeError("No s'ha trobat la variable d'entorn APPDATA.")
            return Path(appdata) / "NominesApp" / "smtp_config.json"

        cfg_path = smtp_config_path()
        if not cfg_path.exists():
            self.mostrar_error(
                "Error",
                f"No existeix el fitxer de configuració SMTP:\n{cfg_path}\n\n"
                "Crea'l i enganxa-hi el JSON de Mailjet."
            )
            return

        try:
            cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
        except Exception as e:
            self.mostrar_error("Error", f"No es pot llegir el JSON:\n{cfg_path}\n\n{e}")
            return

        # Log per evitar confusions (això et dirà QUIN config usa)
        self.escriure_log(f"SMTP config: {cfg_path}")
        self.escriure_log(f"SMTP host/port: {cfg.get('smtp_host')}:{cfg.get('smtp_port')}")
        self.escriure_log(f"From: {cfg.get('from_name')} <{cfg.get('from_email')}>")

        # --- Preparar nom adjunt segons mes ---
        noms_mes = {
            1: "gener", 2: "febrer", 3: "març", 4: "abril",
            5: "maig", 6: "juny", 7: "juliol", 8: "agost",
            9: "setembre", 10: "octubre", 11: "novembre", 12: "desembre"
        }
        nom_mes = noms_mes.get(mes, str(mes))
        nom_adjunt = f"Nòmina {nom_mes}.pdf"

        # --- Enviament SMTP ---
        def enviar_mail_smtp(to_email: str, subject: str, body_text: str, pdf_path: str):
            msg = EmailMessage()
            msg["Subject"] = subject
            msg["From"] = f"{cfg.get('from_name','Nòmines')} <{cfg['from_email']}>"
            msg["To"] = to_email
            msg.set_content(body_text)

            with open(pdf_path, "rb") as f:
                data = f.read()
            msg.add_attachment(
                data,
                maintype="application",
                subtype="pdf",
                filename=os.path.basename(pdf_path)
            )

            host = cfg["smtp_host"]
            port = int(cfg.get("smtp_port", 587))
            user = cfg["smtp_user"]
            pwd = cfg["smtp_pass"]
            use_starttls = bool(cfg.get("use_starttls", True))

            context = ssl.create_default_context()
            with smtplib.SMTP(host, port, timeout=30) as s:
                # DEBUG SMTP: activa-ho si vols veure el diàleg SMTP al terminal
                # s.set_debuglevel(1)
                s.ehlo()
                if use_starttls and port != 465:
                    s.starttls(context=context)
                    s.ehlo()
                s.login(user, pwd)
                s.send_message(msg)

        self.escriure_log("➡️ Enviant correus per SMTP (Mailjet)...")

        temp_dir = tempfile.mkdtemp(prefix="nomines_")
        try:
            for fitxer in os.listdir(self.SORTIDA_DIR):
                if not fitxer.lower().endswith(".pdf"):
                    continue

                dni = os.path.splitext(fitxer)[0]
                email_destinatari = self.mapping.get(dni)

                if not email_destinatari:
                    self.escriure_log(f"❌ No hi ha correu per {dni}, no s'envia")
                    continue

                original_pdf = os.path.abspath(os.path.join(self.SORTIDA_DIR, fitxer))
                temp_pdf = os.path.join(temp_dir, nom_adjunt)

                try:
                    shutil.copyfile(original_pdf, temp_pdf)
                except Exception as e:
                    self.escriure_log(f"❌ Error copiant PDF ({fitxer}): {e}")
                    continue

                try:
                    enviar_mail_smtp(
                        to_email=email_destinatari,
                        subject=assumpte,
                        body_text=cos,
                        pdf_path=temp_pdf
                    )
                    self.escriure_log(f"📧 Enviat (SMTP): {nom_adjunt} a {email_destinatari} (origen {fitxer})")

                except Exception as e:
                    self.escriure_log(f"❌ Error SMTP enviant a {email_destinatari}: {repr(e)}")
                    self.escriure_log(traceback.format_exc())

        finally:
            try:
                shutil.rmtree(temp_dir)
            except Exception:
                pass

        self.escriure_log("✅ Procés d'enviament completat.")




if __name__ == "__main__":
    app = QApplication(sys.argv)
    finestra = NominesApp()
    finestra.resize(800, 700)
    finestra.show()
    sys.exit(app.exec())



