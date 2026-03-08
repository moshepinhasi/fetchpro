"""
FetchPro — מנהל הורדות מקצועי v5.5
תכונות: הורדה מקוטעת, Resume, תור עדיפויות, throttle, hash,
        חילוץ אוטומטי, System Tray, התראות, היסטוריה SQLite,
        Drag & Drop, yt-dlp, FTP, תזמון, ייבוא/ייצוא, נושא כהה/בהיר,
        Auto-retry, bandwidth limit כולל, חיפוש היסטוריה, שורת סטטוס,
        BitTorrent, YouTube/מדיה עם בחירת איכות, הורדת מוזיקה MP3,
        Persistent Queue, Proxy/SOCKS5, תזמון טווח שעות,
        כיבוי אחרי תור, בדיקת דיסק, Watchdog, REST API,
        סטטיסטיקות שימוש, Tags להורדות

Requirements (core):
    pip install requests
Optional (enables extra features):
    pip install pystray pillow plyer yt-dlp libtorrent
"""

from __future__ import annotations

# ── stdlib ────────────────────────────────────────────────────────────────────
import collections
import ftplib
import hashlib
import http.server
import json
import logging
import os
import platform
import re
import shutil
import sqlite3
import subprocess
import sys
import tarfile
import threading
import time
import tkinter as tk
import tkinter.filedialog as filedialog
import tkinter.messagebox as messagebox
import tkinter.ttk as ttk
import zipfile
from collections.abc import Callable
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field, asdict
from datetime import datetime, timedelta
from enum import Enum
from pathlib import Path
from urllib.parse import unquote, urlparse

# ── Optional deps ─────────────────────────────────────────────────────────────
try:
    import requests
    REQUESTS_OK = True
except ModuleNotFoundError:
    REQUESTS_OK = False

try:
    import pystray
    from PIL import Image as PilImage
    TRAY_OK = True
except (ImportError, ModuleNotFoundError):
    TRAY_OK = False

try:
    from plyer import notification as plyer_notif
    PLYER_OK = True
except (ImportError, ModuleNotFoundError):
    PLYER_OK = False

try:
    import yt_dlp
    YTDLP_OK = True
except (ImportError, ModuleNotFoundError):
    YTDLP_OK = False

try:
    import libtorrent as lt
    LIBTORRENT_OK = True
except (ImportError, ModuleNotFoundError):
    lt = None          # type: ignore[assignment]
    LIBTORRENT_OK = False

# ── Startup check ─────────────────────────────────────────────────────────────
if not REQUESTS_OK:
    _r = tk.Tk(); _r.withdraw()
    messagebox.showerror("חסרה תלות",
                         "הספרייה 'requests' לא מותקנת.\n\nהרץ:\n  pip install requests")
    sys.exit(1)

# ── Logging ───────────────────────────────────────────────────────────────────
import logging.handlers as _log_handlers

logging.basicConfig(level=logging.WARNING,
                    format="%(asctime)s %(levelname)s %(name)s: %(message)s")
logger = logging.getLogger(__name__)

def _setup_file_log() -> None:
    """Add rotating file handler after APP_DIR is known."""
    try:
        log_path = Path.home() / ".fetchpro" / "fetchpro.log"
        log_path.parent.mkdir(parents=True, exist_ok=True)
        fh = _log_handlers.RotatingFileHandler(
            log_path, maxBytes=2_000_000, backupCount=3, encoding="utf-8")
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
        logging.getLogger().addHandler(fh)
        logging.getLogger().setLevel(logging.DEBUG)
    except (OSError, PermissionError):
        pass

# ---------------------------------------------------------------------------
# Theme engine — dark / light
# ---------------------------------------------------------------------------
_DARK = dict(
    BG_DEEP="#070b12", BG_CARD="#0d1117", BG_HOVER="#111827",
    BORDER="#1e2530", ACCENT="#00e5ff", ACCENT_BG="#0a2a33",
    GREEN="#00e676", GREEN_DIM="#1a4a1a", RED="#ff1744",
    YELLOW="#ffd740", TEXT_MAIN="#e2e8f0", TEXT_DIM="#4a5568",
    TEXT_MUTED="#2d3748",
)
_LIGHT = dict(
    BG_DEEP="#f0f4f8", BG_CARD="#ffffff", BG_HOVER="#e2e8f0",
    BORDER="#cbd5e0", ACCENT="#0077aa", ACCENT_BG="#dbeafe",
    GREEN="#2f855a", GREEN_DIM="#c6f6d5", RED="#c53030",
    YELLOW="#b7791f", TEXT_MAIN="#1a202c", TEXT_DIM="#718096",
    TEXT_MUTED="#a0aec0",
)

class Theme:
    """Mutable singleton — swap palettes at runtime."""
    _p = _DARK

    @classmethod
    def toggle(cls) -> None:
        cls._p = _LIGHT if cls._p is _DARK else _DARK

    @classmethod
    def is_dark(cls) -> bool:
        return cls._p is _DARK

    def __getattr__(self, name: str) -> str:
        try:
            return Theme._p[name]
        except KeyError:
            raise AttributeError(name)

T = Theme()

# ---------------------------------------------------------------------------
# PyInstaller / EXE compatibility
# ---------------------------------------------------------------------------
def _get_exe_path() -> Path:
    """Return the real path to this process's executable (works for .py and .exe)."""
    if getattr(sys, "frozen", False):
        # Running as PyInstaller EXE
        return Path(sys.executable)
    return Path(sys.argv[0]).resolve()

EXE_PATH = _get_exe_path()

# ---------------------------------------------------------------------------
# Internationalisation (i18n) — CLEAN, no circular refs, no duplicates
# ---------------------------------------------------------------------------
_STRINGS: dict[str, dict[str, str]] = {
    # ── App chrome ──────────────────────────────────────────────────────────
    "app_title":          {"he":"מנהל הורדות",          "en":"Download Manager",               "ar":"مدير التنزيلات",       "ru":"Менеджер загрузок",    "es":"Gestor de descargas",          "fr":"Gestionnaire de téléchargements"},
    "quit":               {"he":"✕ צא",                 "en":"✕ Quit",                         "ar":"✕ خروج",               "ru":"✕ Выход",              "es":"✕ Salir",                      "fr":"✕ Quitter"},
    "settings":           {"he":"⚙ הגדרות",             "en":"⚙ Settings",                     "ar":"⚙ الإعدادات",          "ru":"⚙ Настройки",          "es":"⚙ Ajustes",                    "fr":"⚙ Paramètres"},
    "history":            {"he":"📜 היסטוריה",           "en":"📜 History",                     "ar":"📜 السجل",              "ru":"📜 История",            "es":"📜 Historial",                  "fr":"📜 Historique"},
    "export":             {"he":"📤 ייצא",               "en":"📤 Export",                      "ar":"📤 تصدير",              "ru":"📤 Экспорт",            "es":"📤 Exportar",                   "fr":"📤 Exporter"},
    "import_btn":         {"he":"📥 ייבוא",              "en":"📥 Import",                      "ar":"📥 استيراد",            "ru":"📥 Импорт",             "es":"📥 Importar",                   "fr":"📥 Importer"},
    "retry_all":          {"he":"↺ נסה שנית הכל",        "en":"↺ Retry All",                   "ar":"↺ إعادة المحاولة",     "ru":"↺ Повторить всё",      "es":"↺ Reintentar todo",             "fr":"↺ Tout relancer"},
    "stats":              {"he":"📊 סטטיסטיקות",         "en":"📊 Statistics",                  "ar":"📊 الإحصائيات",         "ru":"📊 Статистика",         "es":"📊 Estadísticas",               "fr":"📊 Statistiques"},
    "media":              {"he":"🎬 מדיה",               "en":"🎬 Media",                       "ar":"🎬 وسائط",              "ru":"🎬 Медиа",              "es":"🎬 Media",                      "fr":"🎬 Médias"},
    "theme":              {"he":"🌗 נושא",               "en":"🌗 Theme",                       "ar":"🌗 السمة",              "ru":"🌗 Тема",               "es":"🌗 Tema",                       "fr":"🌗 Thème"},
    "torrent":            {"he":"🧲 טורנט",              "en":"🧲 Torrent",                     "ar":"🧲 تورنت",              "ru":"🧲 Торрент",            "es":"🧲 Torrent",                    "fr":"🧲 Torrent"},

    # ── Download actions ─────────────────────────────────────────────────────
    "download":           {"he":"⬇ הורד",                "en":"⬇ Download",                    "ar":"⬇ تنزيل",              "ru":"⬇ Скачать",            "es":"⬇ Descargar",                  "fr":"⬇ Télécharger"},
    "pause":              {"he":"⏸ השהה",                "en":"⏸ Pause",                       "ar":"⏸ إيقاف مؤقت",        "ru":"⏸ Пауза",              "es":"⏸ Pausar",                     "fr":"⏸ Pause"},
    "resume":             {"he":"▶ המשך",                "en":"▶ Resume",                      "ar":"▶ استئناف",            "ru":"▶ Продолжить",         "es":"▶ Reanudar",                   "fr":"▶ Reprendre"},
    "cancel":             {"he":"✕ בטל",                 "en":"✕ Cancel",                      "ar":"✕ إلغاء",              "ru":"✕ Отмена",             "es":"✕ Cancelar",                   "fr":"✕ Annuler"},
    "retry":              {"he":"↺ נסה שנית",             "en":"↺ Retry",                       "ar":"↺ أعد المحاولة",       "ru":"↺ Повторить",          "es":"↺ Reintentar",                 "fr":"↺ Réessayer"},
    "open_file":          {"he":"📂 פתח",                "en":"📂 Open",                        "ar":"📂 فتح",               "ru":"📂 Открыть",           "es":"📂 Abrir",                      "fr":"📂 Ouvrir"},
    "open_folder":        {"he":"📁 תיקייה",             "en":"📁 Folder",                      "ar":"📁 المجلد",             "ru":"📁 Папка",             "es":"📁 Carpeta",                    "fr":"📁 Dossier"},
    "delete":             {"he":"🗑 מחק",                "en":"🗑 Delete",                      "ar":"🗑 حذف",               "ru":"🗑 Удалить",           "es":"🗑 Eliminar",                   "fr":"🗑 Supprimer"},
    "copy_url":           {"he":"🔗 העתק URL",           "en":"🔗 Copy URL",                   "ar":"🔗 نسخ الرابط",         "ru":"🔗 Скопировать URL",   "es":"🔗 Copiar URL",                "fr":"🔗 Copier l'URL"},
    "apply":              {"he":"✓ החל",                 "en":"✓ Apply",                       "ar":"✓ تطبيق",              "ru":"✓ Применить",          "es":"✓ Aplicar",                    "fr":"✓ Appliquer"},
    "cancel_btn":         {"he":"ביטול",                 "en":"Cancel",                        "ar":"إلغاء",                "ru":"Отмена",               "es":"Cancelar",                     "fr":"Annuler"},
    "save":               {"he":"💾 שמור",               "en":"💾 Save",                        "ar":"💾 حفظ",               "ru":"💾 Сохранить",         "es":"💾 Guardar",                    "fr":"💾 Enregistrer"},
    "close":              {"he":"סגור",                  "en":"Close",                         "ar":"إغلاق",                "ru":"Закрыть",              "es":"Cerrar",                        "fr":"Fermer"},
    "add_download":       {"he":"⬇ הוסף הורדה",          "en":"⬇ Add Download",                "ar":"⬇ إضافة تنزيل",        "ru":"⬇ Добавить загрузку",  "es":"⬇ Agregar descarga",           "fr":"⬇ Ajouter un téléchargement"},
    "verify_url":         {"he":"🔍 בדוק URL",           "en":"🔍 Verify URL",                  "ar":"🔍 تحقق من الرابط",    "ru":"🔍 Проверить URL",     "es":"🔍 Verificar URL",              "fr":"🔍 Vérifier l'URL"},
    "clear":              {"he":"✕ נקה",                 "en":"✕ Clear",                       "ar":"✕ مسح",                "ru":"✕ Очистить",           "es":"✕ Limpiar",                    "fr":"✕ Effacer"},

    # ── Context menu ─────────────────────────────────────────────────────────
    "ctx_copy_url":       {"he":"העתק URL",              "en":"Copy URL",                      "ar":"نسخ الرابط",           "ru":"Копировать URL",       "es":"Copiar URL",                   "fr":"Copier l'URL"},
    "ctx_open_folder":    {"he":"פתח תיקיית יעד",        "en":"Open destination folder",       "ar":"فتح مجلد الوجهة",      "ru":"Открыть папку",        "es":"Abrir carpeta destino",         "fr":"Ouvrir le dossier"},
    "ctx_open_file":      {"he":"פתח קובץ",              "en":"Open file",                     "ar":"فتح الملف",             "ru":"Открыть файл",         "es":"Abrir archivo",                "fr":"Ouvrir le fichier"},
    "ctx_cancel":         {"he":"בטל הורדה",             "en":"Cancel download",               "ar":"إلغاء التنزيل",        "ru":"Отменить загрузку",    "es":"Cancelar descarga",             "fr":"Annuler le téléchargement"},
    "ctx_retry":          {"he":"נסה שנית",              "en":"Retry",                         "ar":"أعد المحاولة",         "ru":"Повторить",            "es":"Reintentar",                   "fr":"Réessayer"},
    "ctx_delete_partial": {"he":"מחק קובץ חלקי",         "en":"Delete partial file",           "ar":"حذف الملف الجزئي",     "ru":"Удалить частичный файл","es":"Eliminar archivo parcial",     "fr":"Supprimer fichier partiel"},
    "ctx_remove":         {"he":"הסר מהרשימה",           "en":"Remove from list",              "ar":"إزالة من القائمة",     "ru":"Удалить из списка",    "es":"Eliminar de la lista",          "fr":"Retirer de la liste"},
    "ctx_open_folder2":   {"he":"פתח תיקייה",            "en":"Open folder",                   "ar":"فتح المجلد",            "ru":"Открыть папку",        "es":"Abrir carpeta",                "fr":"Ouvrir le dossier"},
    "ctx_delete_record":  {"he":"מחק רשומה",             "en":"Delete record",                 "ar":"حذف السجل",             "ru":"Удалить запись",       "es":"Eliminar registro",             "fr":"Supprimer l'enregistrement"},

    # ── Clipboard / text editing ─────────────────────────────────────────────
    "paste":              {"he":"📋 הדבק",               "en":"📋 Paste",                       "ar":"📋 لصق",               "ru":"📋 Вставить",          "es":"📋 Pegar",                      "fr":"📋 Coller"},
    "cut":                {"he":"✂ גזור",                "en":"✂ Cut",                         "ar":"✂ قص",                "ru":"✂ Вырезать",           "es":"✂ Cortar",                     "fr":"✂ Couper"},
    "copy_action":        {"he":"📄 העתק",               "en":"📄 Copy",                        "ar":"📄 نسخ",               "ru":"📄 Копировать",        "es":"📄 Copiar",                     "fr":"📄 Copier"},
    "paste_from_clip":    {"he":"📋 הדבק מלוח הגזירים", "en":"📋 Paste from clipboard",        "ar":"📋 لصق من الحافظة",    "ru":"📋 Из буфера обмена", "es":"📋 Pegar del portapapeles",      "fr":"📋 Coller depuis le presse-papiers"},
    "select_all":         {"he":"בחר הכל",               "en":"Select all",                    "ar":"تحديد الكل",           "ru":"Выбрать всё",          "es":"Seleccionar todo",              "fr":"Tout sélectionner"},
    "clear_all":          {"he":"נקה הכל",               "en":"Clear all",                     "ar":"مسح الكل",             "ru":"Очистить всё",         "es":"Limpiar todo",                  "fr":"Tout effacer"},
    "paste_short":        {"he":"הדבק  Ctrl+V",          "en":"Paste  Ctrl+V",                 "ar":"لصق  Ctrl+V",          "ru":"Вставить  Ctrl+V",     "es":"Pegar  Ctrl+V",                "fr":"Coller  Ctrl+V"},
    "copy_short":         {"he":"העתק  Ctrl+C",          "en":"Copy  Ctrl+C",                  "ar":"نسخ  Ctrl+C",          "ru":"Копировать  Ctrl+C",   "es":"Copiar  Ctrl+C",               "fr":"Copier  Ctrl+C"},
    "cut_short":          {"he":"גזור  Ctrl+X",          "en":"Cut    Ctrl+X",                 "ar":"قص    Ctrl+X",         "ru":"Вырезать  Ctrl+X",     "es":"Cortar  Ctrl+X",               "fr":"Couper  Ctrl+X"},

    # ── Statuses ─────────────────────────────────────────────────────────────
    "status_pending":     {"he":"ממתין...",               "en":"Pending...",                    "ar":"في الانتظار...",        "ru":"Ожидание...",           "es":"Pendiente...",                  "fr":"En attente..."},
    "status_queued":      {"he":"בתור",                  "en":"Queued",                        "ar":"في الطابور",            "ru":"В очереди",             "es":"En cola",                       "fr":"En file"},
    "status_downloading": {"he":"מוריד...",               "en":"Downloading...",                "ar":"جارٍ التنزيل...",       "ru":"Загрузка...",           "es":"Descargando...",                "fr":"Téléchargement..."},
    "status_paused":      {"he":"⏸ מושהה",               "en":"⏸ Paused",                     "ar":"⏸ متوقف",              "ru":"⏸ Пауза",              "es":"⏸ Pausado",                    "fr":"⏸ Pausé"},
    "status_merging":     {"he":"ממזג חלקים...",          "en":"Merging parts...",              "ar":"دمج الأجزاء...",        "ru":"Объединение...",        "es":"Uniendo partes...",             "fr":"Fusion des parties..."},
    "status_hashing":     {"he":"בודק hash...",           "en":"Verifying hash...",             "ar":"التحقق من الهاش...",   "ru":"Проверка хэша...",      "es":"Verificando hash...",           "fr":"Vérification hash..."},
    "status_extracting":  {"he":"מחלץ...",                "en":"Extracting...",                 "ar":"جارٍ الاستخراج...",    "ru":"Распаковка...",         "es":"Extrayendo...",                 "fr":"Extraction..."},
    "status_seeding":     {"he":"מזרע...",                "en":"Seeding...",                    "ar":"جارٍ البذر...",         "ru":"Раздача...",            "es":"Sembrando...",                  "fr":"Partage..."},
    "status_scanning":    {"he":"סורק וירוסים...",        "en":"Scanning for viruses...",       "ar":"فحص الفيروسات...",     "ru":"Сканирование...",       "es":"Escaneando virus...",           "fr":"Analyse antivirus..."},
    "status_done":        {"he":"הושלם ✓",               "en":"Done ✓",                        "ar":"اكتمل ✓",              "ru":"Готово ✓",              "es":"Completado ✓",                  "fr":"Terminé ✓"},
    "status_failed":      {"he":"נכשל ✗",                "en":"Failed ✗",                      "ar":"فشل ✗",                "ru":"Ошибка ✗",              "es":"Error ✗",                       "fr":"Échec ✗"},
    "status_cancelled":   {"he":"בוטל",                  "en":"Cancelled",                     "ar":"ملغى",                 "ru":"Отменено",              "es":"Cancelado",                     "fr":"Annulé"},
    "status_starting":    {"he":"⏰ מתחיל...",            "en":"⏰ Starting...",                 "ar":"⏰ جارٍ البدء...",     "ru":"⏰ Запуск...",          "es":"⏰ Iniciando...",                "fr":"⏰ Démarrage..."},
    "status_merging2":    {"he":"ממזג חלקים...",          "en":"Merging parts...",              "ar":"دمج الأجزاء...",        "ru":"Объединение...",        "es":"Uniendo partes...",             "fr":"Fusion..."},
    "ready":              {"he":"מוכן להורדה",            "en":"Ready to download",             "ar":"جاهز للتنزيل",         "ru":"Готово к загрузке",    "es":"Listo para descargar",          "fr":"Prêt à télécharger"},
    "verified_copied":    {"he":"URL הועתק ✓",            "en":"URL copied ✓",                  "ar":"تم نسخ الرابط ✓",     "ru":"URL скопирован ✓",     "es":"URL copiado ✓",                 "fr":"URL copié ✓"},
    "pasted_ok":          {"he":"✓ הודבק",               "en":"✓ Pasted",                      "ar":"✓ تم اللصق",           "ru":"✓ Вставлено",          "es":"✓ Pegado",                      "fr":"✓ Collé"},
    "verifying":          {"he":"מאמת...",                "en":"Verifying...",                  "ar":"جارٍ التحقق...",        "ru":"Проверка...",           "es":"Verificando...",                "fr":"Vérification..."},

    # ── URL input area ───────────────────────────────────────────────────────
    "url_placeholder":    {"he":"הדבק קישורים כאן — שורה אחת לכל קישור (http/https/ftp)",
                           "en":"Paste links here — one per line (http/https/ftp)",
                           "ar":"الصق الروابط هنا — رابط في كل سطر",
                           "ru":"Вставьте ссылки — по одной на строку",
                           "es":"Pega los enlaces aquí — uno por línea",
                           "fr":"Collez les liens ici — un par ligne"},
    "drop_hint":          {"he":"גרור קבצים לכאן",       "en":"Drop files here",               "ar":"اسحب الملفات هنا",     "ru":"Перетащите файлы сюда","es":"Arrastra archivos aquí",         "fr":"Déposez les fichiers ici"},
    "ctrl_enter_hint":    {"he":"Ctrl+Enter = הוסף",     "en":"Ctrl+Enter = Add",              "ar":"Ctrl+Enter = إضافة",   "ru":"Ctrl+Enter = Добавить","es":"Ctrl+Enter = Agregar",           "fr":"Ctrl+Entrée = Ajouter"},
    "add_valid":          {"he":"הוסף קישורים תקינים",  "en":"Add valid links",               "ar":"إضافة الروابط الصحيحة","ru":"Добавить корректные",  "es":"Agregar enlaces válidos",        "fr":"Ajouter les liens valides"},
    "verify_links":       {"he":"אמת קישורים",           "en":"Verify links",                  "ar":"التحقق من الروابط",    "ru":"Проверить ссылки",     "es":"Verificar enlaces",             "fr":"Vérifier les liens"},
    "no_url_warning":     {"he":"הכנס URL תחילה",        "en":"Enter a URL first",             "ar":"أدخل رابطاً أولاً",    "ru":"Введите URL",          "es":"Introduce un URL primero",      "fr":"Saisissez une URL d'abord"},
    "url_missing_title":  {"he":"URL חסר",               "en":"Missing URL",                   "ar":"الرابط مفقود",         "ru":"URL отсутствует",      "es":"URL faltante",                  "fr":"URL manquant"},
    "verify_urls_btn":    {"he":"🔍 אמת URLs",           "en":"🔍 Verify URLs",                 "ar":"🔍 التحقق من الروابط", "ru":"🔍 Проверить URLs",    "es":"🔍 Verificar URLs",             "fr":"🔍 Vérifier les URLs"},

    # ── Tabs ─────────────────────────────────────────────────────────────────
    "tab_all":            {"he":"הכל",                   "en":"All",                           "ar":"الكل",                 "ru":"Все",                  "es":"Todo",                          "fr":"Tout"},
    "tab_active":         {"he":"פעיל",                  "en":"Active",                        "ar":"نشط",                  "ru":"Активные",             "es":"Activo",                        "fr":"Actif"},
    "tab_done":           {"he":"הושלם",                 "en":"Done",                          "ar":"مكتمل",                "ru":"Готово",               "es":"Completado",                    "fr":"Terminé"},
    "tab_failed":         {"he":"נכשל",                  "en":"Failed",                        "ar":"فاشل",                 "ru":"Ошибки",               "es":"Error",                         "fr":"Échec"},

    # ── Settings sections ─────────────────────────────────────────────────────
    "settings_title":     {"he":"⚙ הגדרות FetchPro",    "en":"⚙ FetchPro Settings",           "ar":"⚙ إعدادات FetchPro",  "ru":"⚙ Настройки FetchPro","es":"⚙ Ajustes FetchPro",            "fr":"⚙ Paramètres FetchPro"},
    "sec_general":        {"he":"⚙ כללי",               "en":"⚙ General",                     "ar":"⚙ عام",               "ru":"⚙ Общие",              "es":"⚙ General",                    "fr":"⚙ Général"},
    "sec_network":        {"he":"🌐 רשת",               "en":"🌐 Network",                     "ar":"🌐 الشبكة",             "ru":"🌐 Сеть",              "es":"🌐 Red",                        "fr":"🌐 Réseau"},
    "sec_ytdlp":          {"he":"🎬 YouTube / מדיה",    "en":"🎬 YouTube / Media",             "ar":"🎬 يوتيوب / الوسائط",  "ru":"🎬 YouTube / Медиа",   "es":"🎬 YouTube / Medios",           "fr":"🎬 YouTube / Médias"},
    "sec_torrent":        {"he":"🧲 טורנטים",           "en":"🧲 Torrents",                    "ar":"🧲 التورنت",            "ru":"🧲 Торренты",          "es":"🧲 Torrents",                   "fr":"🧲 Torrents"},
    "sec_virustotal":     {"he":"🛡 VirusTotal",         "en":"🛡 VirusTotal",                  "ar":"🛡 فيروس توتال",       "ru":"🛡 VirusTotal",         "es":"🛡 VirusTotal",                 "fr":"🛡 VirusTotal"},
    "sec_proxy":          {"he":"🌐 Proxy / SOCKS5",    "en":"🌐 Proxy / SOCKS5",             "ar":"🌐 الوكيل / SOCKS5",   "ru":"🌐 Прокси / SOCKS5",   "es":"🌐 Proxy / SOCKS5",            "fr":"🌐 Proxy / SOCKS5"},
    "sec_automation":     {"he":"⚙ אוטומציה",           "en":"⚙ Automation",                  "ar":"⚙ الأتمتة",            "ru":"⚙ Автоматизация",      "es":"⚙ Automatización",              "fr":"⚙ Automatisation"},
    "sec_schedule":       {"he":"📅 תזמון טווח שעות",  "en":"📅 Time-range scheduling",       "ar":"📅 جدولة النطاق الزمني","ru":"📅 Расписание",        "es":"📅 Horario programado",          "fr":"📅 Plage horaire"},
    "sec_startup":        {"he":"🚀 הפעלה",              "en":"🚀 Startup",                     "ar":"🚀 بدء التشغيل",       "ru":"🚀 Автозапуск",        "es":"🚀 Inicio automático",           "fr":"🚀 Démarrage"},
    "sec_language":       {"he":"🌍 שפה",               "en":"🌍 Language",                    "ar":"🌍 اللغة",              "ru":"🌍 Язык",              "es":"🌍 Idioma",                     "fr":"🌍 Langue"},

    # ── Settings labels ───────────────────────────────────────────────────────
    "startup_win":        {"he":"הפעל FetchPro עם Windows","en":"Launch FetchPro with Windows","ar":"تشغيل FetchPro مع ويندوز","ru":"Запускать с Windows","es":"Iniciar con Windows",            "fr":"Lancer avec Windows"},
    "startup_win_na":     {"he":"הפעלה אוטומטית זמינה רק ב-Windows",
                           "en":"Auto-start available on Windows only",
                           "ar":"البدء التلقائي متاح على ويندوز فقط",
                           "ru":"Автозапуск доступен только в Windows",
                           "es":"Inicio automático solo en Windows",
                           "fr":"Démarrage automatique Windows uniquement"},
    "startup_enabled":    {"he":"✓ מופעל",              "en":"✓ Enabled",                     "ar":"✓ مفعّل",              "ru":"✓ Включено",           "es":"✓ Habilitado",                  "fr":"✓ Activé"},
    "startup_disabled":   {"he":"✗ מושבת",              "en":"✗ Disabled",                    "ar":"✗ معطّل",              "ru":"✗ Отключено",          "es":"✗ Deshabilitado",               "fr":"✗ Désactivé"},
    "dark_theme":         {"he":"נושא כהה",             "en":"Dark theme",                    "ar":"السمة الداكنة",        "ru":"Тёмная тема",          "es":"Tema oscuro",                   "fr":"Thème sombre"},
    "notify_done":        {"he":"התראה בסיום הורדה",   "en":"Notify on download done",       "ar":"تنبيه عند اكتمال التنزيل","ru":"Уведомление по окончании","es":"Notificar al terminar",     "fr":"Notifier à la fin"},
    "save_dir_label":     {"he":"📁 תיקיית שמירה:",     "en":"📁 Save folder:",               "ar":"📁 مجلد الحفظ:",       "ru":"📁 Папка сохранения:", "es":"📁 Carpeta de guardado:",        "fr":"📁 Dossier de sauvegarde:"},
    "max_concurrent":     {"he":"הורדות מקביליות:",     "en":"Max parallel downloads:",       "ar":"التنزيلات المتوازية:", "ru":"Макс. параллельных:",  "es":"Máx. paralelas:",               "fr":"Max. parallèles:"},
    "bw_limit":           {"he":"הגבלת מהירות (KB/s, 0=ללא):","en":"Bandwidth limit (KB/s, 0=none):","ar":"حد عرض النطاق:","ru":"Лимит скорости (КБ/с):","es":"Límite de ancho de banda:","fr":"Limite de bande passante:"},
    "language_label":     {"he":"שפת ממשק:",             "en":"Interface language:",           "ar":"لغة الواجهة:",         "ru":"Язык интерфейса:",     "es":"Idioma de interfaz:",           "fr":"Langue de l'interface:"},
    "zero_no_limit":      {"he":"(0 = ללא הגבלה)",      "en":"(0 = no limit)",                "ar":"(0 = بلا حد)",         "ru":"(0 = без ограничений)","es":"(0 = sin límite)",              "fr":"(0 = sans limite)"},
    "proxy_type":         {"he":"סוג Proxy:",            "en":"Proxy type:",                   "ar":"نوع الوكيل:",          "ru":"Тип прокси:",          "es":"Tipo de proxy:",                "fr":"Type de proxy:"},
    "select_folder":      {"he":"בחר תיקייה",           "en":"Choose folder",                 "ar":"اختر مجلداً",          "ru":"Выбрать папку",        "es":"Elegir carpeta",                "fr":"Choisir un dossier"},
    "select_torrent":     {"he":"פתח קובץ .torrent",    "en":"Open .torrent file",            "ar":"افتح ملف .torrent",   "ru":"Открыть .torrent",     "es":"Abrir archivo .torrent",        "fr":"Ouvrir fichier .torrent"},
    "sched_from":         {"he":"מ:",                   "en":"From:",                         "ar":"من:",                  "ru":"С:",                   "es":"Desde:",                        "fr":"De:"},
    "sched_to":           {"he":"עד:",                  "en":"To:",                           "ar":"حتى:",                 "ru":"До:",                  "es":"Hasta:",                        "fr":"Jusqu'à:"},
    "schedule_btn":       {"he":"תזמן",                 "en":"Schedule",                      "ar":"جدولة",                "ru":"Расписание",           "es":"Programar",                     "fr":"Planifier"},
    "shutdown_action":    {"he":"פעולת כיבוי:",         "en":"Shutdown action:",              "ar":"إجراء الإيقاف:",       "ru":"Действие при завершении:","es":"Acción de apagado:",         "fr":"Action à l'arrêt:"},
    "max_dl_label":       {"he":"מקסימום הורדות:",      "en":"Max downloads:",                "ar":"الحد الأقصى للتنزيل:", "ru":"Макс. загрузок:",      "es":"Máx. descargas:",               "fr":"Max. téléchargements:"},
    "adv_settings":       {"he":"⚙ הגדרות הורדה",      "en":"⚙ Download settings",           "ar":"⚙ إعدادات التنزيل",   "ru":"⚙ Настройки загрузки", "es":"⚙ Configuración de descarga",  "fr":"⚙ Paramètres de téléchargement"},
    "show_key":           {"he":"הצג",                  "en":"Show",                          "ar":"إظهار",                "ru":"Показать",             "es":"Mostrar",                       "fr":"Afficher"},
    "best_quality":       {"he":"הכי טוב",              "en":"Best",                          "ar":"الأفضل",               "ru":"Лучшее",               "es":"Mejor",                         "fr":"Meilleur"},
    "vt_free_key":        {"he":"API key חינמי: virustotal.com/gui/join-us (עד 4/דקה)",
                           "en":"Free API key: virustotal.com/gui/join-us (4 req/min)",
                           "ar":"مفتاح API مجاني: virustotal.com (4 طلبات/دقيقة)",
                           "ru":"Бесплатный ключ: virustotal.com (4 запр./мин)",
                           "es":"API key gratis: virustotal.com (4 req./min)",
                           "fr":"Clé API gratuite: virustotal.com (4 req./min)"},
    "vt_test_btn":        {"he":"🔍 בדוק API Key",      "en":"🔍 Test API Key",               "ar":"🔍 اختبار مفتاح API", "ru":"🔍 Проверить ключ",    "es":"🔍 Probar API Key",             "fr":"🔍 Tester la clé API"},

    # ── Media dialog ──────────────────────────────────────────────────────────
    "media_dialog_title": {"he":"🎬 הורדת מדיה — YouTube / מוזיקה",
                           "en":"🎬 Media Download — YouTube / Music",
                           "ar":"🎬 تنزيل الوسائط — يوتيوب / موسيقى",
                           "ru":"🎬 Загрузка медиа — YouTube / Музыка",
                           "es":"🎬 Descarga de medios — YouTube / Música",
                           "fr":"🎬 Téléchargement — YouTube / Musique"},
    "media_dialog_hdr":   {"he":"🎬 הורדת מדיה — YouTube / מוזיקה",
                           "en":"🎬 Media Download — YouTube / Music",
                           "ar":"🎬 تنزيل الوسائط",
                           "ru":"🎬 Загрузка медиа",
                           "es":"🎬 Descarga de medios",
                           "fr":"🎬 Téléchargement médias"},
    "media_url":          {"he":"🔗 כתובת URL:",         "en":"🔗 URL:",                       "ar":"🔗 الرابط:",            "ru":"🔗 URL:",              "es":"🔗 URL:",                       "fr":"🔗 URL:"},
    "media_type":         {"he":"📥 סוג הורדה:",         "en":"📥 Download type:",             "ar":"📥 نوع التنزيل:",      "ru":"📥 Тип загрузки:",     "es":"📥 Tipo de descarga:",           "fr":"📥 Type de téléchargement:"},
    "media_video":        {"he":"🎬 וידאו",              "en":"🎬 Video",                      "ar":"🎬 فيديو",              "ru":"🎬 Видео",             "es":"🎬 Video",                      "fr":"🎬 Vidéo"},
    "media_audio_only":   {"he":"🎵 אודיו בלבד (MP3)",  "en":"🎵 Audio only (MP3)",           "ar":"🎵 صوت فقط (MP3)",     "ru":"🎵 Только аудио (MP3)","es":"🎵 Solo audio (MP3)",           "fr":"🎵 Audio seulement (MP3)"},
    "media_quality":      {"he":"📺 איכות וידאו:",       "en":"📺 Video quality:",             "ar":"📺 جودة الفيديو:",     "ru":"📺 Качество видео:",   "es":"📺 Calidad de vídeo:",           "fr":"📺 Qualité vidéo:"},
    "media_audio_format": {"he":"🎵 פורמט אודיו:",       "en":"🎵 Audio format:",              "ar":"🎵 صيغة الصوت:",       "ru":"🎵 Формат аудио:",     "es":"🎵 Formato de audio:",           "fr":"🎵 Format audio:"},
    "media_options":      {"he":"📋 אפשרויות:",          "en":"📋 Options:",                   "ar":"📋 الخيارات:",          "ru":"📋 Параметры:",        "es":"📋 Opciones:",                  "fr":"📋 Options:"},
    "media_playlist":     {"he":"הורד פלייליסט שלם",    "en":"Download entire playlist",      "ar":"تنزيل قائمة التشغيل","ru":"Скачать весь плейлист","es":"Descargar lista completa",       "fr":"Télécharger toute la liste"},
    "media_thumbnail":    {"he":"הטמע תמונה ממוזערת",   "en":"Embed thumbnail",               "ar":"تضمين الصورة المصغرة","ru":"Встроить обложку",     "es":"Incluir miniatura",             "fr":"Intégrer la vignette"},
    "media_metadata":     {"he":"הוסף מטא-דאטה",        "en":"Add metadata",                  "ar":"إضافة البيانات الوصفية","ru":"Добавить метаданные","es":"Agregar metadatos",             "fr":"Ajouter les métadonnées"},
    "media_tags":         {"he":"🏷 תגיות (מופרדות בפסיק):",
                           "en":"🏷 Tags (comma-separated):",
                           "ar":"🏷 العلامات:",
                           "ru":"🏷 Теги (через запятую):",
                           "es":"🏷 Etiquetas (separadas por coma):",
                           "fr":"🏷 Tags (séparés par virgule):"},
    "add_to_download":    {"he":"⬇ הוסף להורדה",        "en":"⬇ Add to Download",             "ar":"⬇ إضافة للتنزيل",     "ru":"⬇ Добавить",          "es":"⬇ Agregar descarga",            "fr":"⬇ Ajouter au téléchargement"},
    "embed_thumb_audio":  {"he":"הטמע תמונה בקובץ אודיו","en":"Embed thumbnail in audio file","ar":"تضمين صورة في ملف الصوت","ru":"Обложка в аудиофайле","es":"Miniatura en archivo de audio","fr":"Vignette dans le fichier audio"},
    "add_metadata_audio": {"he":"הוסף מטא-דאטה (כותרת, אמן)","en":"Add metadata (title, artist)","ar":"إضافة بيانات وصفية","ru":"Метаданные (название, автор)","es":"Añadir metadatos","fr":"Ajouter métadonnées"},
    "media_default_fmt":  {"he":"פורמט ברירת מחדל:",    "en":"Default format:",               "ar":"الصيغة الافتراضية:",   "ru":"Формат по умолчанию:","es":"Formato predeterminado:",        "fr":"Format par défaut:"},
    "media_sites_list":   {"he":"YouTube · Vimeo · TikTok · SoundCloud · Instagram · Twitter · Reddit · ועוד",
                           "en":"YouTube · Vimeo · TikTok · SoundCloud · Instagram · Twitter · Reddit · and more",
                           "ar":"يوتيوب · فيميو · تيك توك · ساوندكلاود · إنستغرام · تويتر · ريديت · والمزيد",
                           "ru":"YouTube · Vimeo · TikTok · SoundCloud · Instagram · Twitter · Reddit · и другие",
                           "es":"YouTube · Vimeo · TikTok · SoundCloud · Instagram · Twitter · Reddit · y más",
                           "fr":"YouTube · Vimeo · TikTok · SoundCloud · Instagram · Twitter · Reddit · et plus"},
    "no_ytdlp":           {"he":"⚠ yt-dlp לא מותקן — הרץ: pip install yt-dlp",
                           "en":"⚠ yt-dlp not installed — run: pip install yt-dlp",
                           "ar":"⚠ yt-dlp غير مثبت — شغّل: pip install yt-dlp",
                           "ru":"⚠ yt-dlp не установлен — выполните: pip install yt-dlp",
                           "es":"⚠ yt-dlp no instalado — ejecuta: pip install yt-dlp",
                           "fr":"⚠ yt-dlp non installé — exécutez: pip install yt-dlp"},
    "no_ytdlp_short":     {"he":"(yt-dlp לא מותקן — YouTube מושבת)",
                           "en":"(yt-dlp not installed — YouTube disabled)",
                           "ar":"(yt-dlp غير مثبت — يوتيوب معطّل)",
                           "ru":"(yt-dlp не установлен — YouTube отключён)",
                           "es":"(yt-dlp no instalado — YouTube desactivado)",
                           "fr":"(yt-dlp non installé — YouTube désactivé)"},
    "no_libtorrent":      {"he":"(libtorrent לא מותקן — טורנטים מושבתים)",
                           "en":"(libtorrent not installed — torrents disabled)",
                           "ar":"(libtorrent غير مثبت — التورنت معطّل)",
                           "ru":"(libtorrent не установлен — торренты отключены)",
                           "es":"(libtorrent no instalado — torrents desactivados)",
                           "fr":"(libtorrent non installé — torrents désactivés)"},

    # ── Stats dialog ───────────────────────────────────────────────────────────
    "stats_title":        {"he":"📊 סטטיסטיקות שימוש",  "en":"📊 Usage Statistics",           "ar":"📊 إحصائيات الاستخدام","ru":"📊 Статистика",        "es":"📊 Estadísticas",               "fr":"📊 Statistiques d'utilisation"},
    "stats_total_files":  {"he":"📦 סה\"כ הורדות",      "en":"📦 Total downloads",            "ar":"📦 إجمالي التنزيلات", "ru":"📦 Всего загрузок",    "es":"📦 Total descargas",            "fr":"📦 Total téléchargements"},
    "stats_total_bytes":  {"he":"💾 סה\"כ נפח",         "en":"💾 Total size",                 "ar":"💾 الحجم الإجمالي",   "ru":"💾 Всего размер",      "es":"💾 Tamaño total",               "fr":"💾 Taille totale"},
    "stats_fastest":      {"he":"⚡ שיא מהירות",          "en":"⚡ Peak speed",                  "ar":"⚡ أقصى سرعة",         "ru":"⚡ Пиковая скорость",   "es":"⚡ Velocidad máxima",           "fr":"⚡ Vitesse maximale"},
    "stats_sessions":     {"he":"🔄 סשנים",              "en":"🔄 Sessions",                   "ar":"🔄 الجلسات",           "ru":"🔄 Сеансов",           "es":"🔄 Sesiones",                   "fr":"🔄 Sessions"},
    "stats_session_files":{"he":"📥 הסשן הנוכחי",        "en":"📥 Current session",            "ar":"📥 الجلسة الحالية",    "ru":"📥 Текущий сеанс",     "es":"📥 Sesión actual",              "fr":"📥 Session actuelle"},
    "stats_session_bytes":{"he":"📊 נפח הסשן",           "en":"📊 Session size",               "ar":"📊 حجم الجلسة",        "ru":"📊 Размер сеанса",     "es":"📊 Tamaño de sesión",           "fr":"📊 Taille de session"},
    "stats_active_now":   {"he":"🟢 פעיל עכשיו",         "en":"🟢 Active now",                 "ar":"🟢 نشط الآن",          "ru":"🟢 Сейчас активно",    "es":"🟢 Activo ahora",               "fr":"🟢 Actif maintenant"},
    "active_now":         {"he":"פעיל עכשיו",            "en":"Active now",                    "ar":"نشط الآن",             "ru":"Сейчас активно",       "es":"Activo ahora",                  "fr":"Actif maintenant"},

    # ── History panel ─────────────────────────────────────────────────────────
    "history_title":      {"he":"📜 היסטוריית הורדות",  "en":"📜 Download History",           "ar":"📜 سجل التنزيلات",    "ru":"📜 История загрузок",  "es":"📜 Historial de descargas",     "fr":"📜 Historique des téléchargements"},
    "csv_export":         {"he":"ייצוא CSV",             "en":"Export CSV",                    "ar":"تصدير CSV",            "ru":"Экспорт CSV",          "es":"Exportar CSV",                  "fr":"Exporter CSV"},
    "csv_export_btn":     {"he":"📋 ייצוא CSV",          "en":"📋 Export CSV",                 "ar":"📋 تصدير CSV",         "ru":"📋 Экспорт CSV",       "es":"📋 Exportar CSV",               "fr":"📋 Exporter CSV"},
    "export_history_csv": {"he":"ייצא היסטוריה ל-CSV",  "en":"Export history to CSV",         "ar":"تصدير السجل إلى CSV", "ru":"Экспорт истории CSV",  "es":"Exportar historial CSV",        "fr":"Exporter historique CSV"},
    "export_links":       {"he":"ייצא קישורים",          "en":"Export links",                  "ar":"تصدير الروابط",        "ru":"Экспорт ссылок",       "es":"Exportar enlaces",              "fr":"Exporter les liens"},
    "import_links":       {"he":"ייבוא קישורים",         "en":"Import links",                  "ar":"استيراد الروابط",      "ru":"Импорт ссылок",        "es":"Importar enlaces",              "fr":"Importer les liens"},
    "about":              {"he":"אודות FetchPro",        "en":"About FetchPro",                "ar":"حول FetchPro",         "ru":"О программе FetchPro", "es":"Acerca de FetchPro",            "fr":"À propos de FetchPro"},

    # ── Notifications ────────────────────────────────────────────────────────
    "queue_restored":     {"he":"✓ שוחזרו {n} הורדות מהסשן הקודם",
                           "en":"✓ Restored {n} downloads from previous session",
                           "ar":"✓ تم استعادة {n} تنزيل من الجلسة السابقة",
                           "ru":"✓ Восстановлено {n} загрузок из предыдущего сеанса",
                           "es":"✓ Se restauraron {n} descargas de la sesión anterior",
                           "fr":"✓ {n} téléchargements restaurés depuis la session précédente"},
    "copied":             {"he":"URL הועתק ✓",            "en":"URL copied ✓",                  "ar":"تم نسخ الرابط ✓",     "ru":"URL скопирован ✓",     "es":"URL copiado ✓",                "fr":"URL copié ✓"},
    "no_url":             {"he":"אין קישור להורדה",      "en":"No URL to download",            "ar":"لا يوجد رابط للتنزيل","ru":"Нет URL для загрузки", "es":"No hay URL para descargar",     "fr":"Pas d'URL à télécharger"},
    "add_download_adv":   {"he":"הוסף הורדה",            "en":"Add download",                  "ar":"إضافة تنزيل",          "ru":"Добавить загрузку",    "es":"Agregar descarga",              "fr":"Ajouter un téléchargement"},
    "lang_restart_warn":  {"he":"⚠ שינוי שפה ייכנס לתוקף בהפעלה הבאה",
                           "en":"⚠ Language change takes effect on next restart",
                           "ar":"⚠ تغيير اللغة يتطلب إعادة التشغيل",
                           "ru":"⚠ Смена языка вступит в силу при перезапуске",
                           "es":"⚠ El cambio de idioma se aplica al reiniciar",
                           "fr":"⚠ Le changement de langue prend effet au redémarrage"},
    "choose_lang":        {"he":"🌍  בחר שפת ממשק",       "en":"🌍  Choose interface language",  "ar":"🌍  اختر لغة الواجهة", "ru":"🌍  Выберите язык интерфейса","es":"🌍  Elige el idioma de interfaz","fr":"🌍  Choisissez la langue d'interface"},

    # ── Keyboard hint ─────────────────────────────────────────────────────────
    "shortcuts":          {"he":"Space=השהה  Del=בטל  R=נסה שנית  Ctrl+V=הדבק",
                           "en":"Space=Pause  Del=Cancel  R=Retry  Ctrl+V=Paste",
                           "ar":"مسافة=إيقاف  Del=إلغاء  R=إعادة  Ctrl+V=لصق",
                           "ru":"Пробел=Пауза  Del=Отмена  R=Повтор  Ctrl+V=Вставить",
                           "es":"Espacio=Pausar  Del=Cancelar  R=Reintentar  Ctrl+V=Pegar",
                           "fr":"Espace=Pause  Suppr=Annuler  R=Réessayer  Ctrl+V=Coller"},
}


_LANG = "he"   # default — changed by user in Settings

def _t(key: str, **kwargs) -> str:
    """Translate key to current language, with optional format substitutions."""
    lang_dict = _STRINGS.get(key, {})
    text = lang_dict.get(_LANG) or lang_dict.get("he") or key
    return text.format(**kwargs) if kwargs else text

SUPPORTED_LANGS: dict[str, str] = {
    "he": "עברית",
    "en": "English",
    "ar": "العربية",
    "ru": "Русский",
    "es": "Español",
    "fr": "Français",
}

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
APP_NAME         = "FetchPro"
APP_DIR          = Path.home() / ".fetchpro"
DB_PATH          = APP_DIR / "history.db"
STATE_DIR        = APP_DIR / "resume"
SETTINGS_FILE    = APP_DIR / "settings.json"
QUEUE_FILE       = APP_DIR / "queue.json"
STATS_FILE       = APP_DIR / "stats.json"
REST_API_PORT    = 9100

MAX_DL_WORKERS    = 16
MAX_VFY_WORKERS   = 6
CHUNK_SIZE        = 524_288     # 512 KB per chunk — 8× faster for large files
MULTIPART_PARTS   = 16          # parallel segments per file
MULTIPART_MIN     = 1_000_000   # use multipart if file >= 1 MB
CONNECT_TIMEOUT   = 8
READ_TIMEOUT      = 20
POLL_MS           = 60          # UI refresh — ~16 fps
PROGRESS_THROTTLE = 0.04        # 25 UI updates/sec
SPEED_WINDOW_SEC  = 1.5         # rolling speed window (shorter = more reactive)
BRIDGE_PORT       = 9099        # Chrome extension bridge

# HTTP session pool — one session per download thread for TCP keep-alive
_SESSION_POOL: dict[int, "requests.Session"] = {}
_SESSION_LOCK  = threading.Lock()
_PROXY_SETTINGS: dict = {}   # updated by app when settings change


def _build_proxy_dict(settings: "Settings") -> dict:
    """Build requests proxies dict from settings."""
    if not settings.proxy_enabled or not settings.proxy_host:
        return {}
    auth = f"{settings.proxy_user}:{settings.proxy_pass}@" if settings.proxy_user else ""
    proxy_url = f"{settings.proxy_type}://{auth}{settings.proxy_host}:{settings.proxy_port}"
    return {"http": proxy_url, "https": proxy_url}


def _get_session(proxies: dict | None = None) -> "requests.Session":
    """Return a per-thread requests.Session with optimised pool settings."""
    tid = threading.get_ident()
    with _SESSION_LOCK:
        if tid not in _SESSION_POOL:
            s = requests.Session()
            adapter = requests.adapters.HTTPAdapter(
                pool_connections=8,
                pool_maxsize=32,
                max_retries=requests.adapters.Retry(
                    total=3, backoff_factor=0.3,
                    status_forcelist={500, 502, 503, 504},
                    allowed_methods={"GET", "HEAD"},
                ),
            )
            s.mount("http://",  adapter)
            s.mount("https://", adapter)
            s.headers.update({
                "User-Agent":      "FetchPro/5.5",
                "Accept-Encoding": "identity",
                "Connection":      "keep-alive",
            })
            _SESSION_POOL[tid] = s
        s = _SESSION_POOL[tid]
        if proxies is not None:
            s.proxies.update(proxies)
        elif _PROXY_SETTINGS:
            s.proxies.update(_PROXY_SETTINGS)
        return s

# File-type → auto-category subdirectory
CATEGORY_MAP: dict[str, str] = {
    **{e: "Videos"    for e in ("mp4","mkv","avi","mov","wmv","flv","webm")},
    **{e: "Music"     for e in ("mp3","flac","wav","aac","ogg","opus","m4a")},
    **{e: "Images"    for e in ("png","jpg","jpeg","gif","webp","bmp","svg","tiff")},
    **{e: "Documents" for e in ("pdf","doc","docx","xls","xlsx","ppt","pptx","txt","csv","epub")},
    **{e: "Archives"  for e in ("zip","rar","7z","tar","gz","xz","bz2","zst")},
    **{e: "Programs"  for e in ("exe","msi","dmg","pkg","deb","rpm","apk")},
}

# File-type → emoji icon
FILE_ICONS: dict[str, str] = {
    **{e: "🎬" for e in ("mp4","mkv","avi","mov","wmv","flv","webm")},
    **{e: "🎵" for e in ("mp3","flac","wav","aac","ogg","opus","m4a")},
    **{e: "🖼" for e in ("png","jpg","jpeg","gif","webp","bmp","svg","tiff")},
    **{e: "📄" for e in ("pdf","doc","docx","txt","epub")},
    **{e: "📊" for e in ("xls","xlsx","csv")},
    **{e: "📦" for e in ("zip","rar","7z","tar","gz","xz","bz2","zst")},
    **{e: "⚙" for e in ("exe","msi","dmg","pkg","deb","rpm","apk")},
    **{e: "🖥" for e in ("iso","img","bin")},
    **{e: "🐍" for e in ("py","pyw")},
}

def _file_icon(filename: str) -> str:
    ext = Path(filename).suffix.lstrip(".").lower()
    return FILE_ICONS.get(ext, "📥")

# Priority levels
class Priority(Enum):
    HIGH   = "גבוהה"
    NORMAL = "רגילה"
    LOW    = "נמוכה"

PRIORITY_COLOR: dict[Priority, str] = {
    Priority.HIGH:   "#ff4444",
    Priority.NORMAL: "#00e5ff",
    Priority.LOW:    "#888888",
}

# Filename sanitizer — removes chars illegal in Windows/Linux filenames
_ILLEGAL_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

def _sanitize_filename(name: str) -> str:
    """Remove illegal filesystem characters, prevent path traversal, and trim whitespace."""
    # Strip any directory separators to prevent path traversal attacks
    name = os.path.basename(name)
    name = _ILLEGAL_CHARS.sub("_", name).strip(". ")
    # Remove any remaining path traversal patterns
    name = name.replace("..", "_")
    return name or "download"

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------
@dataclass
class Settings:
    save_dir:              str  = str(Path.home() / "Downloads")
    max_concurrent:        int  = 4
    multipart:             bool = True
    auto_extract:          bool = False
    auto_open:             bool = False
    auto_categorize:       bool = False
    verify_hash:           bool = False
    dark_theme:            bool = True
    notify_done:           bool = True
    system_tray:           bool = True
    default_hash_algo:     str  = "sha256"
    startup_with_windows:  bool = False
    # v5.1 additions
    auto_retry:            bool = True
    max_retries:           int  = 3
    global_bw_limit_kbps:  int  = 0
    window_geometry:       str  = "920x720"
    confirm_delete_file:   bool = True
    # v5.2 additions
    clipboard_monitor:     bool = False   # auto-detect URLs copied to clipboard
    sound_on_done:         bool = False   # play sound when download completes
    max_file_size_mb:      int  = 0       # 0 = no limit; cancel if exceeded
    warn_duplicates:       bool = True    # warn if same URL already queued
    # v5.5 additions (torrent)
    torrent_max_dl_kbps:   int  = 0       # 0 = unlimited
    torrent_max_ul_kbps:   int  = 50      # upload limit (KB/s); 0 = unlimited
    torrent_seeding:       bool = False   # continue seeding after download completes
    torrent_seed_ratio:    float = 1.0    # stop seeding at this share ratio
    # v5.5 additions
    proxy_enabled:         bool = False
    proxy_type:            str  = "http"  # "http" | "socks5"
    proxy_host:            str  = ""
    proxy_port:            int  = 8080
    proxy_user:            str  = ""
    proxy_pass:            str  = ""
    schedule_range_enabled: bool = False
    schedule_range_start:   str  = "02:00"  # HH:MM
    schedule_range_end:     str  = "06:00"
    shutdown_after_queue:  bool = False
    shutdown_action:       str  = "sleep"   # "sleep" | "hibernate" | "shutdown"
    persistent_queue:      bool = True      # save/restore queue between sessions
    ytdlp_default_format:  str  = "best"    # "best"|"bestvideo"|"bestaudio"|"1080p"|"720p"|"480p"|"360p"|"mp3"|"m4a"
    ytdlp_prefer_audio:    bool = False     # default to audio-only
    ytdlp_embed_thumbnail: bool = True      # embed thumbnail in audio files
    ytdlp_add_metadata:    bool = True      # write title/artist tags
    virustotal_api_key:    str  = ""        # optional VT key for hash scanning
    virustotal_enabled:    bool = True      # enable/disable VT scanning
    disk_check_enabled:    bool = True      # warn if not enough disk space
    watchdog_timeout_min:  int  = 5         # restart stuck downloads after N minutes
    stats_enabled:         bool = True
    language:              str  = "he"   # "he"|"en"|"ar"|"ru"|"es"|"fr"

    @classmethod
    def load(cls) -> "Settings":
        try:
            data = json.loads(SETTINGS_FILE.read_text())
            valid = {f for f in cls.__dataclass_fields__}
            return cls(**{k: v for k, v in data.items() if k in valid})
        except (OSError, json.JSONDecodeError, TypeError, ValueError):
            return cls()

    def save(self) -> None:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        SETTINGS_FILE.write_text(json.dumps(asdict(self), indent=2))


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------
class DownloadStatus(Enum):
    PENDING     = "ממתין"
    QUEUED      = "בתור"
    DOWNLOADING = "מוריד..."
    MERGING     = "ממזג..."
    HASHING     = "מאמת..."
    EXTRACTING  = "מחלץ..."
    PAUSED      = "מושהה"
    SEEDING     = "מזרע..."     # torrent: upload-back phase
    SCANNING    = "סורק וירוסים..."  # VirusTotal scan
    DONE        = "הושלם ✓"
    FAILED      = "שגיאה ✗"
    CANCELLED   = "בוטל"

_TERMINAL = {DownloadStatus.DONE, DownloadStatus.FAILED, DownloadStatus.CANCELLED}
_ACTIVE   = {DownloadStatus.DOWNLOADING, DownloadStatus.MERGING,
             DownloadStatus.HASHING, DownloadStatus.EXTRACTING,
             DownloadStatus.SEEDING, DownloadStatus.SCANNING}


@dataclass
class DownloadItem:
    url:          str
    save_dir:     Path
    filename:     str  = ""
    status:       DownloadStatus = DownloadStatus.PENDING
    priority:     int  = 0          # higher = more urgent
    progress:     float = 0.0       # 0–100
    downloaded_bytes: int = 0
    total_bytes:  int = 0
    speed_bps:    float = 0.0
    eta_seconds:  float = -1.0
    error_msg:    str = ""
    hash_algo:    str = ""          # "md5" | "sha256" | ""
    expected_hash: str = ""         # user-supplied expected checksum
    actual_hash:  str = ""          # computed after download
    scheduled_at: datetime | None = None   # None = start now
    throttle_bps: int = 0           # 0 = unlimited
    multipart:    bool = True
    auto_extract: bool = False
    auto_open:    bool = False
    note:         str  = ""
    retry_count:  int  = 0
    added_at:     str  = ""
    dl_priority:  str  = "NORMAL"   # "HIGH" | "NORMAL" | "LOW"
    # Torrent-specific (unused for non-torrent downloads)
    torrent_seeds:   int = 0
    torrent_peers:   int = 0
    torrent_pieces:  str = ""        # e.g. "120/240"
    torrent_is_magnet: bool = False
    # Media/yt-dlp specific
    media_format:    str = ""        # "" = auto | "mp3"|"m4a"|"best"|"1080p" etc.
    media_is_audio:  bool = False    # True = audio-only extraction
    media_playlist:  bool = False    # True = download entire playlist
    media_title:     str = ""        # video title from yt-dlp info
    # Tags
    tags:            str = ""        # comma-separated tags
    # Watchdog
    _last_progress_bytes: int = field(default=0, repr=False, compare=False)
    _last_progress_time:  float = field(default_factory=time.monotonic, repr=False, compare=False)
    _cancel_event:    threading.Event = field(default_factory=threading.Event, repr=False, compare=False)
    _pause_event:     threading.Event = field(default_factory=threading.Event, repr=False, compare=False)
    _last_ui_update:  float           = field(default_factory=float, repr=False, compare=False)

    def __post_init__(self) -> None:
        if not self.filename:
            self.filename = _derive_filename(self.url)
        if not self.added_at:
            self.added_at = datetime.now().isoformat(timespec="seconds")

    @property
    def destination(self) -> Path:
        return self.save_dir / self.filename

    def cancel(self) -> None:
        self._cancel_event.set()
        self._pause_event.clear()

    def pause(self) -> None:
        self._pause_event.set()
        self.status = DownloadStatus.PAUSED

    def resume(self) -> None:
        self._pause_event.clear()
        self._cancel_event.clear()
        self.status = DownloadStatus.DOWNLOADING

    def reset_for_retry(self) -> None:
        self._cancel_event.clear()
        self._pause_event.clear()
        self.progress = 0.0
        self.downloaded_bytes = 0
        self.total_bytes = 0
        self.speed_bps = 0.0
        self.eta_seconds = -1.0
        self.error_msg = ""
        self.actual_hash = ""
        self._last_ui_update = 0.0
        self.status = DownloadStatus.PENDING


# ---------------------------------------------------------------------------
# History DB
# ---------------------------------------------------------------------------
class HistoryDB:
    """Simple SQLite3 history store. Thread-safe via a reentrant lock."""

    def __init__(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.RLock()
        self._con  = sqlite3.connect(str(path), check_same_thread=False)
        self._migrate()

    def _migrate(self) -> None:
        with self._lock:
            self._con.execute("""
                CREATE TABLE IF NOT EXISTS history (
                    id       INTEGER PRIMARY KEY AUTOINCREMENT,
                    url      TEXT NOT NULL,
                    filename TEXT NOT NULL,
                    save_dir TEXT NOT NULL,
                    status   TEXT NOT NULL,
                    size     INTEGER DEFAULT 0,
                    hash     TEXT DEFAULT '',
                    finished TEXT NOT NULL
                )
            """)
            self._con.commit()

    def record(self, item: DownloadItem) -> None:
        with self._lock:
            self._con.execute(
                "INSERT INTO history(url,filename,save_dir,status,size,hash,finished) "
                "VALUES (?,?,?,?,?,?,?)",
                (item.url, item.filename, str(item.save_dir),
                 item.status.name, item.downloaded_bytes,
                 item.actual_hash, datetime.now().isoformat(timespec="seconds")),
            )
            self._con.commit()

    def fetch(self, limit: int = 200) -> list[dict]:
        with self._lock:
            rows = self._con.execute(
                "SELECT id,url,filename,save_dir,status,size,hash,finished "
                "FROM history ORDER BY id DESC LIMIT ?", (limit,)
            ).fetchall()
        keys = ("id","url","filename","save_dir","status","size","hash","finished")
        return [dict(zip(keys, r)) for r in rows]

    def search(self, query: str, limit: int = 200) -> list[dict]:
        with self._lock:
            pattern = f"%{query}%"
            rows = self._con.execute(
                "SELECT id,url,filename,save_dir,status,size,hash,finished "
                "FROM history WHERE filename LIKE ? OR url LIKE ? "
                "ORDER BY id DESC LIMIT ?", (pattern, pattern, limit)
            ).fetchall()
        keys = ("id","url","filename","save_dir","status","size","hash","finished")
        return [dict(zip(keys, r)) for r in rows]

    def delete_by_id(self, row_id: int) -> None:
        with self._lock:
            self._con.execute("DELETE FROM history WHERE id=?", (row_id,))
            self._con.commit()

    def clear(self) -> None:
        with self._lock:
            self._con.execute("DELETE FROM history")
            self._con.commit()

    def close(self) -> None:
        with self._lock:
            self._con.close()


# ---------------------------------------------------------------------------
# Filename helpers
# ---------------------------------------------------------------------------
_claimed_names: dict[str, set[str]] = {}
_claimed_lock  = threading.Lock()


def _derive_filename(url: str) -> str:
    try:
        url = url.strip()
        # Magnet link — extract display name (dn= param) or fallback
        if url.startswith("magnet:"):
            m = re.search(r"[?&]dn=([^&]+)", url)
            if m:
                import urllib.parse as _up
                return _sanitize_filename(_up.unquote_plus(m.group(1)))
            # Use info-hash as name
            m2 = re.search(r"xt=urn:btih:([a-fA-F0-9]{40}|[A-Z2-7]{32})", url)
            return f"torrent-{m2.group(1)[:12]}" if m2 else "torrent"
        # Local .torrent file path
        if url.lower().endswith(".torrent") and not url.startswith("http"):
            return Path(url).stem   # will be updated to torrent name after metadata
        parsed = urlparse(url)
        # For FTP and HTTP
        path = unquote(parsed.path.rstrip("/").split("/")[-1]).split("?")[0]
        return path if path else "download"
    except Exception:
        return "download"


def _deduplicate_filename(base: str, save_dir: Path) -> str:
    stem, suffix = Path(base).stem, Path(base).suffix
    key = str(save_dir)
    n, result = 1, base
    with _claimed_lock:
        claimed = _claimed_names.setdefault(key, set())
        while (result in claimed
               or (save_dir / result).exists()
               or (save_dir / (result + ".part")).exists()):
            result = f"{stem} ({n}){suffix}"
            n += 1
        claimed.add(result)
    return result


def _release_name(filename: str, save_dir: Path) -> None:
    with _claimed_lock:
        s = _claimed_names.get(str(save_dir))
        if s:
            s.discard(filename)


# ---------------------------------------------------------------------------
# Format helpers
# ---------------------------------------------------------------------------
def _fmt_bytes(n: float) -> str:
    if n < 0:
        return "—"
    for u in ("B","KB","MB","GB"):
        if n < 1024:
            return f"{n:.1f} {u}"
        n /= 1024
    return f"{n:.1f} TB"


def _fmt_speed(bps: float) -> str:
    return f"{_fmt_bytes(bps)}/s"


def _fmt_eta(s: float) -> str:
    if s < 0:
        return ""
    s = int(s)
    if s < 60:   return f"{s}s"
    if s < 3600: return f"{s//60}m {s%60}s"
    return f"{s//3600}h {(s%3600)//60}m"


def _parse_cl(raw: str) -> int:
    try:
        v = int(raw.strip())
        return v if v >= 0 else 0
    except Exception:
        return 0


# ---------------------------------------------------------------------------
# Platform helpers
# ---------------------------------------------------------------------------
def _open_file(path: Path) -> None:
    """Open a file with the default application."""
    try:
        if platform.system() == "Windows":
            os.startfile(str(path))
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception as e:
        logger.warning("Could not open %s: %s", path, e)


def _send_notification(title: str, message: str) -> None:
    """Send a desktop notification via plyer, or Windows/macOS fallback."""
    if PLYER_OK:
        try:
            plyer_notif.notify(title=title, message=message, app_name=APP_NAME, timeout=5)
            return
        except Exception:
            pass
    # Windows toast via PowerShell
    if platform.system() == "Windows":
        try:
            script = (
                f'[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, '
                f'ContentType = WindowsRuntime] | Out-Null;'
                f'$xml = [Windows.UI.Notifications.ToastNotificationManager]'
                f'::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02);'
                f'$xml.GetElementsByTagName("text")[0].AppendChild($xml.CreateTextNode("{title}")) | Out-Null;'
                f'$xml.GetElementsByTagName("text")[1].AppendChild($xml.CreateTextNode("{message}")) | Out-Null;'
                f'$toast = [Windows.UI.Notifications.ToastNotification]::new($xml);'
                f'[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("{APP_NAME}").Show($toast)'
            )
            subprocess.Popen(["powershell","-WindowStyle","Hidden","-Command",script],
                             creationflags=0x08000000 if platform.system()=="Windows" else 0)
        except Exception:
            pass


def _compute_hash(path: Path, algo: str) -> str:
    h = hashlib.new(algo)
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _try_extract(path: Path, dest: Path) -> bool:
    """Try to auto-extract archive. Returns True on success."""
    try:
        if zipfile.is_zipfile(str(path)):
            with zipfile.ZipFile(str(path)) as z:
                z.extractall(str(dest))
            return True
        if tarfile.is_tarfile(str(path)):
            with tarfile.open(str(path)) as t:
                t.extractall(str(dest))
            return True
    except Exception as e:
        logger.warning("Extraction failed for %s: %s", path, e)
    return False


# ---------------------------------------------------------------------------
# Resume state
# ---------------------------------------------------------------------------
def _save_resume_state(item: DownloadItem) -> None:
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    key = hashlib.md5(item.url.encode()).hexdigest()[:16]
    data = {
        "url": item.url,
        "filename": item.filename,
        "save_dir": str(item.save_dir),
        "downloaded_bytes": item.downloaded_bytes,
        "total_bytes": item.total_bytes,
    }
    (STATE_DIR / f"{key}.json").write_text(json.dumps(data))


def _clear_resume_state(item: DownloadItem) -> None:
    key = hashlib.md5(item.url.encode()).hexdigest()[:16]
    f = STATE_DIR / f"{key}.json"
    try:
        f.unlink(missing_ok=True)
    except Exception:
        pass


def _load_resume_bytes(item: DownloadItem) -> int:
    """Return the number of bytes already downloaded, if a resume state exists."""
    key = hashlib.md5(item.url.encode()).hexdigest()[:16]
    f = STATE_DIR / f"{key}.json"
    try:
        data = json.loads(f.read_text())
        if (data.get("url") == item.url
                and (item.save_dir / (item.filename + ".part")).exists()):
            return int(data.get("downloaded_bytes", 0))
    except Exception:
        pass
    return 0


# ---------------------------------------------------------------------------
# Download engines
# ---------------------------------------------------------------------------

def _throttle_sleep(throttle_bps: int, chunk_len: int, t_start: float) -> None:
    """Sleep if needed to enforce a speed cap."""
    if throttle_bps <= 0:
        return
    expected = chunk_len / throttle_bps
    elapsed  = time.monotonic() - t_start
    gap = expected - elapsed
    if gap > 0:
        time.sleep(gap)


# ---------------------------------------------------------------------------
# Windows Taskbar Progress (ITaskbarList3)
# ---------------------------------------------------------------------------
class _TaskbarProgress:
    """Set progress in the Windows taskbar button. No-op on non-Windows."""
    TBPF_NOPROGRESS    = 0
    TBPF_INDETERMINATE = 0x1
    TBPF_NORMAL        = 0x2
    TBPF_ERROR         = 0x4
    TBPF_PAUSED        = 0x8

    def __init__(self) -> None:
        self._taskbar = None
        if platform.system() != "Windows":
            return
        try:
            import ctypes
            import ctypes.wintypes
            self._ctypes = ctypes
            taskbar_clsid = ctypes.POINTER(ctypes.c_int)
            self._taskbar = ctypes.windll.shell32.SHGetPropertyStoreForWindow  # placeholder
            # Use CoCreateInstance via ctypes
            import comtypes.client  # type: ignore
            self._taskbar = comtypes.client.CreateObject(
                "{56FDF344-FD6D-11d0-958A-006097C9A090}",
                interface=comtypes.gen.TaskbarLib.ITaskbarList3  # type: ignore
            )
        except Exception:
            self._taskbar = None

    def set(self, hwnd: int, current: int, total: int) -> None:
        if self._taskbar is None:
            return
        try:
            self._taskbar.SetProgressState(hwnd, self.TBPF_NORMAL)
            self._taskbar.SetProgressValue(hwnd, current, total)
        except Exception:
            pass

    def clear(self, hwnd: int) -> None:
        if self._taskbar is None:
            return
        try:
            self._taskbar.SetProgressState(hwnd, self.TBPF_NOPROGRESS)
        except Exception:
            pass

_TASKBAR = _TaskbarProgress()


# ---------------------------------------------------------------------------
# Sound on completion
# ---------------------------------------------------------------------------
def _play_done_sound() -> None:
    """Play a short beep when a download completes."""
    try:
        if platform.system() == "Windows":
            import winsound
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        elif platform.system() == "Linux":
            subprocess.run(["paplay", "/usr/share/sounds/freedesktop/stereo/complete.oga"],
                           capture_output=True, timeout=2)
        elif platform.system() == "Darwin":
            subprocess.run(["afplay", "/System/Library/Sounds/Glass.aiff"],
                           capture_output=True, timeout=2)
    except Exception:
        pass
# ---------------------------------------------------------------------------
class _BandwidthLimiter:
    """Simple token bucket for global download speed cap."""
    def __init__(self) -> None:
        self._lock     = threading.Lock()
        self._tokens   = 0.0
        self._last     = time.monotonic()
        self._limit    = 0   # bytes/sec; 0 = unlimited

    def set_limit(self, kbps: int) -> None:
        with self._lock:
            self._limit  = kbps * 1024
            self._tokens = float(self._limit)

    def consume(self, nbytes: int) -> None:
        """Block until we are allowed to send nbytes."""
        if self._limit <= 0:
            return
        while True:
            with self._lock:
                now = time.monotonic()
                elapsed = now - self._last
                self._last = now
                self._tokens = min(self._limit, self._tokens + elapsed * self._limit)
                if self._tokens >= nbytes:
                    self._tokens -= nbytes
                    return
                wait = (nbytes - self._tokens) / self._limit
            time.sleep(min(wait, 0.05))

_BW_LIMITER = _BandwidthLimiter()


def _perform_download(
    item: DownloadItem,
    on_progress: Callable[[DownloadItem], None],
    semaphore: threading.Semaphore,
    settings: Settings,
) -> None:
    """Main download dispatcher with auto-retry."""
    with semaphore:
        max_tries = (settings.max_retries + 1) if settings.auto_retry else 1
        for attempt in range(max_tries):
            _do_download(item, on_progress, settings)
            if item.status != DownloadStatus.FAILED:
                break
            # Only retry on transient network errors, not HTTP 4xx or user cancel
            _RETRIABLE = (
                "חיבור", "timeout", "פג זמן", "Connection", "Timeout",
                "ConnectionReset", "BrokenPipe", "RemoteDisconnected",
                "EOF", "IncompleteRead", "ChunkedEncodingError",
                "שגיאת חיבור", "שגיאת בקשה",
            )
            retriable = item.error_msg and any(k in item.error_msg for k in _RETRIABLE)
            if not retriable or attempt >= max_tries - 1:
                break
            item.retry_count += 1
            wait_sec = 2 ** attempt   # 1s, 2s, 4s …
            item.error_msg = f"מנסה שנית ({item.retry_count}/{settings.max_retries}) בעוד {wait_sec}s..."
            on_progress(item)
            # Interruptible sleep
            for _ in range(wait_sec * 10):
                if item._cancel_event.is_set():
                    break
                time.sleep(0.1)
            if item._cancel_event.is_set():
                item.status = DownloadStatus.CANCELLED
                on_progress(item)
                return
            # Reset for retry
            item.status       = DownloadStatus.PENDING
            item.error_msg    = ""
            item.progress     = 0.0
            item.downloaded_bytes = 0
            item.speed_bps    = 0.0
            item.eta_seconds  = -1.0
            item._last_ui_update = 0.0


def _do_download(
    item: DownloadItem,
    on_progress: Callable[[DownloadItem], None],
    settings: Settings,
) -> None:
    if item._cancel_event.is_set():
        item.status = DownloadStatus.CANCELLED
        on_progress(item)
        return

    # ── Schedule range check ──────────────────────────────────────────────
    if settings.schedule_range_enabled and not item.scheduled_at:
        try:
            now_t = datetime.now().time()
            start_t = datetime.strptime(settings.schedule_range_start, "%H:%M").time()
            end_t   = datetime.strptime(settings.schedule_range_end,   "%H:%M").time()
            in_range = (start_t <= now_t <= end_t) if start_t <= end_t else (now_t >= start_t or now_t <= end_t)
            if not in_range:
                item.status = DownloadStatus.QUEUED
                on_progress(item)
                while True:
                    if item._cancel_event.is_set():
                        item.status = DownloadStatus.CANCELLED
                        on_progress(item)
                        return
                    now_t = datetime.now().time()
                    in_range = (start_t <= now_t <= end_t) if start_t <= end_t else (now_t >= start_t or now_t <= end_t)
                    if in_range:
                        break
                    time.sleep(30)
        except Exception:
            pass  # bad time format — ignore scheduling

    # ── Scheduled download (specific time) ───────────────────────────────
    if item.scheduled_at and item.scheduled_at > datetime.now():
        item.status = DownloadStatus.QUEUED
        on_progress(item)
        while datetime.now() < item.scheduled_at:
            if item._cancel_event.is_set():
                item.status = DownloadStatus.CANCELLED
                on_progress(item)
                return
            time.sleep(1)

    # ── Disk space check ──────────────────────────────────────────────────
    if settings.disk_check_enabled:
        try:
            stat = shutil.disk_usage(item.save_dir)
            free_mb = stat.free // (1024 * 1024)
            # Warn if less than 500 MB free
            if free_mb < 500:
                item.status    = DownloadStatus.FAILED
                item.error_msg = f"מקום דיסק לא מספיק: {free_mb} MB פנוי (נדרש לפחות 500 MB)"
                on_progress(item)
                return
        except Exception:
            pass

    scheme = urlparse(item.url).scheme.lower()
    try:
        if scheme in ("ftp", "ftps"):
            _ftp_download(item, on_progress)
        elif _is_torrent(item.url):
            _torrent_download(item, on_progress, settings)
        elif YTDLP_OK and _is_media_url(item.url):
            _ytdlp_download(item, on_progress, settings)
        elif (settings.multipart and item.multipart
              and _supports_multipart(item.url)):
            _multipart_download(item, on_progress, settings)
        else:
            _http_download(item, on_progress, settings)
    finally:
        _release_name(item.filename, item.save_dir)

    # Post-download steps
    if item.status == DownloadStatus.DONE:
        _post_process(item, on_progress, settings)


def _http_download(item: DownloadItem, on_progress: Callable,
                   settings: Settings | None = None) -> None:
    """Single-stream HTTP/HTTPS download with resume support and real-time speed."""
    item.status = DownloadStatus.DOWNLOADING
    on_progress(item)

    resume_from = _load_resume_bytes(item)
    tmp = item.destination.with_suffix(item.destination.suffix + ".part")
    downloaded = resume_from

    # Rolling speed window: list of (monotonic_time, bytes)
    speed_samples: collections.deque[tuple[float, int]] = collections.deque()
    window_bytes = 0

    session = _get_session()
    try:
        headers: dict[str, str] = {}
        if resume_from > 0 and tmp.exists():
            headers["Range"] = f"bytes={resume_from}-"

        resp = session.get(item.url, stream=True,
                           timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
                           headers=headers, allow_redirects=True)
        try:
            if resp.status_code == 416:   # Range not satisfiable → restart
                resp.close()              # ← close old response before new request
                resume_from = 0
                downloaded  = 0
                resp = session.get(item.url, stream=True,
                                   timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
                                   allow_redirects=True)
            resp.raise_for_status()

            cl = _parse_cl(resp.headers.get("Content-Length", ""))
            item.total_bytes = (cl + resume_from) if cl > 0 else 0

            # Max file size guard
            max_mb = settings.max_file_size_mb if settings else 0
            if max_mb > 0 and item.total_bytes > max_mb * 1_048_576:
                item.status = DownloadStatus.FAILED
                item.error_msg = f"הקובץ גדול מ-{max_mb} MB — ביטול"
                on_progress(item)
                return

            mode = "ab" if resume_from > 0 else "wb"
            with tmp.open(mode, buffering=CHUNK_SIZE * 2) as fh:
                for chunk in resp.iter_content(chunk_size=CHUNK_SIZE):
                    if item._cancel_event.is_set():
                        item.status = DownloadStatus.CANCELLED
                        _save_resume_state(item)
                        return
                    while item._pause_event.is_set():
                        if item._cancel_event.is_set():
                            item.status = DownloadStatus.CANCELLED
                            _save_resume_state(item)
                            return
                        time.sleep(0.05)
                    if not chunk:
                        continue

                    t_now = time.monotonic()
                    fh.write(chunk)
                    cl_len = len(chunk)
                    downloaded += cl_len

                    # Global bandwidth cap
                    _BW_LIMITER.consume(cl_len)

                    # --- Real-time speed: rolling window ---
                    speed_samples.append((t_now, cl_len))
                    window_bytes += cl_len
                    cutoff = t_now - SPEED_WINDOW_SEC
                    while speed_samples and speed_samples[0][0] < cutoff:
                        _, ob = speed_samples.popleft()
                        window_bytes -= ob
                    if len(speed_samples) > 1:
                        elapsed = t_now - speed_samples[0][0]
                        item.speed_bps = window_bytes / max(elapsed, 1e-9)
                    elif cl_len > 0:
                        item.speed_bps = cl_len / max(t_now - (speed_samples[0][0] if speed_samples else t_now), 1e-9)

                    item.downloaded_bytes = downloaded
                    item.progress = (downloaded / item.total_bytes * 100) if item.total_bytes > 0 else 0.0

                    if item.total_bytes > 0 and item.speed_bps > 0:
                        rem = item.total_bytes - downloaded
                        item.eta_seconds = rem / item.speed_bps if rem > 0 else -1.0
                    else:
                        item.eta_seconds = -1.0

                    if t_now - item._last_ui_update >= PROGRESS_THROTTLE:
                        item._last_ui_update = t_now
                        on_progress(item)

                    _throttle_sleep(item.throttle_bps, cl_len, t_now)

        finally:
            resp.close()

        tmp.replace(item.destination)
        item.progress = 100.0
        item.downloaded_bytes = item.total_bytes if item.total_bytes > 0 else downloaded
        item.status = DownloadStatus.DONE
        item.speed_bps = 0.0
        item.eta_seconds = -1.0
        _clear_resume_state(item)

    except requests.exceptions.Timeout:
        item.status = DownloadStatus.FAILED
        item.error_msg = "פג זמן החיבור"
        _save_resume_state(item)
    except requests.exceptions.HTTPError as e:
        item.status = DownloadStatus.FAILED
        item.error_msg = f"HTTP {e.response.status_code}"
    except requests.exceptions.ConnectionError:
        item.status = DownloadStatus.FAILED
        item.error_msg = "שגיאת חיבור לרשת"
        _save_resume_state(item)
    except requests.exceptions.RequestException as e:
        item.status = DownloadStatus.FAILED
        item.error_msg = f"שגיאת בקשה: {e}"
    except OSError as e:
        item.status = DownloadStatus.FAILED
        item.error_msg = f"שגיאת שמירה: {e.strerror or e}"
    finally:
        if tmp.exists() and item.status != DownloadStatus.DONE:
            try:
                tmp.unlink()
            except OSError:
                pass
        on_progress(item)


_multipart_cache: dict[str, bool] = {}
_multipart_cache_lock = threading.Lock()

def _supports_multipart(url: str) -> bool:
    """Check if server supports byte ranges. Result is cached per URL."""
    with _multipart_cache_lock:
        if url in _multipart_cache:
            return _multipart_cache[url]
    try:
        session = _get_session()
        resp = session.head(url, timeout=(5, 8), allow_redirects=True)
        cl  = _parse_cl(resp.headers.get("Content-Length",""))
        ar  = resp.headers.get("Accept-Ranges","").lower()
        result = ar == "bytes" and cl >= MULTIPART_MIN
    except Exception:
        result = False
    with _multipart_cache_lock:
        _multipart_cache[url] = result
    return result


def _multipart_download(item: DownloadItem, on_progress: Callable,
                        settings: Settings | None = None) -> None:
    """Download in N parallel segments then merge — maximises bandwidth.
    Supports pause/resume and global bandwidth limiter.
    """
    item.status = DownloadStatus.DOWNLOADING
    on_progress(item)

    parts_dir: Path | None = None
    try:
        session = _get_session()
        resp = session.head(item.url, timeout=(5, 8), allow_redirects=True)
        resp.raise_for_status()
        total = _parse_cl(resp.headers.get("Content-Length",""))
        if total <= 0:
            return _http_download(item, on_progress, settings)

        item.total_bytes = total

        # Max file size guard
        max_mb = settings.max_file_size_mb if settings else 0
        if max_mb > 0 and total > max_mb * 1_048_576:
            item.status = DownloadStatus.FAILED
            item.error_msg = f"הקובץ גדול מ-{max_mb} MB — ביטול"
            on_progress(item)
            return

        n_parts   = min(MULTIPART_PARTS, max(1, total // (CHUNK_SIZE * 2)))
        part_size = total // n_parts
        parts_dir = item.destination.parent / f".{item.filename}.parts"
        parts_dir.mkdir(parents=True, exist_ok=True)

        # Shared speed tracking across all part threads
        speed_lock   = threading.Lock()
        speed_samples: collections.deque[tuple[float, int]] = collections.deque()
        window_bytes = [0]
        downloaded_total = [0]
        part_errors: list[str] = []

        def _update_speed(now: float, nbytes: int) -> None:
            with speed_lock:
                speed_samples.append((now, nbytes))
                window_bytes[0] += nbytes
                downloaded_total[0] += nbytes
                cutoff = now - SPEED_WINDOW_SEC
                while speed_samples and speed_samples[0][0] < cutoff:
                    _, ob = speed_samples.popleft()
                    window_bytes[0] -= ob
                if len(speed_samples) > 1:
                    elapsed = now - speed_samples[0][0]
                    item.speed_bps = window_bytes[0] / max(elapsed, 1e-9)
                item.downloaded_bytes = downloaded_total[0]
                item.progress = downloaded_total[0] / total * 100
                if item.total_bytes > 0 and item.speed_bps > 0:
                    rem = item.total_bytes - downloaded_total[0]
                    item.eta_seconds = rem / item.speed_bps if rem > 0 else -1.0

        def _download_part(idx: int, start: int, end: int) -> None:
            part_path = parts_dir / f"part{idx:04d}"  # type: ignore[operator]
            hdrs = {"Range": f"bytes={start}-{end}"}
            try:
                with session.get(item.url, stream=True,
                                 timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
                                 headers=hdrs) as r:
                    r.raise_for_status()
                    with part_path.open("wb", buffering=CHUNK_SIZE * 2) as fh:
                        for chunk in r.iter_content(chunk_size=CHUNK_SIZE):
                            if item._cancel_event.is_set():
                                return
                            # Pause support in multipart
                            while item._pause_event.is_set():
                                if item._cancel_event.is_set():
                                    return
                                time.sleep(0.05)
                            if not chunk:
                                continue
                            fh.write(chunk)
                            t_now = time.monotonic()
                            cl_len = len(chunk)
                            # Global bandwidth cap in multipart
                            _BW_LIMITER.consume(cl_len)
                            _update_speed(t_now, cl_len)
                            if t_now - item._last_ui_update >= PROGRESS_THROTTLE:
                                item._last_ui_update = t_now
                                on_progress(item)
            except Exception as e:
                logger.error("Multipart part %d failed: %s", idx, e)
                with speed_lock:
                    part_errors.append(f"חלק {idx}: {e}")

        ranges = []
        for i in range(n_parts):
            s = i * part_size
            e = (s + part_size - 1) if i < n_parts - 1 else (total - 1)
            ranges.append((i, s, e))

        with ThreadPoolExecutor(max_workers=n_parts) as pool:
            futures = [pool.submit(_download_part, i, s, e) for i, s, e in ranges]
            for f in futures:
                f.result()

        if item._cancel_event.is_set():
            item.status = DownloadStatus.CANCELLED
            return

        if part_errors:
            # Fall back to single-stream on partial failure
            logger.warning("Multipart had %d failed parts, falling back", len(part_errors))
            item.downloaded_bytes = 0
            item.progress = 0.0
            item.speed_bps = 0.0
            shutil.rmtree(str(parts_dir), ignore_errors=True)
            parts_dir = None
            return _http_download(item, on_progress, settings)

        # Merge parts
        item.status = DownloadStatus.MERGING
        item.speed_bps = 0.0
        on_progress(item)
        tmp = item.destination.with_suffix(item.destination.suffix + ".part")
        with tmp.open("wb", buffering=CHUNK_SIZE * 4) as out:
            for i in range(n_parts):
                part_path = parts_dir / f"part{i:04d}"
                with part_path.open("rb") as inp:
                    shutil.copyfileobj(inp, out, length=CHUNK_SIZE * 4)

        shutil.rmtree(str(parts_dir), ignore_errors=True)
        parts_dir = None
        tmp.replace(item.destination)
        item.progress = 100.0
        item.downloaded_bytes = total
        item.status = DownloadStatus.DONE
        item.speed_bps = 0.0
        item.eta_seconds = -1.0

    except Exception as e:
        item.status = DownloadStatus.FAILED
        item.error_msg = str(e)
        logger.error("Multipart download failed: %s", e, exc_info=True)
    finally:
        # Always clean up parts dir, even on unexpected exception
        if parts_dir is not None and parts_dir.exists():
            shutil.rmtree(str(parts_dir), ignore_errors=True)
        on_progress(item)


def _ftp_download(item: DownloadItem, on_progress: Callable) -> None:
    """FTP/FTPS download via ftplib — with pause and bandwidth limiting."""
    item.status = DownloadStatus.DOWNLOADING
    on_progress(item)

    tmp = item.destination.with_suffix(item.destination.suffix + ".part")
    downloaded = 0
    speed_window: collections.deque[tuple[float, int]] = collections.deque()
    window_bytes = 0

    try:
        parsed = urlparse(item.url)
        host   = parsed.hostname or ""
        port   = parsed.port or 21
        user   = parsed.username or "anonymous"
        passwd = parsed.password or "fetchpro@"
        path   = parsed.path

        cls = ftplib.FTP_TLS if item.url.lower().startswith("ftps://") else ftplib.FTP
        with cls() as ftp:
            ftp.connect(host, port, timeout=CONNECT_TIMEOUT)
            ftp.login(user, passwd)
            if isinstance(ftp, ftplib.FTP_TLS):
                ftp.prot_p()

            try:
                item.total_bytes = ftp.size(path) or 0
            except Exception:
                item.total_bytes = 0

            with tmp.open("wb") as fh:
                def _callback(data: bytes) -> None:
                    nonlocal downloaded, window_bytes
                    # Pause support in FTP
                    while item._pause_event.is_set():
                        if item._cancel_event.is_set():
                            raise ftplib.error_reply("cancelled")
                        time.sleep(0.05)
                    if item._cancel_event.is_set():
                        raise ftplib.error_reply("cancelled")
                    fh.write(data)
                    cl_len = len(data)
                    downloaded += cl_len
                    # Global bandwidth cap
                    _BW_LIMITER.consume(cl_len)
                    now = time.monotonic()
                    speed_window.append((now, cl_len))
                    window_bytes += cl_len
                    cutoff = now - 2.0
                    while speed_window and speed_window[0][0] < cutoff:
                        _, ob = speed_window.popleft()
                        window_bytes -= ob
                    if len(speed_window) > 1:
                        elapsed = now - speed_window[0][0]
                        item.speed_bps = window_bytes / max(elapsed, 1e-9)
                    item.downloaded_bytes = downloaded
                    item.progress = (downloaded / item.total_bytes * 100) if item.total_bytes > 0 else 0.0
                    if now - item._last_ui_update >= PROGRESS_THROTTLE:
                        item._last_ui_update = now
                        on_progress(item)

                ftp.retrbinary(f"RETR {path}", _callback, blocksize=CHUNK_SIZE)

        tmp.replace(item.destination)
        item.progress = 100.0
        item.downloaded_bytes = item.total_bytes if item.total_bytes > 0 else downloaded
        item.status = DownloadStatus.DONE
        item.speed_bps = 0.0
        item.eta_seconds = -1.0

    except Exception as e:
        if item._cancel_event.is_set():
            item.status = DownloadStatus.CANCELLED
        else:
            item.status = DownloadStatus.FAILED
            item.error_msg = str(e)
            logger.error("FTP download failed: %s", e)
    finally:
        if tmp.exists() and item.status != DownloadStatus.DONE:
            try:
                tmp.unlink()
            except OSError:
                pass
        on_progress(item)


def _is_media_url(url: str) -> bool:
    """Return True if yt-dlp should handle this URL."""
    # Known video/audio platforms
    media_domains = (
        "youtube.com", "youtu.be", "vimeo.com", "dailymotion.com",
        "twitch.tv", "soundcloud.com", "twitter.com", "x.com",
        "instagram.com", "tiktok.com", "facebook.com", "reddit.com",
        "bilibili.com", "niconico.jp", "nicovideo.jp", "rumble.com",
        "odysee.com", "bitchute.com", "kick.com", "streamable.com",
        "mixcloud.com", "bandcamp.com", "spotify.com",
    )
    host = urlparse(url).hostname or ""
    return any(d in host for d in media_domains)


# ─── yt-dlp format helpers ────────────────────────────────────────────────────

# Maps user-friendly format name → yt-dlp format selector
_YTDLP_FORMAT_MAP: dict[str, str] = {
    "best":      "bestvideo[ext=mp4]+bestaudio[ext=m4a]/bestvideo+bestaudio/best",
    "4k":        "bestvideo[height<=2160][ext=mp4]+bestaudio[ext=m4a]/best[height<=2160]",
    "1080p":     "bestvideo[height<=1080][ext=mp4]+bestaudio[ext=m4a]/best[height<=1080]",
    "720p":      "bestvideo[height<=720][ext=mp4]+bestaudio[ext=m4a]/best[height<=720]",
    "480p":      "bestvideo[height<=480][ext=mp4]+bestaudio[ext=m4a]/best[height<=480]",
    "360p":      "bestvideo[height<=360][ext=mp4]+bestaudio[ext=m4a]/best[height<=360]",
    "bestaudio": "bestaudio/best",
    "mp3":       "bestaudio/best",   # + postprocessor
    "m4a":       "bestaudio[ext=m4a]/bestaudio/best",
    "opus":      "bestaudio[ext=opus]/bestaudio/best",
    "wav":       "bestaudio/best",   # + postprocessor
}

# Audio formats that trigger post-processing extraction
_AUDIO_EXTRACT_FORMATS = {"mp3", "m4a", "opus", "wav", "aac", "flac"}


def _build_ytdlp_opts(item: "DownloadItem", hook: Callable, settings: "Settings") -> dict:
    """Build yt-dlp options dict based on item format preferences."""
    fmt = item.media_format or settings.ytdlp_default_format or "best"
    is_audio = item.media_is_audio or fmt in _AUDIO_EXTRACT_FORMATS

    # Base options
    opts: dict = {
        "outtmpl": str(item.save_dir / "%(title)s.%(ext)s"),
        "progress_hooks": [hook],
        "quiet": True,
        "no_warnings": True,
        "noplaylist": not item.media_playlist,
        "writethumbnail": False,
        "writeinfojson": False,
        "merge_output_format": "mp4",
        "concurrent_fragment_downloads": 4,
    }

    # Proxy support
    if settings.proxy_enabled and settings.proxy_host:
        proxy_url = (
            f"{settings.proxy_type}://"
            f"{settings.proxy_user + ':' + settings.proxy_pass + '@' if settings.proxy_user else ''}"
            f"{settings.proxy_host}:{settings.proxy_port}"
        )
        opts["proxy"] = proxy_url

    if is_audio:
        # Audio extraction mode
        audio_fmt = fmt if fmt in _AUDIO_EXTRACT_FORMATS else "mp3"
        opts["format"] = _YTDLP_FORMAT_MAP.get("bestaudio", "bestaudio/best")
        opts["postprocessors"] = [
            {
                "key": "FFmpegExtractAudio",
                "preferredcodec": audio_fmt,
                "preferredquality": "0",  # best quality
            }
        ]
        if settings.ytdlp_embed_thumbnail:
            opts["postprocessors"].append({"key": "EmbedThumbnail"})
            opts["writethumbnail"] = True
        if settings.ytdlp_add_metadata:
            opts["postprocessors"].append({"key": "FFmpegMetadata", "add_metadata": True})
        opts["outtmpl"] = str(item.save_dir / "%(title)s.%(ext)s")
    else:
        # Video mode
        opts["format"] = _YTDLP_FORMAT_MAP.get(fmt, fmt)

    return opts


def _ytdlp_fetch_formats(url: str) -> list[dict]:
    """Fetch available formats for a URL (blocking). Returns list of format dicts."""
    if not YTDLP_OK:
        return []
    try:
        with yt_dlp.YoutubeDL({"quiet": True, "no_warnings": True}) as ydl:
            info = ydl.extract_info(url, download=False)
            if not info:
                return []
            formats = info.get("formats", [])
            title = info.get("title", "")
            duration = info.get("duration", 0)
            thumbnail = info.get("thumbnail", "")
            return formats, title, duration, thumbnail
    except Exception as e:
        logger.warning("yt-dlp format fetch failed: %s", e)
        return [], "", 0, ""


def _ytdlp_download(item: "DownloadItem", on_progress: Callable,
                    settings: "Settings | None" = None) -> None:
    """yt-dlp based download — supports video quality selection and audio extraction."""
    from types import SimpleNamespace
    if settings is None:
        settings = Settings()

    item.status = DownloadStatus.DOWNLOADING
    on_progress(item)

    def _hook(d: dict) -> None:
        if item._cancel_event.is_set():
            raise yt_dlp.utils.DownloadCancelled()
        # Pause support
        while item._pause_event.is_set():
            if item._cancel_event.is_set():
                raise yt_dlp.utils.DownloadCancelled()
            time.sleep(0.1)
        if d["status"] == "downloading":
            item.downloaded_bytes = d.get("downloaded_bytes") or 0
            item.total_bytes      = d.get("total_bytes") or d.get("total_bytes_estimate") or 0
            item.speed_bps        = d.get("speed") or 0.0
            item.eta_seconds      = d.get("eta") or -1.0
            if item.total_bytes > 0:
                item.progress = item.downloaded_bytes / item.total_bytes * 100
            now = time.monotonic()
            item._last_progress_bytes = item.downloaded_bytes
            item._last_progress_time  = now
            if now - item._last_ui_update >= PROGRESS_THROTTLE:
                item._last_ui_update = now
                on_progress(item)
        elif d["status"] == "finished":
            item.progress    = 100.0
            item.speed_bps   = 0.0
            item.eta_seconds = -1.0
            on_progress(item)

    opts = _build_ytdlp_opts(item, _hook, settings)
    try:
        with yt_dlp.YoutubeDL(opts) as ydl:
            info = ydl.extract_info(item.url, download=True)
            if info:
                # For playlists, info is the playlist; for singles, it's the video
                entry = info
                if info.get("_type") == "playlist":
                    entries = info.get("entries") or []
                    entry = entries[0] if entries else info
                fname = ydl.prepare_filename(entry)
                item.filename   = Path(fname).name
                item.media_title = info.get("title", item.filename)
                # Audio: extension changes after post-processing
                if item.media_is_audio or (item.media_format in _AUDIO_EXTRACT_FORMATS):
                    audio_fmt = item.media_format if item.media_format in _AUDIO_EXTRACT_FORMATS else "mp3"
                    stem = Path(item.filename).stem
                    item.filename = f"{stem}.{audio_fmt}"
        item.status = DownloadStatus.DONE
    except yt_dlp.utils.DownloadCancelled:
        item.status = DownloadStatus.CANCELLED
    except Exception as e:
        item.status    = DownloadStatus.FAILED
        item.error_msg = str(e)
        logger.error("yt-dlp download failed %s: %s", item.url, e)
    finally:
        on_progress(item)



# ---------------------------------------------------------------------------
# BitTorrent engine  (requires: pip install libtorrent)
# ---------------------------------------------------------------------------

# Shared libtorrent session — one per process, created lazily
_lt_session: "lt.session | None" = None
_lt_session_lock = threading.Lock()


def _get_lt_session() -> "lt.session":
    """Return (or create) the shared libtorrent session."""
    global _lt_session
    with _lt_session_lock:
        if _lt_session is None:
            settings = lt.settings_pack()
            settings[lt.settings_pack.alert_mask] = (
                lt.alert.category_t.status_notification |
                lt.alert.category_t.error_notification
            )
            settings[lt.settings_pack.listen_interfaces] = "0.0.0.0:6881"
            _lt_session = lt.session(settings)
        return _lt_session


def _is_torrent(url: str) -> bool:
    """Return True if url is a magnet link or ends with .torrent."""
    u = url.strip()
    return u.startswith("magnet:") or u.lower().endswith(".torrent")


def _torrent_download(
    item: DownloadItem,
    on_progress: Callable[[DownloadItem], None],
    settings: "Settings",
) -> None:
    """Download via BitTorrent (magnet link or .torrent file/URL)."""
    if not LIBTORRENT_OK:
        item.status = DownloadStatus.FAILED
        item.error_msg = "libtorrent לא מותקן — הרץ: pip install libtorrent"
        on_progress(item)
        return

    item.status = DownloadStatus.DOWNLOADING
    on_progress(item)

    ses = _get_lt_session()

    # Apply bandwidth limits from settings
    sp = lt.settings_pack()
    dl_limit = (settings.torrent_max_dl_kbps * 1024) if settings.torrent_max_dl_kbps > 0 else 0
    ul_limit = (settings.torrent_max_ul_kbps * 1024) if settings.torrent_max_ul_kbps > 0 else 0
    sp[lt.settings_pack.download_rate_limit] = dl_limit
    sp[lt.settings_pack.upload_rate_limit]   = ul_limit
    ses.apply_settings(sp)

    save_dir = str(item.save_dir)
    handle: "lt.torrent_handle | None" = None

    try:
        url = item.url.strip()

        if url.startswith("magnet:"):
            # Magnet link
            params = lt.parse_magnet_uri(url)
            params.save_path = save_dir
            handle = ses.add_torrent(params)
            item.torrent_is_magnet = True
            # Wait for metadata (DHT lookup)
            item.status = DownloadStatus.QUEUED
            on_progress(item)
            timeout_at = time.monotonic() + 120  # 2 min max for metadata
            while not handle.has_metadata():
                if item._cancel_event.is_set():
                    raise RuntimeError("cancelled")
                if time.monotonic() > timeout_at:
                    raise RuntimeError("timeout — לא נמצאו מקורות לטורנט זה")
                time.sleep(0.5)
            # Now we have metadata — update filename
            ti = handle.torrent_file()
            item.filename = ti.name() if ti else (item.filename or "torrent")
            item.total_bytes = ti.total_size() if ti else 0
            item.status = DownloadStatus.DOWNLOADING
            on_progress(item)

        elif url.lower().endswith(".torrent"):
            # Fetch the .torrent file first, then add it
            import tempfile as _tmp
            torrent_data: bytes
            if url.startswith(("http://", "https://")):
                sess = _get_session()
                resp = sess.get(url, timeout=(10, 30))
                resp.raise_for_status()
                torrent_data = resp.content
            else:
                # Local file path
                torrent_data = Path(url).read_bytes()

            ti = lt.torrent_info(lt.bdecode(torrent_data))
            atp = lt.add_torrent_params()
            atp.ti = ti
            atp.save_path = save_dir
            handle = ses.add_torrent(atp)
            item.filename = ti.name()
            item.total_bytes = ti.total_size()
        else:
            raise ValueError(f"לא זוהה כטורנט: {url}")

        # ── Download loop ─────────────────────────────────────────────
        speed_window: collections.deque[tuple[float, int]] = collections.deque()
        prev_bytes = 0

        while True:
            if item._cancel_event.is_set():
                raise RuntimeError("cancelled")

            # Pause support
            if item._pause_event.is_set():
                handle.pause()
                item.status = DownloadStatus.PAUSED
                on_progress(item)
                while item._pause_event.is_set():
                    if item._cancel_event.is_set():
                        raise RuntimeError("cancelled")
                    time.sleep(0.2)
                handle.resume()
                item.status = DownloadStatus.DOWNLOADING
                on_progress(item)

            s = handle.status()

            # Progress
            item.progress        = s.progress * 100.0
            item.downloaded_bytes = s.total_done
            item.total_bytes      = s.total_wanted
            item.torrent_seeds    = s.num_seeds
            item.torrent_peers    = s.num_peers
            num_pieces   = handle.torrent_file().num_pieces() if handle.has_metadata() else 0
            done_pieces  = s.num_pieces
            item.torrent_pieces  = f"{done_pieces}/{num_pieces}" if num_pieces else ""

            # Speed (rolling 3 s window)
            now = time.monotonic()
            delta_bytes = s.total_done - prev_bytes
            if delta_bytes > 0:
                speed_window.append((now, delta_bytes))
            prev_bytes = s.total_done
            cutoff = now - 3.0
            while speed_window and speed_window[0][0] < cutoff:
                speed_window.popleft()
            if speed_window:
                elapsed = now - speed_window[0][0] if len(speed_window) > 1 else 1.0
                total_w = sum(b for _, b in speed_window)
                item.speed_bps = total_w / max(elapsed, 0.001)
            else:
                item.speed_bps = float(s.download_rate)

            # ETA
            if item.speed_bps > 0 and item.total_bytes > item.downloaded_bytes:
                item.eta_seconds = (item.total_bytes - item.downloaded_bytes) / item.speed_bps
            else:
                item.eta_seconds = -1.0

            if now - item._last_ui_update >= PROGRESS_THROTTLE:
                item._last_ui_update = now
                on_progress(item)

            # Check if download is complete
            if s.is_seeding or s.progress >= 1.0:
                item.progress        = 100.0
                item.downloaded_bytes = item.total_bytes
                item.speed_bps       = 0.0
                item.eta_seconds     = -1.0
                break

            # Check for error states
            if s.state == lt.torrent_status.error:
                raise RuntimeError(f"libtorrent שגיאה: {s.errc.message()}")

            time.sleep(0.5)

        # ── Seeding phase (optional) ──────────────────────────────────
        if settings.torrent_seeding and handle is not None:
            item.status = DownloadStatus.SEEDING
            on_progress(item)
            seed_start = time.monotonic()
            while True:
                if item._cancel_event.is_set():
                    break
                s = handle.status()
                ratio = s.all_time_upload / max(s.all_time_download, 1)
                item.speed_bps = float(s.upload_rate)
                now = time.monotonic()
                if now - item._last_ui_update >= PROGRESS_THROTTLE:
                    item._last_ui_update = now
                    on_progress(item)
                if ratio >= settings.torrent_seed_ratio:
                    break
                time.sleep(1.0)

        item.status   = DownloadStatus.DONE
        item.speed_bps = 0.0
        on_progress(item)

    except RuntimeError as e:
        if "cancelled" in str(e).lower() or item._cancel_event.is_set():
            item.status = DownloadStatus.CANCELLED
        else:
            item.status    = DownloadStatus.FAILED
            item.error_msg = str(e)
        on_progress(item)
    except Exception as e:
        item.status    = DownloadStatus.FAILED
        item.error_msg = str(e)
        logger.error("Torrent download failed %s: %s", item.url, e)
        on_progress(item)
    finally:
        # Remove torrent from session (keep files on disk)
        if handle is not None:
            try:
                ses.remove_torrent(handle)
            except Exception:
                pass


def _open_torrent_file(save_dir: str) -> str | None:
    """Open file dialog to pick a .torrent file; return path or None."""
    path = filedialog.askopenfilename(
        title=_t("select_torrent"),
        initialdir=save_dir,
        filetypes=[("Torrent files", "*.torrent"), ("All files", "*.*")],
    )
    return path or None


def _virustotal_scan(item: DownloadItem, on_progress: Callable,
                     settings: Settings) -> bool:
    """
    Scan a downloaded file via VirusTotal API v3.
    Returns True = clean / unknown, False = threat detected (file quarantined).
    Requires settings.virustotal_api_key to be set.
    """
    if not settings.virustotal_enabled:
        return True   # scanning disabled by user
    api_key = settings.virustotal_api_key.strip()
    if not api_key or not item.destination.exists():
        return True   # no key → skip scan, treat as clean

    item.status = DownloadStatus.SCANNING
    item.error_msg = ""
    on_progress(item)

    VT_BASE = "https://www.virustotal.com/api/v3"
    headers = {"x-apikey": api_key}

    try:
        # Step 1: compute SHA-256 (reuse if already computed)
        sha256 = item.actual_hash if item.hash_algo == "sha256" else ""
        if not sha256:
            sha256 = _compute_hash(item.destination, "sha256")

        session = _get_session()

        # Step 2: check if VT already has a report for this hash
        resp = session.get(f"{VT_BASE}/files/{sha256}",
                           headers=headers, timeout=(10, 20))

        if resp.status_code == 200:
            # VT knows this file — read existing report
            data = resp.json()
        elif resp.status_code == 404:
            # VT has no record — upload the file for scanning
            item.error_msg = "מעלה לסריקה..."
            on_progress(item)

            file_size = item.destination.stat().st_size
            if file_size > 32 * 1024 * 1024:
                # Files > 32 MB require a special upload URL
                url_resp = session.get(f"{VT_BASE}/files/upload_url",
                                       headers=headers, timeout=(10, 20))
                url_resp.raise_for_status()
                upload_url = url_resp.json()["data"]
            else:
                upload_url = f"{VT_BASE}/files"

            with item.destination.open("rb") as fh:
                upload_resp = session.post(
                    upload_url, headers=headers,
                    files={"file": (item.filename, fh, "application/octet-stream")},
                    timeout=(30, 300),
                )
            upload_resp.raise_for_status()
            analysis_id = upload_resp.json()["data"]["id"]

            # Step 3: poll analysis result (up to 3 min)
            item.error_msg = "מחכה לתוצאות סריקה..."
            on_progress(item)
            deadline = time.monotonic() + 180
            while time.monotonic() < deadline:
                if item._cancel_event.is_set():
                    return True   # cancelled → treat as clean
                time.sleep(15)
                poll = session.get(f"{VT_BASE}/analyses/{analysis_id}",
                                   headers=headers, timeout=(10, 30))
                if poll.status_code == 200:
                    poll_data = poll.json()
                    status = poll_data.get("data", {}).get("attributes", {}).get("status", "")
                    if status == "completed":
                        # Get the file report by SHA-256
                        sha_resp = session.get(f"{VT_BASE}/files/{sha256}",
                                               headers=headers, timeout=(10, 30))
                        if sha_resp.status_code == 200:
                            data = sha_resp.json()
                            break
            else:
                # Timed out waiting
                item.error_msg = ""
                return True

        else:
            resp.raise_for_status()
            return True

        # Step 4: evaluate result
        stats = (data.get("data", {})
                     .get("attributes", {})
                     .get("last_analysis_stats", {}))
        malicious  = stats.get("malicious",  0)
        suspicious = stats.get("suspicious", 0)
        total      = sum(stats.values()) or 1

        if malicious >= 3:
            # Threat confirmed — quarantine (rename file)
            q_path = item.destination.with_suffix(item.destination.suffix + ".quarantine")
            try:
                item.destination.rename(q_path)
            except Exception:
                pass
            item.status    = DownloadStatus.FAILED
            item.error_msg = (
                f"⚠ VirusTotal: {malicious}/{total} מנועים זיהו איום! "
                f"הקובץ הועבר לקוורנטינה."
            )
            on_progress(item)
            logger.warning("VT threat detected: %s — %d/%d malicious", item.filename, malicious, total)
            return False
        elif malicious > 0 or suspicious >= 3:
            # Suspicious — warn but keep file
            item.error_msg = (
                f"⚠ VirusTotal: {malicious} זיהויים, {suspicious} חשודים מתוך {total}. "
                f"בדוק לפני פתיחה."
            )
            logger.warning("VT suspicious: %s — %d malicious %d suspicious/%d", item.filename, malicious, suspicious, total)
            # Keep as DONE but with warning in error_msg
            return True
        else:
            item.error_msg = f"✓ VirusTotal: נקי ({total} מנועים)"
            logger.info("VT clean: %s (%d engines)", item.filename, total)
            return True

    except Exception as e:
        logger.warning("VirusTotal scan failed for %s: %s", item.filename, e)
        item.error_msg = f"VT שגיאה: {e}"
        return True   # scan error → don't block the download


def _post_process(
    item: DownloadItem,
    on_progress: Callable,
    settings: Settings,
) -> None:
    """Hash check, auto-extract, auto-categorize, auto-open, notification."""

    # Hash verification
    algo = item.hash_algo or (settings.default_hash_algo if settings.verify_hash else "")
    if algo and item.destination.exists():
        item.status = DownloadStatus.HASHING
        on_progress(item)
        item.actual_hash = _compute_hash(item.destination, algo)
        if item.expected_hash and item.actual_hash.lower() != item.expected_hash.lower():
            item.status = DownloadStatus.FAILED
            item.error_msg = f"Hash אינו תואם! צפוי: {item.expected_hash[:12]}… קיבלנו: {item.actual_hash[:12]}…"
            on_progress(item)
            return
        item.status = DownloadStatus.DONE

    # Auto-categorize (move to type subfolder)
    if settings.auto_categorize and item.destination.exists():
        ext = item.destination.suffix.lstrip(".").lower()
        sub = CATEGORY_MAP.get(ext)
        if sub:
            dest_dir = item.save_dir / sub
            dest_dir.mkdir(parents=True, exist_ok=True)
            new_path = dest_dir / item.filename
            if not new_path.exists():
                item.destination.rename(new_path)
                item.save_dir = dest_dir

    # Auto-extract
    if (item.auto_extract or settings.auto_extract) and item.destination.exists():
        item.status = DownloadStatus.EXTRACTING
        on_progress(item)
        ext_dir = item.destination.parent / item.destination.stem
        if _try_extract(item.destination, ext_dir):
            logger.info("Extracted %s → %s", item.filename, ext_dir)
        item.status = DownloadStatus.DONE
        on_progress(item)

    # VirusTotal scan
    if settings.virustotal_enabled and settings.virustotal_api_key.strip() and item.destination.exists():
        clean = _virustotal_scan(item, on_progress, settings)
        if not clean:
            return   # file quarantined — stop here
        # If scan returned a warning (error_msg set) keep it visible
        item.status = DownloadStatus.DONE

    # Notification
    if settings.notify_done:
        _send_notification(APP_NAME, f"הורדה הושלמה: {item.filename}")

    # Sound
    if settings.sound_on_done:
        threading.Thread(target=_play_done_sound, daemon=True).start()

    # Auto-open
    if (item.auto_open or settings.auto_open) and item.destination.exists():
        _open_file(item.destination)


# ---------------------------------------------------------------------------
# URL Verification engine
# ---------------------------------------------------------------------------
class VerifyResult(Enum):
    PENDING = "pending"
    OK      = "ok"
    WARNING = "warning"
    FAILED  = "failed"


@dataclass
class VerifyInfo:
    url:            str
    result:         VerifyResult = VerifyResult.PENDING
    status_code:    int  = 0
    content_type:   str  = ""
    content_length: int  = -1
    final_url:      str  = ""
    message:        str  = _t("verifying")


def _verify_url(url: str) -> VerifyInfo:
    info = VerifyInfo(url=url)
    scheme = urlparse(url).scheme.lower()

    if scheme in ("ftp","ftps"):
        try:
            parsed = urlparse(url)
            cls = ftplib.FTP_TLS if scheme == "ftps" else ftplib.FTP
            with cls() as ftp:
                ftp.connect(parsed.hostname or "", parsed.port or 21, timeout=6)
                ftp.login(parsed.username or "anonymous", parsed.password or "")
                try:
                    size = ftp.size(parsed.path) or -1
                except Exception:
                    size = -1
            info.result = VerifyResult.OK
            info.content_length = size
            size_str = f" ({_fmt_bytes(size)})" if size > 0 else ""
            info.message = f"FTP — קובץ נמצא{size_str}"
        except Exception as e:
            info.result = VerifyResult.FAILED
            info.message = str(e)
        return info

    try:
        session = _get_session()
        resp = session.head(url, timeout=(6, 10), allow_redirects=True)
        if resp.status_code in (405, 501):
            resp = session.get(url, timeout=(6, 10),
                               headers={"Range": "bytes=0-0"},
                               allow_redirects=True, stream=True)
            resp.close()

        info.status_code    = resp.status_code
        info.final_url      = resp.url
        info.content_type   = resp.headers.get("Content-Type","").split(";")[0].strip()
        info.content_length = _parse_cl(resp.headers.get("Content-Length",""))

        if 200 <= resp.status_code < 300:
            info.result = VerifyResult.OK
            sz = f" ({_fmt_bytes(info.content_length)})" if info.content_length > 0 else ""
            mp = " · Range✓" if resp.headers.get("Accept-Ranges","").lower()=="bytes" else ""
            info.message = f"HTTP {resp.status_code} — {info.content_type or 'unknown'}{sz}{mp}"
        elif 300 <= resp.status_code < 400:
            info.result  = VerifyResult.WARNING
            info.message = f"HTTP {resp.status_code} — redirect"
        elif resp.status_code == 404:
            info.result  = VerifyResult.FAILED
            info.message = "HTTP 404 — הקובץ לא נמצא"
        elif resp.status_code == 403:
            info.result  = VerifyResult.FAILED
            info.message = "HTTP 403 — גישה נדחתה"
        else:
            info.result  = VerifyResult.FAILED
            info.message = f"HTTP {resp.status_code} — שגיאת שרת"
    except requests.exceptions.Timeout:
        info.result  = VerifyResult.FAILED
        info.message = "פג זמן החיבור"
    except requests.exceptions.ConnectionError:
        info.result  = VerifyResult.FAILED
        info.message = "לא ניתן להגיע לשרת"
    except requests.exceptions.RequestException as e:
        info.result  = VerifyResult.FAILED
        info.message = f"שגיאת רשת: {e}"
    return info


# ---------------------------------------------------------------------------
# Download Card widget
# ---------------------------------------------------------------------------
class DownloadCard(tk.Frame):
    BAR_H = 8
    PAD   = 12

    def __init__(self, parent: tk.Widget, item: DownloadItem,
                 on_pause: Callable, on_resume: Callable,
                 on_cancel: Callable, on_retry: Callable,
                 on_remove: Callable, on_open: Callable,
                 on_up: Callable, on_down: Callable) -> None:
        super().__init__(parent, bg=T.BG_CARD, padx=self.PAD, pady=self.PAD)
        self._item = item
        self._cbs  = dict(pause=on_pause, resume=on_resume, cancel=on_cancel,
                          retry=on_retry, remove=on_remove,
                          open=on_open, up=on_up, down=on_down)
        self._shimmer_x  = 0
        self._shimmer_id: str | None = None
        self._speed_history: collections.deque[float] = collections.deque(maxlen=20)
        self._countdown_id:  str | None = None
        self._build()

    # ── Build ─────────────────────────────────────────────────────────
    def _build(self) -> None:
        self.configure(highlightthickness=1, highlightbackground=T.BORDER)

        top = tk.Frame(self, bg=T.BG_CARD)
        top.pack(fill="x", pady=(0, 6))

        # Priority arrows
        arrow_frame = tk.Frame(top, bg=T.BG_CARD)
        arrow_frame.pack(side="left", padx=(0, 4))
        self._up_btn   = self._make_btn(arrow_frame, "▲", T.TEXT_DIM, lambda: self._cbs["up"](self._item), size=9)
        self._down_btn = self._make_btn(arrow_frame, "▼", T.TEXT_DIM, lambda: self._cbs["down"](self._item), size=9)
        self._up_btn.pack()
        self._down_btn.pack()

        # Icon
        ext = self._item.filename.rsplit(".", 1)[-1].lower() if "." in self._item.filename else ""
        icons = {"pdf":"📄","zip":"📦","rar":"📦","7z":"📦","gz":"📦","tar":"📦",
                 "exe":"⚙️","dmg":"💿","iso":"💿","apk":"📱",
                 "mp4":"🎬","mkv":"🎬","avi":"🎬","mov":"🎬",
                 "mp3":"🎵","flac":"🎵","wav":"🎵","aac":"🎵",
                 "png":"🖼️","jpg":"🖼️","jpeg":"🖼️","gif":"🖼️","webp":"🖼️",
                 "csv":"📊","xlsx":"📊","doc":"📝","docx":"📝","txt":"📝"}
        icon_frame = tk.Frame(top, bg="#1a1f2e", width=36, height=36)
        icon_frame.pack(side="left", padx=6)
        icon_frame.pack_propagate(False)
        tk.Label(icon_frame, text=icons.get(ext,"📁"), bg="#1a1f2e",
                 font=("Segoe UI Emoji", 15)).place(relx=.5, rely=.5, anchor="center")

        # Name + URL
        nf = tk.Frame(top, bg=T.BG_CARD)
        nf.pack(side="left", fill="x", expand=True)

        name_row = tk.Frame(nf, bg=T.BG_CARD)
        name_row.pack(fill="x")
        self._name_lbl = tk.Label(name_row, text=self._item.filename,
                                  bg=T.BG_CARD, fg=T.TEXT_MAIN,
                                  font=("Consolas",10,"bold"),
                                  anchor="w")
        self._name_lbl.pack(side="left")
        # Priority badge
        pri = self._item.dl_priority
        if pri != "NORMAL":
            pri_color = "#ff4444" if pri == "HIGH" else "#888888"
            pri_text  = "⚡HIGH" if pri == "HIGH" else "⬇LOW"
            tk.Label(name_row, text=pri_text, bg=pri_color, fg="white",
                     font=("Consolas",7,"bold"), padx=4, pady=1).pack(side="left", padx=4)
        # Priority border
        if pri == "HIGH":
            self.configure(highlightbackground="#ff4444")
        elif pri == "LOW":
            self.configure(highlightbackground="#555555")

        short = (self._item.url[:68]+"…") if len(self._item.url)>68 else self._item.url
        tk.Label(nf, text=short, bg=T.BG_CARD, fg=T.TEXT_DIM,
                 font=("Consolas",8), anchor="w").pack(fill="x")
        if self._item.note:
            tk.Label(nf, text=f"📌 {self._item.note}", bg=T.BG_CARD, fg=T.YELLOW,
                     font=("Consolas",8), anchor="w").pack(fill="x")

        # Action buttons
        bf = tk.Frame(top, bg=T.BG_CARD)
        bf.pack(side="right", padx=4)
        self._pause_btn  = self._make_btn(bf, "⏸", T.YELLOW,   lambda: self._cbs["pause"](self._item))
        self._resume_btn = self._make_btn(bf, "▶",  T.GREEN,    lambda: self._cbs["resume"](self._item))
        self._retry_btn  = self._make_btn(bf, "↺",  T.ACCENT,   lambda: self._cbs["retry"](self._item))
        self._open_btn   = self._make_btn(bf, "📂", T.TEXT_DIM, lambda: self._cbs["open"](self._item))
        self._remove_btn = self._make_btn(bf, "🗑",  T.TEXT_DIM, lambda: self._cbs["remove"](self._item))
        self._cancel_btn = self._make_btn(bf, "✕",  T.RED,      lambda: self._cbs["cancel"](self._item))

        # Progress bar with shimmer canvas
        bar_outer = tk.Frame(self, bg=T.BG_CARD)
        bar_outer.pack(fill="x", pady=4)
        self._bar_bg = tk.Canvas(bar_outer, height=self.BAR_H,
                                 bg="#0d1520", highlightthickness=0)
        self._bar_bg.pack(fill="x")
        self._bar_fill    = self._bar_bg.create_rectangle(0, 0, 0, self.BAR_H, fill=T.ACCENT, outline="")
        self._bar_shimmer = self._bar_bg.create_rectangle(-60, 0, -30, self.BAR_H,
                                                           fill="#ffffff", outline="", stipple="gray50")
        # Speedometer mini-bars (10 columns)
        self._speed_bars: list[int] = []
        bar_row = tk.Frame(self, bg=T.BG_CARD)
        bar_row.pack(fill="x", pady=(0, 3))
        self._speed_canvas = tk.Canvas(bar_row, height=18, bg=T.BG_CARD, highlightthickness=0)
        self._speed_canvas.pack(side="left", fill="x", expand=True)

        # Info row
        info = tk.Frame(self, bg=T.BG_CARD)
        info.pack(fill="x")
        self._status_lbl = tk.Label(info, text="", bg=T.BG_CARD, fg=T.ACCENT,
                                    font=("Consolas",9,"bold"), anchor="w")
        self._status_lbl.pack(side="left")
        self._speed_lbl  = tk.Label(info, text="", bg=T.BG_CARD, fg=T.ACCENT,
                                    font=("Consolas",9,"bold"), anchor="w")
        self._speed_lbl.pack(side="left", padx=10)
        self._eta_lbl    = tk.Label(info, text="", bg=T.BG_CARD, fg=T.TEXT_MUTED,
                                    font=("Consolas",9))
        self._eta_lbl.pack(side="left")
        self._pct_lbl    = tk.Label(info, text="0%", bg=T.BG_CARD, fg=T.TEXT_DIM,
                                    font=("Consolas",9,"bold"))
        self._pct_lbl.pack(side="right", padx=10)
        self._size_lbl   = tk.Label(info, text="", bg=T.BG_CARD, fg=T.TEXT_DIM,
                                    font=("Consolas",9))
        self._size_lbl.pack(side="right")

        # Torrent: seeds/peers row (hidden unless torrent download)
        self._torrent_row = tk.Frame(self, bg=T.BG_CARD)
        self._torrent_lbl = tk.Label(self._torrent_row, text="",
                                      bg=T.BG_CARD, fg="#00cc88",
                                      font=("Consolas",8), anchor="w")
        self._torrent_lbl.pack(side="left", padx=4)

        # Countdown row (visible only for QUEUED items)
        self._countdown_frame = tk.Frame(self, bg=T.BG_CARD)
        self._countdown_lbl   = tk.Label(self._countdown_frame, text="",
                                          bg=T.BG_CARD, fg=T.YELLOW,
                                          font=("Consolas",9), anchor="w")
        self._countdown_lbl.pack(side="left")

        # Right-click context menu (bind on whole card)
        self.bind("<Button-3>", self._show_context_menu)
        for child in self.winfo_children():
            child.bind("<Button-3>", self._show_context_menu)

    def _show_context_menu(self, event: tk.Event) -> None:
        item = self._item
        menu = tk.Menu(self, tearoff=0, bg=T.BG_CARD, fg=T.TEXT_MAIN,
                       activebackground=T.ACCENT, activeforeground=T.BG_DEEP,
                       font=("Consolas", 9), bd=0, relief="flat")
        menu.add_command(label=_t("ctx_copy_url"),
                         command=lambda: (self.clipboard_clear(), self.clipboard_append(item.url)))
        menu.add_command(label=_t("ctx_open_folder"),
                         command=lambda: _open_file(item.save_dir))
        if item.status == DownloadStatus.DONE and item.destination.exists():
            menu.add_command(label=_t("ctx_open_file"),
                             command=lambda: _open_file(item.destination))
        menu.add_separator()
        if item.status in _ACTIVE:
            menu.add_command(label=_t("ctx_cancel"), command=lambda: self._cbs["cancel"](item))
        if item.status in (DownloadStatus.FAILED, DownloadStatus.CANCELLED):
            menu.add_command(label=_t("ctx_retry"), command=lambda: self._cbs["retry"](item))
            # Delete partial file option
            part = item.destination.with_suffix(item.destination.suffix + ".part")
            if part.exists():
                menu.add_command(label=_t("ctx_delete_partial"),
                                 command=lambda p=part: self._delete_partial(p))
        menu.add_separator()
        menu.add_command(label=_t("ctx_remove"), command=lambda: self._cbs["remove"](item))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _delete_partial(self, part: Path) -> None:
        if messagebox.askyesno("מחיקת קובץ חלקי",
                                f"למחוק את:\n{part.name}?", parent=self):
            try:
                part.unlink(missing_ok=True)
                self._item.downloaded_bytes = 0
                self._item.progress = 0.0
                self.refresh()
            except OSError as e:
                messagebox.showerror("שגיאה", str(e), parent=self)

    def _make_btn(self, parent: tk.Widget, text: str, color: str,
                  cmd: Callable, size: int = 12) -> tk.Label:
        btn = tk.Label(parent, text=text, bg=T.BG_CARD, fg=color,
                       font=("Segoe UI Emoji", size), cursor="hand2", padx=3)
        btn.pack(side="left")
        btn.bind("<Button-1>", lambda _: cmd())
        btn.bind("<Enter>",    lambda _: btn.configure(bg=T.BG_HOVER))
        btn.bind("<Leave>",    lambda _: btn.configure(bg=T.BG_CARD))
        return btn

    # ── Shimmer animation ─────────────────────────────────────────────
    def _start_shimmer(self) -> None:
        if self._shimmer_id:
            return
        self._shimmer_x = 0
        self._animate_shimmer()

    def _animate_shimmer(self) -> None:
        if not self.winfo_exists():
            return
        if self._item.status not in _ACTIVE:
            self._bar_bg.coords(self._bar_shimmer, -60, 0, -30, self.BAR_H)
            self._shimmer_id = None
            return
        w = self._bar_bg.winfo_width()
        self._shimmer_x = (self._shimmer_x + 12) % (w + 60)
        x0 = self._shimmer_x - 60
        x1 = self._shimmer_x
        fill_w = int(w * self._item.progress / 100) if self._item.progress > 0 else 0
        if x1 > 0 and x0 < fill_w:
            self._bar_bg.coords(self._bar_shimmer, x0, 0, min(x1, fill_w), self.BAR_H)
        else:
            self._bar_bg.coords(self._bar_shimmer, -60, 0, -30, self.BAR_H)
        self._shimmer_id = self.after(30, self._animate_shimmer)  # 33 fps

    def _stop_shimmer(self) -> None:
        if self._shimmer_id:
            try:
                self.after_cancel(self._shimmer_id)
            except Exception:
                pass
            self._shimmer_id = None
        self._bar_bg.coords(self._bar_shimmer, -60, 0, -30, self.BAR_H)

    # ── Speedometer mini-bars ─────────────────────────────────────────
    def _draw_speed_bars(self, color: str) -> None:
        """Draw a mini bargraph of recent speed history."""
        c = self._speed_canvas
        c.delete("all")
        if not self._speed_history or self._item.status not in _ACTIVE:
            return
        w = c.winfo_width()
        h = 18
        n = 20
        bar_w = max(2, (w // n) - 1)
        max_v = max(self._speed_history) or 1
        samples = list(self._speed_history)
        for i, v in enumerate(samples):
            bar_h = max(1, int(v / max_v * h))
            x0 = i * (bar_w + 1)
            x1 = x0 + bar_w
            y0 = h - bar_h
            alpha = 80 + int(175 * i / max(len(samples) - 1, 1))
            # We can't do real alpha in tkinter canvas easily, so use color intensity
            c.create_rectangle(x0, y0, x1, h, fill=color, outline="")

    # ── Countdown ticker for scheduled downloads ───────────────────────
    def _start_countdown(self) -> None:
        if self._countdown_id:
            return
        self._tick_countdown()

    def _tick_countdown(self) -> None:
        if not self.winfo_exists():
            return
        if self._item.status != DownloadStatus.QUEUED or not self._item.scheduled_at:
            self._countdown_frame.pack_forget()
            self._countdown_id = None
            return
        remaining = self._item.scheduled_at - datetime.now()
        if remaining.total_seconds() <= 0:
            self._countdown_lbl.configure(text=_t("status_starting"))
        else:
            total_s = int(remaining.total_seconds())
            h, rem = divmod(total_s, 3600)
            m, s   = divmod(rem, 60)
            if h > 0:
                txt = f"⏰ מתוזמן בעוד {h}:{m:02d}:{s:02d}"
            else:
                txt = f"⏰ מתוזמן בעוד {m:02d}:{s:02d}"
            scheduled_fmt = self._item.scheduled_at.strftime("%d/%m/%Y %H:%M")
            self._countdown_lbl.configure(text=f"{txt}  ({scheduled_fmt})")
        self._countdown_frame.pack(fill="x", pady=(2, 0))
        self._countdown_id = self.after(1000, self._tick_countdown)

    def _stop_countdown(self) -> None:
        if self._countdown_id:
            try:
                self.after_cancel(self._countdown_id)
            except Exception:
                pass
            self._countdown_id = None
        self._countdown_frame.pack_forget()

    # ── Refresh ───────────────────────────────────────────────────────
    def refresh(self) -> None:
        item   = self._item
        status = item.status
        color  = {
            DownloadStatus.DOWNLOADING: T.ACCENT,
            DownloadStatus.MERGING:     T.YELLOW,
            DownloadStatus.HASHING:     "#aa88ff",
            DownloadStatus.EXTRACTING:  T.YELLOW,
            DownloadStatus.SCANNING:    "#ff9900",
            DownloadStatus.DONE:        T.GREEN,
            DownloadStatus.FAILED:      T.RED,
            DownloadStatus.PAUSED:      T.YELLOW,
            DownloadStatus.QUEUED:      T.YELLOW,
            DownloadStatus.SEEDING:     "#00cc88",
        }.get(status, T.TEXT_DIM)

        self._status_lbl.configure(text=status.value, fg=color)
        retry_sfx = f" (ניסיון {item.retry_count})" if item.retry_count > 0 else ""
        # Torrent-specific: show pieces info instead of plain %
        if item.torrent_pieces:
            pct_text = f"{item.progress:.1f}% [{item.torrent_pieces}]{retry_sfx}"
        else:
            pct_text = f"{item.progress:.1f}%{retry_sfx}"
        self._pct_lbl.configure(text=pct_text)

        w = self._bar_bg.winfo_width()
        if w > 1:
            fw = max(2, int(w * item.progress / 100)) if item.progress > 0 else 0
            self._bar_bg.coords(self._bar_fill, 0, 0, fw, self.BAR_H)
            self._bar_bg.itemconfigure(self._bar_fill, fill=color)

        # Real-time speed display
        if status in _ACTIVE and item.speed_bps > 0:
            self._speed_history.append(item.speed_bps)
            spd_txt = _fmt_speed(item.speed_bps)
            self._speed_lbl.configure(text=f"▲ {spd_txt}", fg=T.ACCENT)
            eta = _fmt_eta(item.eta_seconds)
            self._eta_lbl.configure(text=f"ETA {eta}" if eta else "")
            self._draw_speed_bars(T.ACCENT)
        elif status == DownloadStatus.MERGING:
            self._speed_lbl.configure(text=_t("status_merging2"), fg=T.YELLOW)
            self._eta_lbl.configure(text="")
        elif status == DownloadStatus.FAILED:
            self._speed_lbl.configure(text=item.error_msg, fg=T.RED)
            self._eta_lbl.configure(text="")
            self._speed_canvas.delete("all")
        elif status == DownloadStatus.QUEUED:
            self._speed_lbl.configure(text="", fg=T.TEXT_DIM)
            self._eta_lbl.configure(text="")
            self._speed_canvas.delete("all")
        else:
            self._speed_lbl.configure(text="", fg=T.TEXT_DIM)
            self._eta_lbl.configure(text="")
            self._speed_canvas.delete("all")

        if item.actual_hash:
            short_h = item.actual_hash[:16] + "…"
            self._size_lbl.configure(text=short_h, fg=T.GREEN if item.status==DownloadStatus.DONE else T.TEXT_DIM)
        elif item.total_bytes > 0:
            self._size_lbl.configure(
                text=f"{_fmt_bytes(item.downloaded_bytes)} / {_fmt_bytes(item.total_bytes)}",
                fg=T.TEXT_DIM)
        elif item.downloaded_bytes > 0:
            self._size_lbl.configure(text=_fmt_bytes(item.downloaded_bytes), fg=T.TEXT_DIM)
        else:
            self._size_lbl.configure(text="")

        # Torrent seeds/peers row
        if item.torrent_seeds > 0 or item.torrent_peers > 0:
            seed_txt = f"🌱 {item.torrent_seeds} זורעים  👥 {item.torrent_peers} עמיתים"
            if item.torrent_is_magnet:
                seed_txt = "🧲 " + seed_txt
            self._torrent_lbl.configure(text=seed_txt)
            self._torrent_row.pack(fill="x", pady=(0, 2))
        else:
            self._torrent_row.pack_forget()

        self.configure(highlightbackground=T.ACCENT if status in _ACTIVE else
                       (T.YELLOW if status == DownloadStatus.QUEUED else T.BORDER))

        # Shimmer
        if status in _ACTIVE:
            self._start_shimmer()
        else:
            self._stop_shimmer()

        # Countdown
        if status == DownloadStatus.QUEUED and item.scheduled_at:
            self._start_countdown()
        else:
            self._stop_countdown()

        # Buttons
        for btn in (self._pause_btn, self._resume_btn, self._retry_btn,
                    self._open_btn, self._remove_btn, self._cancel_btn):
            btn.pack_forget()

        match status:
            case DownloadStatus.DOWNLOADING | DownloadStatus.MERGING | \
                 DownloadStatus.HASHING | DownloadStatus.EXTRACTING:
                self._pause_btn.pack(side="left")
                self._cancel_btn.pack(side="left")
            case DownloadStatus.PAUSED:
                self._resume_btn.pack(side="left")
                self._cancel_btn.pack(side="left")
            case DownloadStatus.QUEUED:
                self._cancel_btn.pack(side="left")
            case DownloadStatus.FAILED | DownloadStatus.CANCELLED:
                self._retry_btn.pack(side="left")
                self._remove_btn.pack(side="left")
            case DownloadStatus.DONE:
                self._open_btn.pack(side="left")
                self._remove_btn.pack(side="left")
            case _:
                self._cancel_btn.pack(side="left")

    def destroy(self) -> None:
        self._stop_shimmer()
        self._stop_countdown()
        super().destroy()


# ---------------------------------------------------------------------------
# Settings Dialog
# ---------------------------------------------------------------------------
class SettingsDialog(tk.Toplevel):
    def __init__(self, parent: tk.Widget, settings: Settings,
                 on_save: Callable[[Settings], None]) -> None:
        super().__init__(parent)
        self.title("⚙ הגדרות")
        self.configure(bg=T.BG_DEEP)
        self.geometry("540x640")
        self.minsize(480, 500)
        self.resizable(True, True)
        self.grab_set()
        self._s  = settings
        self._cb = on_save
        self._vars: dict[str, tk.Variable] = {}
        self._build()

    def _build(self) -> None:
        # Title
        tk.Label(self, text=_t("settings_title_lbl"), bg=T.BG_DEEP, fg=T.TEXT_MAIN,
                 font=("Consolas",13,"bold")).pack(padx=24, pady=(18,4), anchor="w")
        tk.Frame(self, bg=T.BORDER, height=1).pack(fill="x")

        # ── Footer FIRST so it's always visible at the bottom ──
        tk.Frame(self, bg=T.BORDER, height=1).pack(fill="x", side="bottom")
        footer = tk.Frame(self, bg=T.BG_DEEP, padx=24, pady=12)
        footer.pack(fill="x", side="bottom")

        def _save() -> None:
            for key, var in self._vars.items():
                setattr(self._s, key, var.get())
            if platform.system() == "Windows":
                ok, msg = _set_windows_startup(self._s.startup_with_windows)
                if not ok:
                    messagebox.showwarning(
                        "שגיאת Startup",
                        f"לא ניתן לשנות את הגדרת ההפעלה:\n{msg}",
                        parent=self
                    )
            # Apply global bandwidth cap immediately
            _BW_LIMITER.set_limit(self._s.global_bw_limit_kbps)
            self._s.save()
            self._cb(self._s)
            self.destroy()

        self._accent_btn(footer, _t("save"), _save).pack(side="right")
        self._flat_btn(footer, _t("close"), self.destroy).pack(side="right", padx=8)

        # ── Scrollable body ──
        canvas = tk.Canvas(self, bg=T.BG_DEEP, highlightthickness=0)
        vsb = tk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        body = tk.Frame(canvas, bg=T.BG_DEEP)
        body_win = canvas.create_window((0, 0), window=body, anchor="nw")

        def _on_body_configure(_event=None) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event) -> None:
            canvas.itemconfig(body_win, width=event.width)

        body.bind("<Configure>", _on_body_configure)
        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.bind("<MouseWheel>", lambda e: canvas.yview_scroll(-1 if e.delta > 0 else 1, "units"))
        canvas.bind("<Button-4>",   lambda _: canvas.yview_scroll(-1, "units"))
        canvas.bind("<Button-5>",   lambda _: canvas.yview_scroll(1,  "units"))

        # ── Helpers ──
        def _section(title: str) -> None:
            tk.Frame(body, bg=T.BORDER, height=1).pack(fill="x", padx=0, pady=(12, 0))
            tk.Label(body, text=title, bg=T.BG_DEEP, fg=T.ACCENT,
                     font=("Consolas",10,"bold")).pack(anchor="w", padx=24, pady=(6,4))

        def _check(label: str, key: str) -> None:
            v = tk.BooleanVar(value=getattr(self._s, key))
            self._vars[key] = v
            tk.Checkbutton(body, text=label, variable=v,
                           bg=T.BG_DEEP, fg=T.TEXT_MAIN, selectcolor=T.BG_CARD,
                           activebackground=T.BG_DEEP, activeforeground=T.ACCENT,
                           font=("Consolas",10), anchor="w",
                           padx=24).pack(fill="x", pady=2)

        def _spin(label: str, key: str, lo: int, hi: int) -> None:
            row = tk.Frame(body, bg=T.BG_DEEP)
            row.pack(fill="x", padx=24, pady=3)
            tk.Label(row, text=label, bg=T.BG_DEEP, fg=T.TEXT_MAIN,
                     font=("Consolas",10), width=32, anchor="w").pack(side="left")
            v = tk.IntVar(value=getattr(self._s, key))
            self._vars[key] = v
            tk.Spinbox(row, from_=lo, to=hi, textvariable=v, width=6,
                       bg=T.BG_CARD, fg=T.TEXT_MAIN, font=("Consolas",10)).pack(side="left")

        def _combo(label: str, key: str, values: list[str]) -> None:
            row = tk.Frame(body, bg=T.BG_DEEP)
            row.pack(fill="x", padx=24, pady=3)
            tk.Label(row, text=label, bg=T.BG_DEEP, fg=T.TEXT_MAIN,
                     font=("Consolas",10), width=32, anchor="w").pack(side="left")
            v = tk.StringVar(value=getattr(self._s, key))
            self._vars[key] = v
            ttk.Combobox(row, textvariable=v, values=values, width=12, state="readonly").pack(side="left")

        # ── Content ──

        # ── 🚀 Startup — at top for easy access ─────────────────────
        _section(_t("sec_startup"))
        if platform.system() == "Windows":
            _startup_actual = _is_windows_startup_enabled()
            self._s.startup_with_windows = _startup_actual
            startup_var_top = tk.BooleanVar(value=_startup_actual)

            def _toggle_startup_top() -> None:
                enable = startup_var_top.get()
                self._s.startup_with_windows = enable
                ok, msg = _set_windows_startup(enable)
                if not ok:
                    messagebox.showerror("שגיאה", msg, parent=self)
                    startup_var_top.set(not enable)
                    self._s.startup_with_windows = not enable
                else:
                    startup_status_lbl.configure(
                        text=_t("startup_enabled") if enable else _t("startup_disabled"),
                        fg=T.GREEN if enable else T.RED)

            chk_top_row = tk.Frame(body, bg=T.BG_DEEP)
            chk_top_row.pack(fill="x", padx=24, pady=(2, 0))
            tk.Checkbutton(chk_top_row, text=_t("startup_win"),
                           variable=startup_var_top, command=_toggle_startup_top,
                           bg=T.BG_DEEP, fg=T.TEXT_MAIN, selectcolor=T.BG_CARD,
                           activebackground=T.BG_DEEP, font=("Consolas", 10),
                           cursor="hand2").pack(side="left")
            startup_status_lbl = tk.Label(
                chk_top_row,
                text=_t("startup_enabled") if _startup_actual else _t("startup_disabled"),
                fg=T.GREEN if _startup_actual else T.RED,
                bg=T.BG_DEEP, font=("Consolas", 9))
            startup_status_lbl.pack(side="left", padx=8)
        else:
            tk.Label(body, text=_t("startup_win_na"),
                     bg=T.BG_DEEP, fg=T.TEXT_MUTED, font=("Consolas", 9),
                     padx=24).pack(anchor="w")

        # ── 🌍 Language — at top for easy access ─────────────────────
        _section(_t("sec_language"))
        lang_top_row = tk.Frame(body, bg=T.BG_DEEP)
        lang_top_row.pack(fill="x", padx=24, pady=4)
        tk.Label(lang_top_row, text=_t("language_label"),
                 bg=T.BG_DEEP, fg=T.TEXT_DIM, font=("Consolas", 9),
                 width=22, anchor="w").pack(side="left")
        lang_var_top = tk.StringVar(value=self._s.language)
        lang_combo_top = ttk.Combobox(
            lang_top_row, textvariable=lang_var_top,
            values=[f"{code}  —  {name}" for code, name in SUPPORTED_LANGS.items()],
            state="readonly", width=26)
        lang_combo_top.pack(side="left")
        for i, (code, _) in enumerate(SUPPORTED_LANGS.items()):
            if code == self._s.language:
                lang_combo_top.current(i)
                break

        def _on_lang_top(*_) -> None:
            val = lang_var_top.get().split("  ")[0].strip()
            if val in SUPPORTED_LANGS:
                global _LANG
                _LANG = val
                self._s.language = val

        lang_combo_top.bind("<<ComboboxSelected>>", _on_lang_top)
        tk.Label(body, text=_t("lang_restart_warn"),
                 bg=T.BG_DEEP, fg=T.YELLOW, font=("Consolas", 8),
                 padx=24).pack(anchor="w", pady=(2, 0))

        # ── Downloads ─────────────────────────────────────────────────
        _section("🔽 הורדות")
        _spin("מספר הורדות מקביליות:", "max_concurrent", 1, 16)
        _check("הורדה מקוטעת (מהירות גבוהה לקבצים גדולים)", "multipart")
        _check("נסה שוב אוטומטית בכשל רשת", "auto_retry")
        _spin("מספר ניסיונות חוזרים:", "max_retries", 1, 10)
        _spin("הגבלת מהירות כוללת (KB/s, 0=ללא):", "global_bw_limit_kbps", 0, 100000)

        _section("📁 קבצים")
        _check("סיווג אוטומטי לתיקיות לפי סוג", "auto_categorize")
        _check("חילוץ אוטומטי של ארכיונים", "auto_extract")
        _check("פתיחה אוטומטית אחרי הורדה", "auto_open")
        _check("אימות Hash אוטומטי", "verify_hash")
        _combo("אלגוריתם Hash:", "default_hash_algo", ["md5","sha256","sha1"])

        _section("🎨 ממשק")
        _check("נושא כהה", "dark_theme")
        _check("System Tray (X מסתיר במגש)", "system_tray")
        _check("התראות שולחן עבודה", "notify_done")
        _check("צליל בסיום הורדה", "sound_on_done")
        _check("אשר לפני מחיקת קובץ חלקי", "confirm_delete_file")
        _check("ניטור לוח גזירה (הוסף URL אוטומטי)", "clipboard_monitor")
        _check("אזהרה על כפילות URL", "warn_duplicates")
        _spin("גודל קובץ מקסימלי (MB, 0=ללא):", "max_file_size_mb", 0, 100000)

        _section("🧲 טורנטים") if LIBTORRENT_OK else "✗ לא מותקן — הרץ: pip install libtorrent"
        torrent_color  = T.GREEN if LIBTORRENT_OK else T.RED
        tk.Label(body, text=torrent_status, bg=T.BG_DEEP, fg=torrent_color,
                 font=("Consolas",9), padx=24).pack(anchor="w", pady=(0,4))
        _check("המשך זריעה לאחר הורדה", "torrent_seeding")
        _spin("יחס זריעה (עצור אחרי):", "torrent_seed_ratio", 0, 100)
        _spin("הגבל הורדה (KB/s, 0=ללא):", "torrent_max_dl_kbps", 0, 100000)
        _spin("הגבל העלאה (KB/s, 0=ללא):", "torrent_max_ul_kbps", 0, 100000)

        # ── VirusTotal ─────────────────────────────────────────────────
        _section("🛡 VirusTotal — סריקת וירוסים")
        _check("הפעל סריקת VirusTotal לאחר הורדה", "virustotal_enabled")
        tk.Label(body,
                 text=_t("vt_free_key"),
                 bg=T.BG_DEEP, fg=T.TEXT_MUTED, font=("Consolas",8), padx=24
                 ).pack(anchor="w", pady=(0,4))
        vt_row = tk.Frame(body, bg=T.BG_DEEP)
        vt_row.pack(fill="x", padx=24, pady=2)
        tk.Label(vt_row, text="API Key:", bg=T.BG_DEEP, fg=T.TEXT_DIM,
                 font=("Consolas",9), width=14, anchor="w").pack(side="left")
        vt_var = tk.StringVar(value=self._s.virustotal_api_key)
        vt_entry = tk.Entry(vt_row, textvariable=vt_var, bg=T.BG_CARD, fg=T.TEXT_MAIN,
                            insertbackground=T.ACCENT, font=("Consolas",9),
                            relief="flat", width=44, show="*")
        vt_entry.pack(side="left")
        vt_var.trace_add("write",
                         lambda *_: setattr(self._s, "virustotal_api_key", vt_var.get()))
        show_vt = tk.BooleanVar(value=False)
        def _toggle_vt():
            vt_entry.configure(show="" if show_vt.get() else "*")
        tk.Checkbutton(vt_row, text=_t("show_key"), variable=show_vt, command=_toggle_vt,
                       bg=T.BG_DEEP, fg=T.TEXT_DIM, selectcolor=T.BG_CARD,
                       activebackground=T.BG_DEEP, font=("Consolas",8),
                       cursor="hand2").pack(side="left", padx=(6,0))
        def _test_vt():
            key = vt_var.get().strip()
            if not key:
                messagebox.showwarning("API Key חסר", "הכנס API key תחילה", parent=self)
                return
            try:
                r = _get_session().get(
                    "https://www.virustotal.com/api/v3/users/me",
                    headers={"x-apikey": key}, timeout=(8, 15))
                if r.status_code == 200:
                    qa = r.json().get("data",{}).get("attributes",{}).get("quotas",{})
                    d  = qa.get("api_requests_daily", {})
                    messagebox.showinfo("✓ VirusTotal",
                        f"API Key תקין!\nשימוש יומי: {d.get('used','?')}/{d.get('allowed','?')}",
                        parent=self)
                else:
                    messagebox.showerror("✗ VirusTotal",
                        f"API Key לא תקין (קוד {r.status_code})", parent=self)
            except Exception as e:
                messagebox.showerror("שגיאה", str(e), parent=self)
        test_btn = tk.Label(body, text=_t("vt_test_btn"), cursor="hand2",
                            bg=T.BG_DEEP, fg=T.ACCENT, font=("Consolas",9),
                            padx=10, pady=3,
                            highlightthickness=1, highlightbackground=T.BORDER)
        test_btn.bind("<Button-1>", lambda _: _test_vt())
        test_btn.bind("<Enter>",    lambda _, w=test_btn: w.configure(bg=T.BG_HOVER))
        test_btn.bind("<Leave>",    lambda _, w=test_btn: w.configure(bg=T.BG_DEEP))
        test_btn.pack(anchor="w", padx=24, pady=(4,0))

        # ── Proxy / SOCKS5 ─────────────────────────────────────────────
        _section("🌐 Proxy / SOCKS5")
        _check("הפעל Proxy", "proxy_enabled")
        proxy_type_row = tk.Frame(body, bg=T.BG_DEEP)
        proxy_type_row.pack(fill="x", padx=24, pady=(0,4))
        tk.Label(proxy_type_row, text=_t("proxy_type"), bg=T.BG_DEEP, fg=T.TEXT_DIM,
                 font=("Consolas",9), width=14, anchor="w").pack(side="left")
        proxy_type_var = tk.StringVar(value=self._s.proxy_type)
        for val, lbl in [("http","HTTP"), ("socks5","SOCKS5")]:
            tk.Radiobutton(proxy_type_row, text=lbl, variable=proxy_type_var, value=val,
                           bg=T.BG_DEEP, fg=T.TEXT_MAIN, selectcolor=T.BG_CARD,
                           activebackground=T.BG_DEEP, font=("Consolas",9),
                           command=lambda: setattr(self._s, "proxy_type", proxy_type_var.get())
                           ).pack(side="left", padx=(0,10))

        def _proxy_entry(label: str, attr: str, width: int = 28) -> None:
            row = tk.Frame(body, bg=T.BG_DEEP)
            row.pack(fill="x", padx=24, pady=2)
            tk.Label(row, text=label, bg=T.BG_DEEP, fg=T.TEXT_DIM,
                     font=("Consolas",9), width=14, anchor="w").pack(side="left")
            var = tk.StringVar(value=str(getattr(self._s, attr, "")))
            e = tk.Entry(row, textvariable=var, bg=T.BG_CARD, fg=T.TEXT_MAIN,
                         insertbackground=T.ACCENT, font=("Consolas",9),
                         relief="flat", width=width)
            e.pack(side="left")
            def _update(*_, a=attr, v=var):
                val = v.get()
                try:
                    setattr(self._s, a, int(val) if a == "proxy_port" else val)
                except Exception:
                    pass
            var.trace_add("write", _update)

        _proxy_entry("Host:", "proxy_host")
        _proxy_entry("Port:", "proxy_port", 8)
        _proxy_entry("משתמש:", "proxy_user")
        _proxy_entry("סיסמה:", "proxy_pass")

        # ── YouTube / Media ────────────────────────────────────────────
        _section("🎬 YouTube / מדיה")
        ytdlp_status = "✓ yt-dlp מותקן" if YTDLP_OK else "✗ לא מותקן — הרץ: pip install yt-dlp"
        tk.Label(body, text=ytdlp_status, bg=T.BG_DEEP,
                 fg=T.GREEN if YTDLP_OK else T.RED,
                 font=("Consolas",9), padx=24).pack(anchor="w", pady=(0,4))

        fmt_row = tk.Frame(body, bg=T.BG_DEEP)
        fmt_row.pack(fill="x", padx=24, pady=(0,4))
        tk.Label(fmt_row, text=_t("media_default_fmt"), bg=T.BG_DEEP, fg=T.TEXT_DIM,
                 font=("Consolas",9), width=22, anchor="w").pack(side="left")
        fmt_var = tk.StringVar(value=self._s.ytdlp_default_format or "best")
        fmt_combo = ttk.Combobox(fmt_row, textvariable=fmt_var, width=12,
                                 values=["best","4k","1080p","720p","480p","360p","mp3","m4a","opus"],
                                 state="readonly")
        fmt_combo.pack(side="left")
        fmt_var.trace_add("write", lambda *_: setattr(self._s, "ytdlp_default_format", fmt_var.get()))

        _check("ברירת מחדל — אודיו בלבד", "ytdlp_prefer_audio")
        _check(_t("embed_thumb_audio"), "ytdlp_embed_thumbnail")
        _check(_t("add_metadata_audio"), "ytdlp_add_metadata")

        # ── Automation ─────────────────────────────────────────────────
        _section("⚙ אוטומציה")
        _check("שמור/שחזר תור בין הפעלות", "persistent_queue")
        _check("כבה/שנה מצב שינה אחרי השלמת התור", "shutdown_after_queue")
        shutdown_row = tk.Frame(body, bg=T.BG_DEEP)
        shutdown_row.pack(fill="x", padx=24, pady=(0,4))
        tk.Label(shutdown_row, text=_t("shutdown_action"), bg=T.BG_DEEP, fg=T.TEXT_DIM,
                 font=("Consolas",9), width=14, anchor="w").pack(side="left")
        sd_var = tk.StringVar(value=self._s.shutdown_action)
        for val, lbl in [("sleep","שנת מחשב"),("hibernate","שינה עמוקה"),("shutdown","כיבוי")]:
            tk.Radiobutton(shutdown_row, text=lbl, variable=sd_var, value=val,
                           bg=T.BG_DEEP, fg=T.TEXT_MAIN, selectcolor=T.BG_CARD,
                           activebackground=T.BG_DEEP, font=("Consolas",9),
                           command=lambda: setattr(self._s, "shutdown_action", sd_var.get())
                           ).pack(side="left", padx=(0,10))

        _check("בדוק מקום דיסק לפני הורדה", "disk_check_enabled")
        _spin("Watchdog — אתחל הורדה תקועה אחרי (דקות):", "watchdog_timeout_min", 1, 60)

        # ── Schedule range ─────────────────────────────────────────────
        _section("📅 תזמון טווח שעות")
        _check("הפעל הורדה רק בטווח שעות מוגדר", "schedule_range_enabled")
        sched_row = tk.Frame(body, bg=T.BG_DEEP)
        sched_row.pack(fill="x", padx=24, pady=(0,4))
        tk.Label(sched_row, text=_t("sched_from"), bg=T.BG_DEEP, fg=T.TEXT_DIM,
                 font=("Consolas",9)).pack(side="left")
        sr_start = tk.StringVar(value=self._s.schedule_range_start)
        tk.Entry(sched_row, textvariable=sr_start, width=6,
                 bg=T.BG_CARD, fg=T.TEXT_MAIN, font=("Consolas",9),
                 relief="flat").pack(side="left", padx=(4,12))
        tk.Label(sched_row, text=_t("sched_to"), bg=T.BG_DEEP, fg=T.TEXT_DIM,
                 font=("Consolas",9)).pack(side="left")
        sr_end = tk.StringVar(value=self._s.schedule_range_end)
        tk.Entry(sched_row, textvariable=sr_end, width=6,
                 bg=T.BG_CARD, fg=T.TEXT_MAIN, font=("Consolas",9),
                 relief="flat").pack(side="left", padx=4)
        tk.Label(sched_row, text="(HH:MM)", bg=T.BG_DEEP, fg=T.TEXT_MUTED,
                 font=("Consolas",8)).pack(side="left", padx=4)
        def _save_sched_range(*_):
            self._s.schedule_range_start = sr_start.get()
            self._s.schedule_range_end   = sr_end.get()
        sr_start.trace_add("write", _save_sched_range)
        sr_end.trace_add("write",   _save_sched_range)

        # ── Startup with Windows ────────────────────────────────────────
        _section(_t("sec_startup"))
        if platform.system() == "Windows":
            # Refresh actual state
            self._s.startup_with_windows = _is_windows_startup_enabled()
            startup_var = tk.BooleanVar(value=self._s.startup_with_windows)

            def _toggle_startup() -> None:
                enable = startup_var.get()
                self._s.startup_with_windows = enable
                ok, msg = _set_windows_startup(enable)
                if not ok:
                    messagebox.showerror("שגיאה", msg, parent=self)
                    startup_var.set(not enable)      # revert
                    self._s.startup_with_windows = not enable
                else:
                    status_lbl.configure(
                        text=_t("startup_enabled") if enable else _t("startup_disabled"),
                        fg=T.GREEN if enable else T.RED)

            chk_row = tk.Frame(body, bg=T.BG_DEEP)
            chk_row.pack(fill="x", padx=24, pady=2)
            tk.Checkbutton(chk_row, text=_t("startup_win"),
                           variable=startup_var, command=_toggle_startup,
                           bg=T.BG_DEEP, fg=T.TEXT_MAIN, selectcolor=T.BG_CARD,
                           activebackground=T.BG_DEEP, font=("Consolas", 10),
                           cursor="hand2").pack(side="left")
            status_lbl = tk.Label(chk_row,
                text=_t("startup_enabled") if self._s.startup_with_windows else _t("startup_disabled"),
                fg=T.GREEN if self._s.startup_with_windows else T.RED,
                bg=T.BG_DEEP, font=("Consolas", 9))
            status_lbl.pack(side="left", padx=8)

            tk.Label(body,
                     text="📁 " + str(APP_DIR / "FetchPro.lnk"),
                     bg=T.BG_DEEP, fg=T.TEXT_MUTED, font=("Consolas", 7),
                     padx=32).pack(anchor="w")
        else:
            tk.Label(body, text=_t("startup_win_na"),
                     bg=T.BG_DEEP, fg=T.TEXT_MUTED, font=("Consolas", 9),
                     padx=24).pack(anchor="w")

        # ── Language ────────────────────────────────────────────────────
        _section(_t("sec_language"))
        lang_row = tk.Frame(body, bg=T.BG_DEEP)
        lang_row.pack(fill="x", padx=24, pady=4)
        tk.Label(lang_row, text=_t("language_label"),
                 bg=T.BG_DEEP, fg=T.TEXT_DIM, font=("Consolas", 9),
                 width=22, anchor="w").pack(side="left")
        lang_var = tk.StringVar(value=self._s.language)
        lang_combo = ttk.Combobox(
            lang_row, textvariable=lang_var,
            values=[f"{code}  —  {name}" for code, name in SUPPORTED_LANGS.items()],
            state="readonly", width=28)
        lang_combo.pack(side="left")
        # Match current lang to combo value
        for i, (code, name) in enumerate(SUPPORTED_LANGS.items()):
            if code == self._s.language:
                lang_combo.current(i)
                break

        def _on_lang_change(*_) -> None:
            val = lang_var.get().split("  ")[0].strip()
            if val in SUPPORTED_LANGS:
                global _LANG
                _LANG = val
                self._s.language = val

        lang_combo.bind("<<ComboboxSelected>>", _on_lang_change)

        tk.Label(body,
                 text=_t("lang_restart_warn"),
                 bg=T.BG_DEEP, fg=T.YELLOW, font=("Consolas", 8), padx=24
                 ).pack(anchor="w", pady=(2, 0))

        # bottom padding
        tk.Frame(body, bg=T.BG_DEEP, height=16).pack()

    def _accent_btn(self, parent: tk.Widget, text: str, cmd: Callable) -> tk.Label:
        b = tk.Label(parent, text=text, bg="#003d4d", fg=T.ACCENT,
                     font=("Consolas",11,"bold"), padx=16, pady=6, cursor="hand2",
                     highlightthickness=1, highlightbackground=T.ACCENT)
        b.bind("<Button-1>", lambda _: cmd())
        return b

    def _flat_btn(self, parent: tk.Widget, text: str, cmd: Callable) -> tk.Label:
        b = tk.Label(parent, text=text, bg=T.BG_DEEP, fg=T.TEXT_DIM,
                     font=("Consolas",10), padx=10, pady=4, cursor="hand2",
                     highlightthickness=1, highlightbackground=T.BORDER)
        b.bind("<Button-1>", lambda _: cmd())
        return b


# ---------------------------------------------------------------------------
# History Panel
# ---------------------------------------------------------------------------
class HistoryPanel(tk.Toplevel):
    def __init__(self, parent: tk.Widget, db: HistoryDB) -> None:
        super().__init__(parent)
        self.title("📜 היסטוריית הורדות")
        self.configure(bg=T.BG_DEEP)
        self.geometry("820x560")
        self.minsize(600, 400)
        self.grab_set()
        self._db   = db
        self._rows: list[dict] = []
        self._build()

    def _build(self) -> None:
        hdr = tk.Frame(self, bg=T.BG_DEEP, padx=20, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text=_t("history_title"), bg=T.BG_DEEP, fg=T.TEXT_MAIN,
                 font=("Consolas",13,"bold")).pack(side="left")

        # Search bar
        search_frame = tk.Frame(hdr, bg=T.BG_DEEP)
        search_frame.pack(side="right")
        tk.Label(search_frame, text="🔍", bg=T.BG_DEEP, fg=T.TEXT_DIM,
                 font=("Consolas",10)).pack(side="left")
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._do_search())
        search_entry = tk.Entry(search_frame, textvariable=self._search_var,
                                bg=T.BG_CARD, fg=T.TEXT_MAIN,
                                insertbackground=T.ACCENT, font=("Consolas",10),
                                relief="flat", width=24)
        search_entry.pack(side="left", padx=6)

        tk.Frame(self, bg=T.BORDER, height=1).pack(fill="x")

        # Treeview
        cols = ("filename","status","size","finished")
        self._tree = ttk.Treeview(self, columns=cols, show="headings", height=18)
        heads  = {"filename":"שם קובץ","status":"סטטוס","size":"גודל","finished":"סיום"}
        widths = {"filename":310,"status":80,"size":90,"finished":130}
        for c in cols:
            self._tree.heading(c, text=heads[c],
                               command=lambda col=c: self._sort_by(col))
            self._tree.column(c, width=widths[c], anchor="w")

        scroll = ttk.Scrollbar(self, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self._tree.pack(fill="both", expand=True, padx=16, pady=8)

        self._tree.bind("<Double-1>",  self._copy_url)
        self._tree.bind("<Button-3>",  self._ctx_menu)
        self._tree.bind("<Delete>",    lambda _: self._delete_selected())

        tk.Frame(self, bg=T.BORDER, height=1).pack(fill="x")
        footer = tk.Frame(self, bg=T.BG_DEEP, padx=16, pady=10)
        footer.pack(fill="x")

        self._count_lbl = tk.Label(footer, text="", bg=T.BG_DEEP, fg=T.TEXT_MUTED,
                                   font=("Consolas",8))
        self._count_lbl.pack(side="left")

        def _clear() -> None:
            if messagebox.askyesno("נקה היסטוריה", "למחוק את כל ההיסטוריה?", parent=self):
                self._db.clear()
                self._refresh()

        def _export_csv() -> None:
            import csv as _csv
            path = filedialog.asksaveasfilename(
                title=_t("csv_export"), defaultextension=".csv",
                filetypes=[("CSV","*.csv"),("כל הקבצים","*.*")], parent=self)
            if not path:
                return
            rows = self._db.fetch(limit=100000)
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = _csv.DictWriter(f, fieldnames=["id","filename","url","save_dir","status","size","hash","finished"])
                writer.writeheader()
                writer.writerows(rows)
            self._count_lbl.configure(text=f"ייוצאו {len(rows)} רשומות ✓", fg=T.GREEN)

        for text, color, cmd in [("רענן", T.ACCENT, self._refresh),
                                  ("📊 CSV", T.TEXT_DIM, _export_csv),
                                  ("מחק נבחרים", T.YELLOW, self._delete_selected),
                                  (_t("clear_all"), T.RED, _clear),
                                  (_t("close"), T.TEXT_DIM, self.destroy)]:
            btn = tk.Label(footer, text=text, bg=T.BG_DEEP, fg=color,
                           font=("Consolas",10), padx=10, pady=4, cursor="hand2",
                           highlightthickness=1, highlightbackground=T.BORDER)
            btn.bind("<Button-1>", lambda _, c=cmd: c())
            btn.pack(side="right", padx=4)

        self._sort_col: str = "finished"
        self._sort_rev: bool = True
        self._refresh()

    def _refresh(self) -> None:
        q = self._search_var.get().strip() if hasattr(self, "_search_var") else ""
        self._rows = self._db.search(q) if q else self._db.fetch()
        self._render()

    def _do_search(self) -> None:
        self._refresh()

    def _sort_by(self, col: str) -> None:
        if self._sort_col == col:
            self._sort_rev = not self._sort_rev
        else:
            self._sort_col = col
            self._sort_rev = True
        self._render()

    def _render(self) -> None:
        col = self._sort_col
        rev = self._sort_rev
        rows = sorted(self._rows, key=lambda r: r.get(col,""), reverse=rev)
        self._tree.delete(*self._tree.get_children())
        for r in rows:
            tag = "done" if r["status"] == "DONE" else ("fail" if r["status"] == "FAILED" else "")
            self._tree.insert("", "end", iid=str(r["id"]),
                              values=(r["filename"], r["status"],
                                      _fmt_bytes(r["size"]), r["finished"]),
                              tags=(tag,))
        self._tree.tag_configure("done", foreground=T.GREEN)
        self._tree.tag_configure("fail", foreground=T.RED)
        self._count_lbl.configure(text=f"{len(rows)} רשומות")

    def _copy_url(self, _event: tk.Event) -> None:
        sel = self._tree.selection()
        if not sel:
            return
        row = next((r for r in self._rows if str(r["id"]) == sel[0]), None)
        if row:
            self.clipboard_clear()
            self.clipboard_append(row["url"])
            # flash feedback
            self._count_lbl.configure(text=_t("verified_copied"), fg=T.GREEN)
            self.after(1500, lambda: self._count_lbl.configure(
                text=f"{len(self._rows)} רשומות", fg=T.TEXT_MUTED))

    def _delete_selected(self) -> None:
        sels = self._tree.selection()
        if not sels:
            return
        if not messagebox.askyesno("מחיקה", f"למחוק {len(sels)} רשומות?", parent=self):
            return
        for iid in sels:
            self._db.delete_by_id(int(iid))
        self._refresh()

    def _ctx_menu(self, event: tk.Event) -> None:
        iid = self._tree.identify_row(event.y)
        if not iid:
            return
        self._tree.selection_set(iid)
        row = next((r for r in self._rows if str(r["id"]) == iid), None)
        if not row:
            return
        menu = tk.Menu(self, tearoff=0, bg=T.BG_CARD, fg=T.TEXT_MAIN,
                       activebackground=T.ACCENT, activeforeground=T.BG_DEEP,
                       font=("Consolas",9))
        menu.add_command(label=_t("ctx_copy_url"),
                         command=lambda: (self.clipboard_clear(), self.clipboard_append(row["url"])))
        menu.add_command(label=_t("ctx_open_folder2"),
                         command=lambda: _open_file(Path(row["save_dir"])))
        fp = Path(row["save_dir"]) / row["filename"]
        if fp.exists():
            menu.add_command(label=_t("ctx_open_file"),
                             command=lambda: _open_file(fp))
        menu.add_separator()
        menu.add_command(label=_t("ctx_delete_record"), command=self._delete_selected)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()


# ---------------------------------------------------------------------------
# Add Download Dialog — advanced options
# ---------------------------------------------------------------------------
class AddDownloadDialog(tk.Toplevel):
    """Dialog for adding a download with advanced options."""

    def __init__(self, parent: tk.Widget, url: str, settings: Settings,
                 on_add: Callable[[DownloadItem], None]) -> None:
        super().__init__(parent)
        self.title("➕ הוסף הורדה")
        self.configure(bg=T.BG_DEEP)
        self.geometry("560x480")
        self.resizable(False, False)
        self.grab_set()
        self._url  = url
        self._s    = settings
        self._cb   = on_add
        self._build()

    def _build(self) -> None:
        tk.Label(self, text=_t("adv_settings"), bg=T.BG_DEEP, fg=T.TEXT_MAIN,
                 font=("Consolas",12,"bold")).pack(padx=24, pady=(18,8), anchor="w")
        tk.Frame(self, bg=T.BORDER, height=1).pack(fill="x")

        body = tk.Frame(self, bg=T.BG_DEEP, padx=24, pady=16)
        body.pack(fill="both", expand=True)

        def _row(label: str) -> tk.Frame:
            r = tk.Frame(body, bg=T.BG_DEEP)
            r.pack(fill="x", pady=5)
            tk.Label(r, text=label, bg=T.BG_DEEP, fg=T.TEXT_DIM,
                     font=("Consolas",9), width=20, anchor="w").pack(side="left")
            return r

        # URL (read-only)
        r = _row("URL:")
        tk.Label(r, text=(self._url[:55]+"…") if len(self._url)>55 else self._url,
                 bg=T.BG_DEEP, fg=T.ACCENT, font=("Consolas",9)).pack(side="left")

        # Filename
        r = _row("שם קובץ:")
        self._fname_var = tk.StringVar(value=_derive_filename(self._url))
        tk.Entry(r, textvariable=self._fname_var, bg=T.BG_CARD, fg=T.TEXT_MAIN,
                 insertbackground=T.ACCENT, font=("Consolas",10), relief="flat",
                 width=32).pack(side="left")

        # Hash
        r = _row("Hash (אופציונלי):")
        self._hash_var = tk.StringVar()
        tk.Entry(r, textvariable=self._hash_var, bg=T.BG_CARD, fg=T.TEXT_MAIN,
                 insertbackground=T.ACCENT, font=("Consolas",10), relief="flat",
                 width=32).pack(side="left")

        # Throttle
        r = _row("מגבלת מהירות (KB/s):")
        self._throttle_var = tk.IntVar(value=0)
        tk.Spinbox(r, from_=0, to=100000, textvariable=self._throttle_var,
                   width=10, bg=T.BG_CARD, fg=T.TEXT_MAIN,
                   font=("Consolas",10)).pack(side="left")
        tk.Label(r, text=_t("zero_no_limit"), bg=T.BG_DEEP, fg=T.TEXT_MUTED,
                 font=("Consolas",8)).pack(side="left", padx=6)

        # Schedule — proper date+time pickers
        r = _row("תזמון הורדה:")
        now = datetime.now()
        self._sched_enabled = tk.BooleanVar(value=False)
        chk = tk.Checkbutton(r, text=_t("schedule_btn"), variable=self._sched_enabled,
                             bg=T.BG_DEEP, fg=T.TEXT_MAIN, selectcolor=T.BG_CARD,
                             activebackground=T.BG_DEEP, font=("Consolas",9),
                             command=self._toggle_schedule)
        chk.pack(side="left")

        self._sched_frame = tk.Frame(r, bg=T.BG_DEEP)
        self._sched_frame.pack(side="left", padx=6)

        spin_cfg = dict(bg=T.BG_CARD, fg=T.TEXT_MAIN, font=("Consolas",10),
                        relief="flat", width=4, justify="center")
        # Day / Month / Year
        self._sched_day   = tk.Spinbox(self._sched_frame, from_=1,  to=31,  **{**spin_cfg, "width":3})
        self._sched_month = tk.Spinbox(self._sched_frame, from_=1,  to=12,  **{**spin_cfg, "width":3})
        self._sched_year  = tk.Spinbox(self._sched_frame, from_=now.year, to=now.year+5, **spin_cfg)
        self._sched_hour  = tk.Spinbox(self._sched_frame, from_=0,  to=23,  **{**spin_cfg, "width":3})
        self._sched_min   = tk.Spinbox(self._sched_frame, from_=0,  to=59,  **{**spin_cfg, "width":3})

        # Preset to now+1h
        preset = now + timedelta(hours=1)
        self._sched_day.delete(0, "end");   self._sched_day.insert(0, str(preset.day))
        self._sched_month.delete(0, "end"); self._sched_month.insert(0, str(preset.month))
        self._sched_year.delete(0, "end");  self._sched_year.insert(0, str(preset.year))
        self._sched_hour.delete(0, "end");  self._sched_hour.insert(0, f"{preset.hour:02d}")
        self._sched_min.delete(0, "end");   self._sched_min.insert(0, f"{preset.minute:02d}")

        for w, sep in [(self._sched_day, "/"), (self._sched_month, "/"),
                       (self._sched_year, " "), (self._sched_hour, ":"),
                       (self._sched_min, None)]:
            w.pack(side="left")
            if sep:
                tk.Label(self._sched_frame, text=sep, bg=T.BG_DEEP, fg=T.TEXT_DIM,
                         font=("Consolas",10)).pack(side="left")

        self._toggle_schedule()  # start hidden

        # Note
        r = _row("הערה:")
        self._note_var = tk.StringVar()
        tk.Entry(r, textvariable=self._note_var, bg=T.BG_CARD, fg=T.TEXT_MAIN,
                 insertbackground=T.ACCENT, font=("Consolas",10), relief="flat",
                 width=32).pack(side="left")

        # Priority
        r = _row("עדיפות:")
        self._priority_var = tk.StringVar(value="NORMAL")
        for val, label, color in [("HIGH","⚡ גבוהה","#ff4444"),
                                   ("NORMAL","● רגילה", T.ACCENT),
                                   ("LOW","⬇ נמוכה","#888888")]:
            tk.Radiobutton(r, text=label, variable=self._priority_var, value=val,
                           bg=T.BG_DEEP, fg=color, selectcolor=T.BG_CARD,
                           activebackground=T.BG_DEEP, font=("Consolas",9)).pack(side="left", padx=4)

        # Checkboxes
        self._multipart_var  = tk.BooleanVar(value=self._s.multipart)
        self._extract_var    = tk.BooleanVar(value=self._s.auto_extract)
        self._open_var       = tk.BooleanVar(value=self._s.auto_open)
        for text, var in [("הורדה מקוטעת", self._multipart_var),
                          ("חלץ אוטומטית", self._extract_var),
                          ("פתח אחרי הורדה", self._open_var)]:
            tk.Checkbutton(body, text=text, variable=var,
                           bg=T.BG_DEEP, fg=T.TEXT_MAIN, selectcolor=T.BG_CARD,
                           activebackground=T.BG_DEEP, font=("Consolas",10),
                           anchor="w").pack(fill="x", pady=2)

        tk.Frame(self, bg=T.BORDER, height=1).pack(fill="x")
        footer = tk.Frame(self, bg=T.BG_DEEP, padx=24, pady=12)
        footer.pack(fill="x")

        def _submit() -> None:
            scheduled = None
            if self._sched_enabled.get():
                try:
                    d = int(self._sched_day.get())
                    mo = int(self._sched_month.get())
                    y = int(self._sched_year.get())
                    h = int(self._sched_hour.get())
                    mi = int(self._sched_min.get())
                    scheduled = datetime(y, mo, d, h, mi)
                    if scheduled <= datetime.now():
                        messagebox.showerror("שגיאה", "זמן התזמון חייב להיות בעתיד", parent=self)
                        return
                except (ValueError, tk.TclError):
                    messagebox.showerror("שגיאה", "תאריך/שעה לא תקינים", parent=self)
                    return
            fname    = _sanitize_filename(self._fname_var.get().strip() or _derive_filename(self._url))
            throttle = self._throttle_var.get() * 1024
            item = DownloadItem(
                url=self._url,
                save_dir=Path(self._s.save_dir),
                filename=fname,
                hash_algo=self._s.default_hash_algo if self._s.verify_hash else "",
                expected_hash=self._hash_var.get().strip(),
                throttle_bps=throttle,
                multipart=self._multipart_var.get(),
                auto_extract=self._extract_var.get(),
                auto_open=self._open_var.get(),
                note=self._note_var.get().strip(),
                scheduled_at=scheduled,
                dl_priority=self._priority_var.get(),
            )
            self.destroy()
            self._cb(item)

        self._accent_btn(footer, _t("add_to_download"), _submit).pack(side="right")
        self._flat_btn(footer, _t("cancel_btn"), self.destroy).pack(side="right", padx=8)

    def _toggle_schedule(self) -> None:
        """Show or hide the date/time spinboxes based on the checkbox."""
        if self._sched_enabled.get():
            for w in self._sched_frame.winfo_children():
                w.configure(state="normal")
            self._sched_frame.pack(side="left", padx=6)
        else:
            for w in self._sched_frame.winfo_children():
                try:
                    w.configure(state="disabled")
                except tk.TclError:
                    pass

    def _accent_btn(self, parent: tk.Widget, text: str, cmd: Callable) -> tk.Label:
        b = tk.Label(parent, text=text, bg="#003d4d", fg=T.ACCENT,
                     font=("Consolas",11,"bold"), padx=16, pady=6, cursor="hand2",
                     highlightthickness=1, highlightbackground=T.ACCENT)
        b.bind("<Button-1>", lambda _: cmd())
        return b

    def _flat_btn(self, parent: tk.Widget, text: str, cmd: Callable) -> tk.Label:
        b = tk.Label(parent, text=text, bg=T.BG_DEEP, fg=T.TEXT_DIM,
                     font=("Consolas",10), padx=10, pady=4, cursor="hand2",
                     highlightthickness=1, highlightbackground=T.BORDER)
        b.bind("<Button-1>", lambda _: cmd())
        return b


# ---------------------------------------------------------------------------
# Persistent Queue  — save/restore download queue between sessions
# ---------------------------------------------------------------------------
class PersistentQueue:
    """Saves the download queue to disk so it survives app restarts."""

    _SAVEABLE_FIELDS = (
        "url", "filename", "added_at", "dl_priority", "hash_algo",
        "expected_hash", "note", "media_format", "media_is_audio",
        "media_playlist", "tags", "throttle_bps", "multipart",
        "auto_extract", "auto_open",
    )

    def save(self, items: list["DownloadItem"]) -> None:
        try:
            data = []
            for item in items:
                if item.status in _TERMINAL:
                    continue          # don't restore completed/failed items
                row: dict = {
                    "save_dir": str(item.save_dir),
                    "status":   item.status.name,
                }
                for f in self._SAVEABLE_FIELDS:
                    row[f] = getattr(item, f, "")
                if item.scheduled_at:
                    row["scheduled_at"] = item.scheduled_at.isoformat()
                data.append(row)
            QUEUE_FILE.parent.mkdir(parents=True, exist_ok=True)
            QUEUE_FILE.write_text(json.dumps(data, indent=2, ensure_ascii=False))
        except Exception as e:
            logger.warning("PersistentQueue.save failed: %s", e)

    def load(self) -> list["DownloadItem"]:
        items: list[DownloadItem] = []
        try:
            if not QUEUE_FILE.exists():
                return items
            data = json.loads(QUEUE_FILE.read_text())
            for row in data:
                save_dir = Path(row.get("save_dir", str(Path.home() / "Downloads")))
                item = DownloadItem(
                    url      = row.get("url", ""),
                    save_dir = save_dir,
                )
                for f in self._SAVEABLE_FIELDS:
                    if f in row:
                        try:
                            setattr(item, f, row[f])
                        except Exception:
                            pass
                if "scheduled_at" in row and row["scheduled_at"]:
                    try:
                        item.scheduled_at = datetime.fromisoformat(row["scheduled_at"])
                    except Exception:
                        pass
                # Restore to PENDING so it gets re-queued
                item.status = DownloadStatus.PENDING
                items.append(item)
        except Exception as e:
            logger.warning("PersistentQueue.load failed: %s", e)
        return items


# ---------------------------------------------------------------------------
# Stats Tracker  — usage statistics
# ---------------------------------------------------------------------------
class StatsTracker:
    """Tracks cumulative download stats across sessions."""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._data = self._load()

    def _load(self) -> dict:
        try:
            if STATS_FILE.exists():
                return json.loads(STATS_FILE.read_text())
        except Exception:
            pass
        return {
            "total_bytes":    0,
            "total_files":    0,
            "total_sessions": 0,
            "fastest_bps":    0,
            "session_bytes":  0,
            "session_files":  0,
        }

    def save(self) -> None:
        try:
            STATS_FILE.parent.mkdir(parents=True, exist_ok=True)
            STATS_FILE.write_text(json.dumps(self._data, indent=2))
        except Exception:
            pass

    def record_done(self, item: "DownloadItem") -> None:
        with self._lock:
            size = item.downloaded_bytes or item.total_bytes
            self._data["total_bytes"]   += size
            self._data["total_files"]   += 1
            self._data["session_bytes"] += size
            self._data["session_files"] += 1
            if item.speed_bps > self._data.get("fastest_bps", 0):
                self._data["fastest_bps"] = item.speed_bps
        self.save()

    def new_session(self) -> None:
        with self._lock:
            self._data["total_sessions"] = self._data.get("total_sessions", 0) + 1
            self._data["session_bytes"] = 0
            self._data["session_files"] = 0
        self.save()

    @property
    def data(self) -> dict:
        with self._lock:
            return dict(self._data)


_STATS = StatsTracker()


# ---------------------------------------------------------------------------
# Watchdog service  — restarts stuck downloads
# ---------------------------------------------------------------------------
class WatchdogService:
    """Monitors active downloads and restarts ones with no progress."""

    def __init__(self, items_ref: Callable, restart_fn: Callable,
                 timeout_min: int = 5) -> None:
        self._get_items = items_ref
        self._restart   = restart_fn
        self._timeout   = timeout_min * 60
        self._thread    = threading.Thread(target=self._run, daemon=True,
                                           name="fp-watchdog")

    def start(self) -> None:
        self._thread.start()

    def _run(self) -> None:
        while True:
            time.sleep(60)
            try:
                now = time.monotonic()
                for item in self._get_items():
                    if item.status != DownloadStatus.DOWNLOADING:
                        continue
                    since_progress = now - item._last_progress_time
                    if since_progress > self._timeout:
                        logger.warning("Watchdog: restarting stuck download %s", item.filename)
                        item.reset_for_retry()
                        self._restart(item)
            except Exception as e:
                logger.debug("Watchdog error: %s", e)


# ---------------------------------------------------------------------------
# REST API  — HTTP control on port 9100
# ---------------------------------------------------------------------------
class _RestHandler(http.server.BaseHTTPRequestHandler):
    """Minimal REST API for external control (n8n, scripts, etc.)."""
    app: "FetchProApp | None" = None

    def do_GET(self) -> None:
        if self.path == "/status":
            self._reply(200, self._status_payload())
        elif self.path == "/stats":
            self._reply(200, _STATS.data)
        elif self.path == "/queue":
            self._reply(200, self._queue_payload())
        elif self.path == "/health":
            self._reply(200, {"ok": True, "version": "5.5"})
        else:
            self._reply(404, {"error": "not found"})

    def do_POST(self) -> None:
        if self.path == "/add":
            try:
                length = int(self.headers.get("Content-Length", 0))
                body   = json.loads(self.rfile.read(length))
                url    = (body.get("url") or "").strip()
                if not url:
                    self._reply(400, {"error": "missing url"}); return
                fmt       = body.get("format", "")
                audio     = bool(body.get("audio", False))
                playlist  = bool(body.get("playlist", False))
                tags      = body.get("tags", "")
                if _RestHandler.app:
                    _RestHandler.app.after(0, lambda: _RestHandler.app._add_url_from_rest(
                        url, fmt, audio, playlist, tags))
                self._reply(200, {"ok": True, "url": url})
            except Exception as exc:
                self._reply(400, {"error": str(exc)})
        elif self.path == "/pause_all":
            if _RestHandler.app:
                _RestHandler.app.after(0, _RestHandler.app._pause_all)
            self._reply(200, {"ok": True})
        elif self.path == "/resume_all":
            if _RestHandler.app:
                _RestHandler.app.after(0, _RestHandler.app._resume_all)
            self._reply(200, {"ok": True})
        elif self.path == "/cancel_all":
            if _RestHandler.app:
                _RestHandler.app.after(0, _RestHandler.app._cancel_all)
            self._reply(200, {"ok": True})
        else:
            self._reply(404, {"error": "not found"})

    def do_OPTIONS(self) -> None:
        self.send_response(200)
        self._cors()
        self.end_headers()

    def _status_payload(self) -> dict:
        if not _RestHandler.app:
            return {}
        with _RestHandler.app._lock:
            items = list(_RestHandler.app._items)
        return {
            "active":    sum(1 for i in items if i.status in _ACTIVE),
            "queued":    sum(1 for i in items if i.status == DownloadStatus.QUEUED),
            "done":      sum(1 for i in items if i.status == DownloadStatus.DONE),
            "failed":    sum(1 for i in items if i.status == DownloadStatus.FAILED),
            "total":     len(items),
        }

    def _queue_payload(self) -> list:
        if not _RestHandler.app:
            return []
        with _RestHandler.app._lock:
            items = list(_RestHandler.app._items)
        return [
            {"url": i.url, "filename": i.filename,
             "status": i.status.name, "progress": round(i.progress, 1),
             "speed_bps": int(i.speed_bps), "tags": i.tags}
            for i in items
        ]

    def _reply(self, code: int, body: object) -> None:
        payload = json.dumps(body, ensure_ascii=False).encode()
        self.send_response(code)
        self._cors()
        self.send_header("Content-Type",   "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def _cors(self) -> None:
        origin = self.headers.get("Origin", "")
        # Only allow requests from localhost origins (prevents CSRF from remote sites)
        allowed = origin if (origin.startswith("http://localhost") or
                             origin.startswith("http://127.0.0.1")) else "http://127.0.0.1"
        self.send_header("Access-Control-Allow-Origin",  allowed)
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def log_message(self, *_) -> None:
        pass


def _start_rest_api(app: "FetchProApp") -> None:
    _RestHandler.app = app
    try:
        server = http.server.ThreadingHTTPServer(("127.0.0.1", REST_API_PORT), _RestHandler)
        t = threading.Thread(target=server.serve_forever, daemon=True, name="fp-rest-api")
        t.start()
        logger.info("REST API running on http://127.0.0.1:%d", REST_API_PORT)
    except OSError:
        logger.warning("REST API port %d in use — API disabled", REST_API_PORT)


# ---------------------------------------------------------------------------
# Main application window
# ---------------------------------------------------------------------------
class FetchProApp(tk.Tk):
    _PLACEHOLDER = _t("url_placeholder")

    def __init__(self) -> None:
        super().__init__()
        APP_DIR.mkdir(parents=True, exist_ok=True)
        STATE_DIR.mkdir(parents=True, exist_ok=True)

        self._settings  = Settings.load()
        if not self._settings.dark_theme:
            Theme._p = _LIGHT

        # Apply saved language
        global _LANG
        _LANG = self._settings.language

        self.title(f"{APP_NAME} v5.5 — מנהל הורדות")
        self.configure(bg=T.BG_DEEP)
        geo = self._settings.window_geometry or "960x740"
        self.geometry(geo)
        self.minsize(700, 520)
        self._apply_dark_titlebar()

        self._items:    list[DownloadItem]      = []
        self._cards:    dict[int, DownloadCard] = {}
        self._tabs:     dict[str, tk.Label]     = {}
        self._db        = HistoryDB(DB_PATH)
        self._semaphore = threading.Semaphore(self._settings.max_concurrent)
        self._executor  = ThreadPoolExecutor(max_workers=MAX_DL_WORKERS,
                                             thread_name_prefix="fp-dl")
        self._vfy_exec  = ThreadPoolExecutor(max_workers=MAX_VFY_WORKERS,
                                             thread_name_prefix="fp-vfy")
        self._lock      = threading.Lock()
        self._canvas_win_id: int | None = None
        self._flash_id:      str | None = None
        self._tray:          object | None = None
        self._pq            = PersistentQueue()
        self._shutdown_pending = False

        # Apply saved bandwidth limit
        _BW_LIMITER.set_limit(self._settings.global_bw_limit_kbps)

        _STATS.new_session()

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._build_ui()

        self.bind("<Configure>", self._on_configure)
        self.bind("<space>",      lambda _: self._kb_pause_resume())
        self.bind("<Delete>",     lambda _: self._kb_cancel_selected())
        self.bind("<r>",          lambda _: self._retry_all_failed())
        self.bind("<Control-a>",  lambda _: self._select_all())
        self.bind("<Control-v>",  lambda _: self._paste_clipboard())
        self.bind("<F5>",         lambda _: self._rebuild_list())

        try:
            self.drop_target_register("DND_Text")  # type: ignore
            self.dnd_bind("<<Drop>>", self._on_dnd_drop)  # type: ignore
        except Exception:
            pass

        self._schedule_refresh()

        self._last_clipboard = ""
        if self._settings.clipboard_monitor:
            self._start_clipboard_monitor()

        if self._settings.system_tray and TRAY_OK:
            threading.Thread(target=self._start_tray, daemon=True).start()

        threading.Thread(target=_start_bridge, args=(self,), daemon=True).start()

        # Watchdog
        self._watchdog = WatchdogService(
            items_ref    = lambda: list(self._items),
            restart_fn   = lambda item: self._executor.submit(
                _perform_download, item, self._thread_safe_update,
                self._semaphore, self._settings),
            timeout_min  = self._settings.watchdog_timeout_min,
        )
        self._watchdog.start()

        # REST API
        _start_rest_api(self)

        # Restore persistent queue
        if self._settings.persistent_queue:
            self.after(500, self._restore_queue)

        # Shutdown-after-queue monitor
        if self._settings.shutdown_after_queue:
            threading.Thread(target=self._shutdown_monitor, daemon=True,
                             name="fp-shutdown-mon").start()

    def _on_configure(self, event: tk.Event) -> None:
        if event.widget is not self:
            return
        # Debounce: cancel previous delayed save and reschedule
        if hasattr(self, "_geo_save_id") and self._geo_save_id:
            try:
                self.after_cancel(self._geo_save_id)
            except Exception:
                pass
        self._geo_save_id = self.after(500, self._save_geometry)

    def _save_geometry(self) -> None:
        try:
            geo = self.geometry()
            if geo and geo != self._settings.window_geometry:
                self._settings.window_geometry = geo
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Clipboard monitor
    # ------------------------------------------------------------------
    def _start_clipboard_monitor(self) -> None:
        def _poll() -> None:
            while True:
                time.sleep(1.5)
                try:
                    txt = self.clipboard_get().strip()
                    if txt != self._last_clipboard and txt.startswith(("http://","https://","ftp://")):
                        self._last_clipboard = txt
                        self.after(0, self._on_clipboard_url, txt)
                except Exception:
                    pass
        threading.Thread(target=_poll, daemon=True).start()

    def _on_clipboard_url(self, url: str) -> None:
        self._url_text.delete("1.0", "end")
        self._url_text.configure(fg=T.TEXT_MAIN)
        self._url_text.insert("1.0", url)
        self._flash(f"📋 URL זוהה בלוח: {url[:60]}{'…' if len(url)>60 else ''}")
        self.deiconify()
        self.lift()

    # ------------------------------------------------------------------
    # Keyboard shortcuts
    # ------------------------------------------------------------------
    def _kb_pause_resume(self) -> None:
        """Space — pause all active / resume all paused."""
        with self._lock:
            items = list(self._items)
        active = [i for i in items if i.status in _ACTIVE]
        paused = [i for i in items if i.status == DownloadStatus.PAUSED]
        if active:
            for i in active: self._pause_item(i)
            self._flash(f"⏸ הושהו {len(active)} הורדות")
        elif paused:
            for i in paused: self._resume_item(i)
            self._flash(f"▶ הומשכו {len(paused)} הורדות")

    def _kb_cancel_selected(self) -> None:
        """Delete — cancel all active downloads."""
        with self._lock:
            items = list(self._items)
        active = [i for i in items if i.status in _ACTIVE]
        for i in active:
            self._cancel_item(i)
        if active:
            self._flash(f"✕ בוטלו {len(active)} הורדות")

    def _select_all(self) -> None:
        """Ctrl+A — paste clipboard."""
        self._url_text.focus_set()
        self._url_text.tag_add("sel", "1.0", "end")

    # ------------------------------------------------------------------
    # Batch operations
    # ------------------------------------------------------------------
    def _retry_all_failed(self) -> None:
        with self._lock:
            items = list(self._items)
        failed = [i for i in items if i.status == DownloadStatus.FAILED]
        for i in failed:
            self._retry_item(i)
        if failed:
            self._flash(f"↺ {len(failed)} הורדות נכשלות — מנסה שנית")
        else:
            self._flash("אין הורדות שנכשלו")

    # ------------------------------------------------------------------
    # CSV export
    # ------------------------------------------------------------------
    def _export_history_csv(self) -> None:
        import csv as _csv
        path = filedialog.asksaveasfilename(
            title=_t("export_history_csv"),
            defaultextension=".csv",
            filetypes=[("CSV","*.csv"),("כל הקבצים","*.*")],
        )
        if not path:
            return
        rows = self._db.fetch(limit=10000)
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = _csv.DictWriter(f, fieldnames=["id","filename","url","save_dir","status","size","hash","finished"])
                writer.writeheader()
                writer.writerows(rows)
            self._flash(f"✓ ייוצאו {len(rows)} רשומות → {Path(path).name}")
        except OSError as e:
            messagebox.showerror("שגיאת ייצוא", str(e), parent=self)

    # ------------------------------------------------------------------
    # Platform / tray
    # ------------------------------------------------------------------
    def _apply_dark_titlebar(self) -> None:
        try:
            import ctypes
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, 20, ctypes.byref(ctypes.c_int(1 if T.is_dark() else 0)), 4)
        except Exception:
            pass

    def _start_tray(self) -> None:
        try:
            img = PilImage.new("RGBA", (64, 64), (0,0,0,0))
            from PIL import ImageDraw
            d = ImageDraw.Draw(img)
            d.ellipse((4,4,60,60), fill=(0,229,255,255))
            d.polygon([(32,16),(48,42),(16,42)], fill=(7,11,18,255))
            icon = pystray.Icon(APP_NAME, img, APP_NAME, menu=pystray.Menu(
                pystray.MenuItem("פתח FetchPro", lambda: self.after(0, self.deiconify), default=True),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("צא לגמרי", lambda: self.after(0, self._quit_app)),
            ))
            self._tray = icon
            icon.run()
        except Exception as e:
            logger.warning("Tray failed: %s", e)

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        # Header
        hdr = tk.Frame(self, bg=T.BG_DEEP, pady=14)
        hdr.pack(fill="x", padx=20)

        logo = tk.Frame(hdr, bg=T.BG_DEEP)
        logo.pack(side="left")
        tk.Label(logo, text="⬇", bg=T.BG_DEEP, fg=T.ACCENT,
                 font=("Segoe UI Emoji",20)).pack(side="left", padx=8)
        tit = tk.Frame(logo, bg=T.BG_DEEP)
        tit.pack(side="left")
        tk.Label(tit, text="FETCHPRO", bg=T.BG_DEEP, fg=T.TEXT_MAIN,
                 font=("Consolas",16,"bold")).pack(anchor="w")
        tk.Label(tit, text="v5.5 — מנהל הורדות מקצועי", bg=T.BG_DEEP, fg=T.TEXT_DIM,
                 font=("Consolas",8)).pack(anchor="w")

        # Stats
        sf = tk.Frame(hdr, bg=T.BG_DEEP)
        sf.pack(side="right")
        self._stat_labels: dict[str, tk.Label] = {}
        for key, label, color in [("active","פעיל",T.ACCENT),("done","הושלם",T.GREEN),
                                   ("failed","שגיאה",T.RED),("speed","מהירות",T.TEXT_DIM)]:
            box = tk.Frame(sf, bg=T.BG_CARD, highlightthickness=1, highlightbackground=T.BORDER,
                           padx=10, pady=6)
            box.pack(side="left", padx=3)
            v = tk.Label(box, text="0", bg=T.BG_CARD, fg=color, font=("Consolas",14,"bold"))
            v.pack()
            tk.Label(box, text=label, bg=T.BG_CARD, fg=T.TEXT_DIM, font=("Consolas",7)).pack()
            self._stat_labels[key] = v

        tk.Frame(self, bg=T.BORDER, height=1).pack(fill="x")

        # Toolbar
        toolbar = tk.Frame(self, bg=T.BG_DEEP, padx=16, pady=6)
        toolbar.pack(fill="x")

        # ── Helper: create a small toolbar button ──────────────────────────
        def _tb_btn(parent, text, cmd, fg=None, bg_hover=None):
            fg = fg or T.TEXT_DIM
            b = tk.Label(parent, text=text, bg=T.BG_DEEP, fg=fg,
                         font=("Consolas", 9), padx=9, pady=3,
                         cursor="hand2", highlightthickness=1,
                         highlightbackground=T.BORDER)
            b.bind("<Button-1>", lambda _, c=cmd: c())
            b.bind("<Enter>",    lambda _, w=b: w.configure(bg=bg_hover or T.BG_HOVER))
            b.bind("<Leave>",    lambda _, w=b: w.configure(bg=T.BG_DEEP))
            b.pack(side="left", padx=2)
            return b

        # ── Primary buttons (always visible) ───────────────────────────────
        self._tb_settings = _tb_btn(toolbar, _t("settings"), self._open_settings)
        self._tb_media    = _tb_btn(toolbar, _t("media"),
                                    self._open_media_dialog,
                                    fg=T.ACCENT if YTDLP_OK else T.TEXT_MUTED)
        self._tb_retry    = _tb_btn(toolbar, _t("retry_all"),
                                    self._retry_all_failed, fg=T.YELLOW)
        self._tb_theme    = _tb_btn(toolbar, _t("theme"), self._toggle_theme)

        # ── Separator ──────────────────────────────────────────────────────
        tk.Frame(toolbar, bg=T.BORDER, width=1, height=18).pack(side="left", padx=6)

        # ── Language quick-picker (inline flags) ───────────────────────────
        self._tb_lang = _tb_btn(toolbar, "🌍 " + SUPPORTED_LANGS.get(_LANG, "שפה"),
                                self._open_language_dialog)

        # ── Separator ──────────────────────────────────────────────────────
        tk.Frame(toolbar, bg=T.BORDER, width=1, height=18).pack(side="left", padx=6)

        # ── ⋯ More dropdown ────────────────────────────────────────────────
        def _show_more_menu(event=None):
            menu = tk.Menu(self, tearoff=0,
                           bg=T.BG_CARD, fg=T.TEXT_MAIN,
                           activebackground=T.BG_HOVER, activeforeground=T.ACCENT,
                           font=("Consolas", 9), bd=0, relief="flat",
                           selectcolor=T.ACCENT)
            menu.add_command(label=_t("history"),    command=self._open_history)
            menu.add_command(label=_t("stats"),      command=self._open_stats)
            menu.add_separator()
            menu.add_command(label=_t("export"),     command=self._export_urls)
            menu.add_command(label=_t("import_btn"), command=self._import_urls)
            menu.add_separator()
            if LIBTORRENT_OK:
                menu.add_command(label=_t("torrent"), command=self._add_torrent_file)
            menu.add_separator()
            menu.add_command(label=_t("csv_export_btn"), command=self._export_history_csv)

            # Position below the ⋯ button
            x = more_btn.winfo_rootx()
            y = more_btn.winfo_rooty() + more_btn.winfo_height() + 2
            menu.tk_popup(x, y)

        more_btn = _tb_btn(toolbar, "⋯ עוד", _show_more_menu)
        self._tb_more = more_btn

        # ── Right side: Quit ───────────────────────────────────────────────
        quit_btn = tk.Label(toolbar, text=_t("quit"), bg=T.BG_DEEP, fg=T.RED,
                            font=("Consolas", 9), padx=9, pady=3,
                            cursor="hand2", highlightthickness=1,
                            highlightbackground=T.BORDER)
        quit_btn.bind("<Button-1>", lambda _: self._quit_app())
        quit_btn.bind("<Enter>",    lambda _: quit_btn.configure(bg="#2a0808"))
        quit_btn.bind("<Leave>",    lambda _: quit_btn.configure(bg=T.BG_DEEP))
        quit_btn.pack(side="right", padx=2)
        self._tb_quit = quit_btn

        # ── Keyboard hint (compact) ────────────────────────────────────────
        hints = tk.Label(toolbar, text="Space · Del · R · Ctrl+V",
                         bg=T.BG_DEEP, fg=T.TEXT_MUTED, font=("Consolas", 7))
        hints.pack(side="right", padx=8)
        self._tb_hints = hints

        # Input
        inp = tk.Frame(self, bg=T.BG_DEEP, padx=20, pady=12)
        inp.pack(fill="x")

        # Dir row
        dr = tk.Frame(inp, bg=T.BG_DEEP)
        dr.pack(fill="x", pady=(0,8))
        tk.Label(dr, text=_t("save_dir_label"), bg=T.BG_DEEP, fg=T.TEXT_DIM,
                 font=("Consolas",9)).pack(side="left")
        self._dir_lbl = tk.Label(dr, text=self._trunc(Path(self._settings.save_dir)),
                                 bg=T.BG_DEEP, fg=T.ACCENT, font=("Consolas",9), cursor="hand2")
        self._dir_lbl.pack(side="left", padx=6)
        self._dir_lbl.bind("<Button-1>", lambda _: self._choose_dir())
        self._flat_btn(dr, "שנה", T.ACCENT, self._choose_dir).pack(side="left")

        # Concurrent slider
        sc = tk.Frame(inp, bg=T.BG_DEEP)
        sc.pack(fill="x", pady=(0,8))
        tk.Label(sc, text=_t("max_dl_label"), bg=T.BG_DEEP, fg=T.TEXT_DIM,
                 font=("Consolas",9)).pack(side="left")
        self._concurrent_var = tk.IntVar(value=self._settings.max_concurrent)
        self._concurrent_lbl = tk.Label(sc, text=str(self._settings.max_concurrent),
                                        bg=T.BG_DEEP, fg=T.ACCENT, font=("Consolas",9,"bold"), width=2)
        self._concurrent_lbl.pack(side="left", padx=4)
        sl = tk.Scale(sc, from_=1, to=16, orient="horizontal",
                      variable=self._concurrent_var,
                      command=self._on_concurrent_change,
                      bg=T.BG_DEEP, fg=T.TEXT_DIM, highlightthickness=0,
                      troughcolor=T.BORDER, activebackground=T.ACCENT,
                      length=140, showvalue=False)
        sl.pack(side="left")

        # URL textarea
        uc = tk.Frame(inp, bg=T.BG_CARD, highlightthickness=1, highlightbackground=T.BORDER,
                      padx=2, pady=2)
        uc.pack(fill="x", pady=(0,8))
        self._url_text = tk.Text(uc, bg=T.BG_CARD, fg=T.TEXT_MAIN,
                                 insertbackground=T.ACCENT, font=("Consolas",10),
                                 height=3, relief="flat", padx=8, pady=6,
                                 wrap="char", selectbackground=T.ACCENT, selectforeground=T.BG_DEEP)
        self._url_text.pack(fill="x")
        self._url_text.insert("1.0", self._PLACEHOLDER)
        self._url_text.configure(fg=T.TEXT_DIM)
        self._url_text.bind("<FocusIn>",       self._url_focus_in)
        self._url_text.bind("<FocusOut>",      self._url_focus_out)
        self._url_text.bind("<Control-Return>", lambda _: self._add_downloads())  # quick-add, no dialog
        self._url_text.bind("<Button-3>",       self._ctx_menu)

        # Button row
        br = tk.Frame(inp, bg=T.BG_DEEP)
        br.pack(fill="x")
        self._accent_btn(br, _t("add_download_adv"), self._add_advanced).pack(side="right")  # opens dialog
        self._flat_btn(br, "נקה שהושלמו", T.GREEN, self._clear_done).pack(side="right", padx=6)
        self._flat_btn(br, "עצור הכל", T.RED, self._cancel_all).pack(side="right", padx=6)
        self._paste_btn = self._flat_btn(br, _t("paste"), T.ACCENT, self._paste_clipboard)
        self._paste_btn.pack(side="left", padx=(0,6))
        self._flat_btn(br, "🔍 אמת", T.YELLOW, self._verify_urls).pack(side="left", padx=(0,6))
        tk.Label(br, text=_t("ctrl_enter_hint"), bg=T.BG_DEEP, fg=T.TEXT_MUTED,
                 font=("Consolas",8)).pack(side="left")

        # Filter tabs
        fr = tk.Frame(self, bg=T.BG_DEEP, padx=20, pady=4)
        fr.pack(fill="x")
        self._filter_var = tk.StringVar(value="all")
        for fk, fl in [("all","הכל"),("downloading","⬇ פעיל"),("done","✓ הושלם"),
                        ("failed","✗ שגיאה"),("paused","⏸ מושהה"),("queued","📅 מתוזמן")]:
            tab = tk.Label(fr, text=fl, cursor="hand2", font=("Consolas",9), padx=10, pady=3)
            tab.bind("<Button-1>", lambda _, k=fk: self._set_filter(k))
            self._style_tab(tab, fk)
            tab.pack(side="left", padx=2)
            self._tabs[fk] = tab

        # About tab (right-aligned, opens license popup)
        about_tab = tk.Label(fr, text=_t("about"), cursor="hand2",
                             font=("Consolas", 9), padx=10, pady=3,
                             bg=T.BG_DEEP, fg=T.TEXT_DIM,
                             highlightthickness=1, highlightbackground=T.BORDER)
        about_tab.bind("<Button-1>", lambda _: self._open_about())
        about_tab.bind("<Enter>", lambda _, w=about_tab: w.configure(bg=T.BG_HOVER))
        about_tab.bind("<Leave>", lambda _, w=about_tab: w.configure(bg=T.BG_DEEP))
        about_tab.pack(side="right", padx=2)

        tk.Frame(self, bg=T.BORDER, height=1).pack(fill="x")

        # List
        lc = tk.Frame(self, bg=T.BG_DEEP)
        lc.pack(fill="both", expand=True, padx=20, pady=12)
        self._canvas = tk.Canvas(lc, bg=T.BG_DEEP, highlightthickness=0)
        vsb = tk.Scrollbar(lc, orient="vertical", command=self._canvas.yview)
        self._list_frame = tk.Frame(self._canvas, bg=T.BG_DEEP)
        self._list_frame.bind("<Configure>",
                              lambda _: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas_win_id = self._canvas.create_window((0,0), window=self._list_frame,
                                                          anchor="nw", width=860)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)
        self._canvas.bind("<MouseWheel>", self._on_scroll)
        self._canvas.bind("<Button-4>", lambda _: self._canvas.yview_scroll(-1,"units"))
        self._canvas.bind("<Button-5>", lambda _: self._canvas.yview_scroll(1,"units"))
        self.bind("<Configure>", self._on_resize)

        # Status bar
        sb = tk.Frame(self, bg="#0a0f1a", pady=5)
        sb.pack(fill="x", side="bottom")
        self._status_lbl = tk.Label(sb, text=_t("ready"), bg="#0a0f1a", fg=T.TEXT_DIM,
                                    font=("Consolas",8), padx=14)
        self._status_lbl.pack(side="left")
        self._total_lbl = tk.Label(sb, text="", bg="#0a0f1a", fg=T.TEXT_DIM,
                                   font=("Consolas",8), padx=14)
        self._total_lbl.pack(side="right")
        if not YTDLP_OK:
            tk.Label(sb, text=_t("no_ytdlp_short"), bg="#0a0f1a", fg=T.TEXT_MUTED,
                     font=("Consolas",7)).pack(side="right", padx=8)
        if not LIBTORRENT_OK:
            tk.Label(sb, text=_t("no_libtorrent"), bg="#0a0f1a", fg=T.TEXT_MUTED,
                     font=("Consolas",7)).pack(side="right", padx=8)

        self._empty_lbl = tk.Label(self._list_frame, bg=T.BG_DEEP, fg=T.TEXT_MUTED,
                                   text="⬇\n\nגרור קישורים, קבצי .torrent, או magnet לחלון\nהדבק למעלה, או לחץ '+ הוסף להורדה'",
                                   font=("Consolas",10), justify="center")
        self._empty_lbl.pack(pady=80)

    # ------------------------------------------------------------------
    # Widget factories
    # ------------------------------------------------------------------
    def _flat_btn(self, parent: tk.Widget, text: str, color: str,
                  cmd: Callable) -> tk.Label:
        b = tk.Label(parent, text=text, cursor="hand2", bg=T.BG_DEEP, fg=color,
                     font=("Consolas",9), padx=10, pady=4,
                     highlightthickness=1, highlightbackground=T.BORDER)
        b.bind("<Button-1>", lambda _: cmd())
        b.bind("<Enter>", lambda _, w=b: w.configure(bg=T.BG_HOVER))
        b.bind("<Leave>", lambda _, w=b: w.configure(bg=T.BG_DEEP))
        return b

    def _accent_btn(self, parent: tk.Widget, text: str, cmd: Callable) -> tk.Label:
        b = tk.Label(parent, text=text, cursor="hand2", bg="#003d4d", fg=T.ACCENT,
                     font=("Consolas",10,"bold"), padx=16, pady=5,
                     highlightthickness=1, highlightbackground=T.ACCENT)
        b.bind("<Button-1>", lambda _: cmd())
        b.bind("<Enter>", lambda _, w=b: w.configure(bg="#005566"))
        b.bind("<Leave>", lambda _, w=b: w.configure(bg="#003d4d"))
        return b

    def _style_tab(self, lbl: tk.Label, key: str) -> None:
        active = self._filter_var.get() == key
        lbl.configure(bg=T.ACCENT_BG if active else T.BG_DEEP,
                      fg=T.ACCENT    if active else T.TEXT_DIM)

    # ------------------------------------------------------------------
    # Events
    # ------------------------------------------------------------------
    def _on_resize(self, _: tk.Event) -> None:
        if self._canvas_win_id is not None:
            self._canvas.itemconfigure(self._canvas_win_id,
                                       width=max(self._canvas.winfo_width()-4, 400))

    def _on_scroll(self, e: tk.Event) -> None:
        units = -(e.delta // 120) if abs(e.delta) >= 120 else -e.delta
        if units:
            self._canvas.yview_scroll(int(units), "units")

    def _on_concurrent_change(self, val: str) -> None:
        n = int(float(val))
        self._concurrent_lbl.configure(text=str(n))
        self._settings.max_concurrent = n
        self._semaphore = threading.Semaphore(n)
        self._settings.save()  # persist immediately

    def _on_dnd_drop(self, event: object) -> None:
        try:
            data: str = getattr(event, "data", "")
            for token in data.split():
                token = token.strip()
                if token.startswith(("http","ftp","magnet:")):
                    self._enqueue_item(DownloadItem(url=token,
                                                    save_dir=Path(self._settings.save_dir)))
                elif token.lower().endswith(".torrent"):
                    # Local file dropped — use file path directly
                    p = Path(token)
                    if p.exists():
                        self._enqueue_item(DownloadItem(url=str(p),
                                                        save_dir=Path(self._settings.save_dir)))
        except Exception as e:
            logger.warning("DnD error: %s", e)

    # ------------------------------------------------------------------
    # URL textarea
    # ------------------------------------------------------------------
    def _url_focus_in(self, _: tk.Event) -> None:
        if self._url_text.get("1.0","end-1c") == self._PLACEHOLDER:
            self._url_text.delete("1.0","end")
            self._url_text.configure(fg=T.TEXT_MAIN)

    def _url_focus_out(self, _: tk.Event | None) -> None:
        if not self._url_text.get("1.0","end-1c").strip():
            self._url_text.insert("1.0", self._PLACEHOLDER)
            self._url_text.configure(fg=T.TEXT_DIM)

    def _read_urls(self) -> list[str]:
        raw = self._url_text.get("1.0","end-1c").strip()
        if raw == self._PLACEHOLDER:
            return []
        results = []
        for line in raw.splitlines():
            u = line.strip()
            if u.startswith(("http://","https://","ftp://","ftps://","magnet:")):
                results.append(u)
            elif u.lower().endswith(".torrent") and Path(u).exists():
                results.append(u)   # local .torrent file path
        return results

    def _paste_clipboard(self) -> None:
        try:
            text = self.clipboard_get().strip()
        except tk.TclError:
            text = ""
        if not text:
            self._flash("הלוח ריק")
            return
        current = self._url_text.get("1.0","end-1c")
        if current == self._PLACEHOLDER or not current.strip():
            self._url_text.delete("1.0","end")
            self._url_text.configure(fg=T.TEXT_MAIN)
            self._url_text.insert("1.0", text)
        else:
            if not current.endswith("\n"):
                self._url_text.insert("end","\n")
            self._url_text.insert("end", text)
        self._url_text.see("end")
        self._paste_btn.configure(text=_t("pasted_ok"), fg=T.GREEN)
        self.after(1200, lambda: self._paste_btn.configure(text=_t("paste"), fg=T.ACCENT))

    def _ctx_menu(self, event: tk.Event) -> None:
        m = tk.Menu(self, tearoff=0, bg=T.BG_CARD, fg=T.TEXT_MAIN,
                    activebackground=T.ACCENT_BG, activeforeground=T.ACCENT,
                    font=("Consolas",9), bd=0)
        has_sel = False
        try:
            self._url_text.index("sel.first")
            has_sel = True
        except tk.TclError:
            pass
        is_ph = self._url_text.get("1.0","end-1c") == self._PLACEHOLDER
        def _sel_all():
            self._url_text.tag_add("sel","1.0","end-1c")
        def _clear():
            self._url_text.delete("1.0","end")
            self._url_focus_out(None)
        m.add_command(label=f"  {_t('paste_short')}", command=self._paste_clipboard)
        m.add_command(label=f"  {_t('copy_short')}", state="normal" if has_sel else "disabled",
                      command=lambda: self._url_text.event_generate("<<Copy>>"))
        m.add_command(label=f"  {_t('cut_short')}", state="normal" if has_sel and not is_ph else "disabled",
                      command=lambda: self._url_text.event_generate("<<Cut>>"))
        m.add_separator()
        m.add_command(label=f"  {_t('select_all')}", state="disabled" if is_ph else "normal", command=_sel_all)
        m.add_command(label=f"  {_t('clear_all')}", state="disabled" if is_ph else "normal", command=_clear)
        m.add_separator()
        m.add_command(label=f"  {_t('verify_urls_btn')}", command=self._verify_urls)
        m.tk_popup(event.x_root, event.y_root)
        m.grab_release()

    # ------------------------------------------------------------------
    # Directory
    # ------------------------------------------------------------------
    def _choose_dir(self) -> None:
        d = filedialog.askdirectory(title=_t("select_folder"),
                                    initialdir=self._settings.save_dir)
        if d:
            self._settings.save_dir = d
            self._dir_lbl.configure(text=self._trunc(Path(d)))
            self._settings.save()  # persist immediately

    @staticmethod
    def _trunc(p: Path, n: int = 55) -> str:
        s = str(p)
        return ("…" + s[-(n-1):]) if len(s) > n else s

    # ------------------------------------------------------------------
    # Download management
    # ------------------------------------------------------------------
    def _add_torrent_file(self) -> None:
        """Open file dialog to pick a .torrent file and enqueue it."""
        if not LIBTORRENT_OK:
            messagebox.showwarning(
                "libtorrent לא מותקן",
                "כדי להוריד טורנטים, התקן את הספרייה:\n\n"
                "  pip install libtorrent\n\n"
                "לאחר ההתקנה הפעל מחדש את FetchPro.",
                parent=self,
            )
            return
        path = _open_torrent_file(self._settings.save_dir)
        if path:
            item = DownloadItem(url=path, save_dir=Path(self._settings.save_dir))
            self._enqueue_item(item)
            self._flash(f"נוסף טורנט: {Path(path).name}")

    def _add_downloads(self) -> None:
        """Quick-add: add all URLs directly without dialog (Ctrl+Enter shortcut)."""
        urls = self._read_urls()
        if not urls:
            self._flash("לא נמצאו קישורים תקינים")
            return
        save_dir = Path(self._settings.save_dir)
        save_dir.mkdir(parents=True, exist_ok=True)

        for url in urls:
            item = DownloadItem(
                url=url, save_dir=save_dir,
                hash_algo=self._settings.default_hash_algo if self._settings.verify_hash else "",
                multipart=self._settings.multipart,
                auto_extract=self._settings.auto_extract,
                auto_open=self._settings.auto_open,
            )
            self._enqueue_item(item)

        self._url_text.delete("1.0","end")
        self._url_focus_out(None)
        self._flash(f"נוספו {len(urls)} קבצים להורדה")

    def _add_advanced(self) -> None:
        """Open the options/schedule dialog for each URL (main + button)."""
        urls = self._read_urls()
        if not urls:
            self._flash("לא נמצאו קישורים תקינים")
            return
        save_dir = Path(self._settings.save_dir)
        save_dir.mkdir(parents=True, exist_ok=True)
        # Open dialog for first URL; remaining are quick-added after
        def _on_add(item: DownloadItem) -> None:
            self._enqueue_item(item)
            rest = urls[1:]
            for u in rest:
                extra = DownloadItem(
                    url=u, save_dir=save_dir,
                    hash_algo=self._settings.default_hash_algo if self._settings.verify_hash else "",
                    multipart=self._settings.multipart,
                    auto_extract=self._settings.auto_extract,
                    auto_open=self._settings.auto_open,
                )
                self._enqueue_item(extra)
            if rest:
                self._flash(f"נוספו עוד {len(rest)} קבצים ישירות")
            self._url_text.delete("1.0","end")
            self._url_focus_out(None)
        AddDownloadDialog(self, urls[0], self._settings, _on_add)

    def _enqueue_item(self, item: DownloadItem) -> None:
        """Deduplicate filename, check for URL duplicates, sort by priority, submit."""
        save_dir = item.save_dir
        if not item.filename or item.filename == "download":
            item.filename = _derive_filename(item.url)
        # Sanitize filename
        item.filename = _sanitize_filename(item.filename)
        item.filename = _deduplicate_filename(item.filename, save_dir)

        # Warn on duplicate URL
        if self._settings.warn_duplicates:
            with self._lock:
                existing_urls = [i.url for i in self._items]
            if item.url in existing_urls:
                if not messagebox.askyesno(
                    "כפילות",
                    f"הקישור כבר נמצא בתור:\n{item.url[:80]}\n\nלהוסיף שוב?",
                    parent=self
                ):
                    return

        with self._lock:
            self._items.append(item)
            # Keep HIGH priority items at top of visual list
            self._items.sort(key=lambda i: (
                0 if i.dl_priority == "HIGH" else
                1 if i.dl_priority == "NORMAL" else 2
            ))

        self._executor.submit(
            _perform_download, item,
            self._thread_safe_update,
            self._semaphore,
            self._settings,
        )

        # Event-based wait instead of busy-poll every 0.5s
        done_event = threading.Event()
        item._done_event = done_event  # type: ignore[attr-defined]

        def _record_done() -> None:
            done_event.wait(timeout=86400)   # max 24h
            self._db.record(item)
        threading.Thread(target=_record_done, daemon=True, name=f"fp-rec-{item.filename[:20]}").start()

        self._rebuild_list()

    def _thread_safe_update(self, item: DownloadItem) -> None:
        try:
            self.after(0, self._refresh_card, item)
        except RuntimeError:
            pass
        # Signal the history-recording thread when done
        if item.status in _TERMINAL:
            ev = getattr(item, "_done_event", None)
            if ev is not None:
                ev.set()
            # Record stats on completion
            if item.status == DownloadStatus.DONE:
                _STATS.record_done(item)
            # Auto-save queue on any terminal state
            self.after(100, self._save_queue)

    def _refresh_card(self, item: DownloadItem) -> None:
        card = self._cards.get(id(item))
        if card and card.winfo_exists():
            card.refresh()
        self._update_stats()

    def _pause_item(self, item: DownloadItem) -> None:
        item.pause()
        self._refresh_card(item)

    def _resume_item(self, item: DownloadItem) -> None:
        # Just clear the pause event — the existing download thread wakes up and continues.
        # Do NOT submit a new _perform_download or two threads will download simultaneously.
        item.resume()
        self._refresh_card(item)

    def _cancel_item(self, item: DownloadItem) -> None:
        item.cancel(); self._refresh_card(item)

    def _retry_item(self, item: DownloadItem) -> None:
        item.reset_for_retry()
        self._executor.submit(_perform_download, item,
                              self._thread_safe_update, self._semaphore, self._settings)
        self._refresh_card(item)

    def _open_item(self, item: DownloadItem) -> None:
        if item.destination.exists():
            _open_file(item.destination)
        else:
            self._flash(f"הקובץ לא נמצא: {item.filename}")

    def _remove_item(self, item: DownloadItem) -> None:
        item.cancel()
        with self._lock:
            self._items = [i for i in self._items if i is not item]
        card = self._cards.pop(id(item), None)
        if card and card.winfo_exists():
            card.destroy()
        self._update_stats()
        self._check_empty()

    def _clear_done(self) -> None:
        for item in [i for i in self._items if i.status == DownloadStatus.DONE]:
            self._remove_item(item)

    def _cancel_all(self) -> None:
        _stoppable = _ACTIVE | {DownloadStatus.PAUSED, DownloadStatus.PENDING, DownloadStatus.QUEUED}
        with self._lock:
            targets = [i for i in self._items if i.status in _stoppable]
        for item in targets:
            item.cancel()
        for item in targets:
            self._refresh_card(item)

    # Priority
    def _move_up(self, item: DownloadItem) -> None:
        with self._lock:
            idx = self._items.index(item)
            if idx > 0:
                self._items[idx], self._items[idx-1] = self._items[idx-1], self._items[idx]
        self._rebuild_list()

    def _move_down(self, item: DownloadItem) -> None:
        with self._lock:
            idx = self._items.index(item)
            if idx < len(self._items) - 1:
                self._items[idx], self._items[idx+1] = self._items[idx+1], self._items[idx]
        self._rebuild_list()

    # ------------------------------------------------------------------
    # Import / Export
    # ------------------------------------------------------------------
    def _export_urls(self) -> None:
        if not self._items:
            self._flash("אין הורדות לייצא")
            return
        path = filedialog.asksaveasfilename(
            title=_t("export_links"), defaultextension=".json",
            filetypes=[("JSON","*.json"),("Text","*.txt"),("All","*")])
        if not path:
            return
        data = [{"url": i.url, "filename": i.filename, "status": i.status.name}
                for i in self._items]
        Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2))
        self._flash(f"יוצאו {len(data)} קישורים → {Path(path).name}")

    def _import_urls(self) -> None:
        path = filedialog.askopenfilename(
            title=_t("import_links"),
            filetypes=[("JSON","*.json"),("Text","*.txt"),("All","*")])
        if not path:
            return
        try:
            text = Path(path).read_text()
            try:
                data = json.loads(text)
                urls = [e["url"] if isinstance(e, dict) else str(e) for e in data]
            except json.JSONDecodeError:
                urls = [l.strip() for l in text.splitlines()
                        if l.strip().startswith(("http","ftp"))]

            self._url_text.delete("1.0","end")
            self._url_text.configure(fg=T.TEXT_MAIN)
            self._url_text.insert("1.0","\n".join(urls))
            self._flash(f"יובאו {len(urls)} קישורים")
        except Exception as e:
            self._flash(f"שגיאת ייבוא: {e}")

    # ------------------------------------------------------------------
    # Verify dialog
    # ------------------------------------------------------------------
    def _verify_urls(self) -> None:
        urls = self._read_urls()
        if not urls:
            self._flash("אין קישורים לאימות")
            return

        win = tk.Toplevel(self)
        win.title(_t("verify_links"))
        win.configure(bg=T.BG_DEEP)
        win.geometry("740x460")
        win.minsize(560,320)
        win.grab_set()

        hdr = tk.Frame(win, bg=T.BG_DEEP, padx=20, pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text=_t("verify_links"), bg=T.BG_DEEP, fg=T.TEXT_MAIN,
                 font=("Consolas",12,"bold")).pack(side="left")
        summary_lbl = tk.Label(hdr, text=f"בודק {len(urls)} קישורים...",
                               bg=T.BG_DEEP, fg=T.TEXT_DIM, font=("Consolas",9))
        summary_lbl.pack(side="right")

        tk.Frame(win, bg=T.BORDER, height=1).pack(fill="x")
        outer = tk.Frame(win, bg=T.BG_DEEP)
        outer.pack(fill="both",expand=True,padx=14,pady=10)
        vc = tk.Canvas(outer, bg=T.BG_DEEP, highlightthickness=0)
        vs = tk.Scrollbar(outer, orient="vertical", command=vc.yview)
        vi = tk.Frame(vc, bg=T.BG_DEEP)
        vi.bind("<Configure>", lambda _: vc.configure(scrollregion=vc.bbox("all")))
        wid = vc.create_window((0,0), window=vi, anchor="nw")
        vc.configure(yscrollcommand=vs.set)
        vc.bind("<Configure>", lambda e: vc.itemconfigure(wid, width=e.width))
        def _vscroll(e: tk.Event) -> None:
            u = -(e.delta//120) if abs(e.delta)>=120 else -e.delta
            if u: vc.yview_scroll(int(u),"units")
        vc.bind("<MouseWheel>", _vscroll)
        vc.bind("<Button-4>", lambda _: vc.yview_scroll(-1,"units"))
        vc.bind("<Button-5>", lambda _: vc.yview_scroll(1,"units"))
        vs.pack(side="right",fill="y"); vc.pack(side="left",fill="both",expand=True)

        row_data: list[dict] = []
        for url in urls:
            row = tk.Frame(vi, bg=T.BG_CARD, highlightthickness=1, highlightbackground=T.BORDER)
            row.pack(fill="x", pady=3)
            ind = tk.Label(row, text="⏳", bg=T.BG_CARD, font=("Segoe UI Emoji",11), width=3)
            ind.pack(side="left", padx=6, pady=6)
            tf = tk.Frame(row, bg=T.BG_CARD)
            tf.pack(side="left",fill="x",expand=True,pady=4)
            short = (url[:70]+"…") if len(url)>70 else url
            tk.Label(tf,text=short,bg=T.BG_CARD,fg=T.TEXT_MAIN,font=("Consolas",9),anchor="w").pack(fill="x")
            msg = tk.Label(tf,text=_t("verifying"),bg=T.BG_CARD,fg=T.TEXT_DIM,font=("Consolas",8),anchor="w")
            msg.pack(fill="x")
            row_data.append({"url":url,"ind":ind,"msg":msg,"row":row,"result":VerifyResult.PENDING})

        tk.Frame(win, bg=T.BORDER, height=1).pack(fill="x")
        footer = tk.Frame(win, bg=T.BG_DEEP, padx=16, pady=10)
        footer.pack(fill="x")

        add_btn_ref: list[tk.Label] = []
        def _add_valid() -> None:
            valid = [rd["url"] for rd in row_data
                     if rd["result"] in (VerifyResult.OK, VerifyResult.WARNING)]
            win.destroy()
            if not valid:
                self._flash("כל הקישורים נכשלו")
                return
            self._url_text.delete("1.0","end")
            self._url_text.configure(fg=T.TEXT_MAIN)
            self._url_text.insert("1.0","\n".join(valid))
            self._add_downloads()

        add_btn = self._accent_btn(footer,_t("add_valid"), _add_valid)
        add_btn.pack(side="right")
        add_btn.configure(fg=T.TEXT_DIM, bg=T.BG_CARD, highlightbackground=T.BORDER)
        add_btn_ref.append(add_btn)
        self._flat_btn(footer,_t("close"),T.TEXT_DIM,win.destroy).pack(side="right",padx=6)

        lock = threading.Lock()
        completed=[0]; ok_n=[0]; warn_n=[0]; fail_n=[0]

        def _on_result(idx: int, info: VerifyInfo) -> None:
            def _upd() -> None:
                if not win.winfo_exists(): return
                rd = row_data[idx]
                match info.result:
                    case VerifyResult.OK:
                        icon,color,border="✅",T.GREEN,T.GREEN_DIM; ok_n[0]+=1
                    case VerifyResult.WARNING:
                        icon,color,border="⚠️",T.YELLOW,"#2a2a0a"; warn_n[0]+=1
                    case _:
                        icon,color,border="❌",T.RED,"#3a0a0a"; fail_n[0]+=1
                rd["result"]=info.result
                rd["ind"].configure(text=icon)
                rd["msg"].configure(text=info.message, fg=color)
                rd["row"].configure(highlightbackground=border)
                with lock:
                    completed[0]+=1; done=completed[0]
                if done>=len(urls):
                    if summary_lbl.winfo_exists():
                        summary_lbl.configure(
                            text=f"✓{ok_n[0]} ⚠{warn_n[0]} ✗{fail_n[0]}",
                            fg=T.GREEN if fail_n[0]==0 else (T.YELLOW if ok_n[0]>0 else T.RED))
                    if add_btn_ref and add_btn_ref[0].winfo_exists():
                        add_btn_ref[0].configure(fg=T.ACCENT,bg="#003d4d",highlightbackground=T.ACCENT)
            try: win.after(0,_upd)
            except RuntimeError: pass

        for i,url in enumerate(urls):
            def _work(ii=i,u=url): _on_result(ii, _verify_url(u))
            self._vfy_exec.submit(_work)

    # ------------------------------------------------------------------
    # Settings / History / Theme
    # ------------------------------------------------------------------
    def _open_settings(self) -> None:
        def _apply(s: Settings) -> None:
            self._settings = s
            self._semaphore = threading.Semaphore(s.max_concurrent)
            self._concurrent_var.set(s.max_concurrent)
            self._concurrent_lbl.configure(text=str(s.max_concurrent))
            if platform.system() == "Windows":
                _set_windows_startup(s.startup_with_windows)
            s.save()
        SettingsDialog(self, self._settings, _apply)

    def _open_history(self) -> None:
        HistoryPanel(self, self._db)

    def _toggle_theme(self) -> None:
        Theme.toggle()
        self._settings.dark_theme = T.is_dark()
        self._settings.save()
        # Light re-color of main window
        self.configure(bg=T.BG_DEEP)
        self._flash("נושא שונה — חלק מהצבעים יתעדכנו בפתיחה מחדש")

    # ------------------------------------------------------------------
    # Filter / list rendering
    # ------------------------------------------------------------------
    def _set_filter(self, key: str) -> None:
        self._filter_var.set(key)
        for k, tab in self._tabs.items():
            self._style_tab(tab, k)
        self._rebuild_list()

    def _rebuild_list(self) -> None:
        for card in list(self._cards.values()):
            if card.winfo_exists():
                card.destroy()
        self._cards.clear()

        filt = self._filter_var.get()
        sm = {"downloading": _ACTIVE,
              "done":       {DownloadStatus.DONE},
              "failed":     {DownloadStatus.FAILED},
              "paused":     {DownloadStatus.PAUSED},
              "queued":     {DownloadStatus.QUEUED}}

        with self._lock:
            visible = [i for i in self._items
                       if filt == "all" or i.status in sm.get(filt, {i.status})]

        for item in reversed(visible):
            card = DownloadCard(
                self._list_frame, item,
                on_pause=self._pause_item, on_resume=self._resume_item,
                on_cancel=self._cancel_item, on_retry=self._retry_item,
                on_remove=self._remove_item, on_open=self._open_item,
                on_up=self._move_up, on_down=self._move_down,
            )
            card.pack(fill="x", pady=6)
            card.refresh()
            self._cards[id(item)] = card

        self._check_empty()
        self._update_stats()

    def _check_empty(self) -> None:
        if self._cards:
            self._empty_lbl.pack_forget()
        else:
            self._empty_lbl.pack(pady=80)

    # ------------------------------------------------------------------
    # Stats / refresh
    # ------------------------------------------------------------------
    def _update_stats(self) -> None:
        with self._lock:
            items = list(self._items)
        active = sum(1 for i in items if i.status in _ACTIVE)
        done   = sum(1 for i in items if i.status == DownloadStatus.DONE)
        failed = sum(1 for i in items if i.status == DownloadStatus.FAILED)
        spd    = sum(i.speed_bps for i in items if i.status in _ACTIVE)
        total  = sum(i.downloaded_bytes for i in items)
        self._stat_labels["active"].configure(text=str(active))
        self._stat_labels["done"].configure(text=str(done))
        self._stat_labels["failed"].configure(text=str(failed))
        self._stat_labels["speed"].configure(text=_fmt_speed(spd) if spd > 0 else "—")
        if total > 0:
            self._total_lbl.configure(text=f"סה״כ הורד: {_fmt_bytes(total)}")
        # Title bar
        if active > 0 and spd > 0:
            self.title(f"{APP_NAME} v5.2 — {_fmt_speed(spd)} ▼ ({active} פעילות)")
        else:
            self.title(f"{APP_NAME} v5.2 — מנהל הורדות")
        # Overall ETA
        eta_txt = ""
        if spd > 0 and active > 0:
            remaining = sum(
                max(0, i.total_bytes - i.downloaded_bytes)
                for i in items if i.status in _ACTIVE and i.total_bytes > 0
            )
            if remaining > 0:
                eta_secs = remaining / spd
                eta_txt = f" | ETA: {_fmt_eta(eta_secs)}"
        bw = self._settings.global_bw_limit_kbps
        bw_txt = f" | מגבלה: {_fmt_speed(bw*1024)}" if bw > 0 else ""
        cb_txt = " | 📋 מנטר לוח" if self._settings.clipboard_monitor else ""
        if not hasattr(self, "_last_flash_time") or (time.monotonic() - getattr(self, "_last_flash_time", 0) > 4):
            self._status_lbl.configure(
                text=f"פעיל: {active}  |  הושלמו: {done}  |  שגיאות: {failed}{eta_txt}{bw_txt}{cb_txt}",
                fg=T.TEXT_DIM
            )

    def _schedule_refresh(self) -> None:
        with self._lock:
            snap = {id(i): i for i in self._items}
        for iid, card in list(self._cards.items()):
            item = snap.get(iid)
            if not item or not card.winfo_exists():
                continue
            if item.status in _TERMINAL:
                continue
            # Dirty check: only refresh if something actually changed
            last = getattr(item, "_last_refresh_state", None)
            cur  = (item.status, round(item.progress, 1), item.downloaded_bytes,
                    round(item.speed_bps, -3))  # round speed to 1KB to avoid noise
            if cur != last:
                item._last_refresh_state = cur  # type: ignore[attr-defined]
                card.refresh()
        self._update_stats()
        self.after(POLL_MS, self._schedule_refresh)

    def _flash(self, msg: str) -> None:
        if self._flash_id:
            try:
                self.after_cancel(self._flash_id)
            except Exception:
                pass
        self._last_flash_time = time.monotonic()
        self._status_lbl.configure(text=msg, fg=T.ACCENT)
        self._flash_id = self.after(4000, lambda: self._status_lbl.configure(text=_t("ready"), fg=T.TEXT_DIM))

    # ------------------------------------------------------------------
    # About dialog
    # ------------------------------------------------------------------
    def _open_about(self) -> None:
        win = tk.Toplevel(self)
        win.title("אודות FetchPro")
        win.configure(bg=T.BG_DEEP)
        win.resizable(False, False)
        win.grab_set()

        # Center the dialog
        win.update_idletasks()
        w, h = 480, 340
        x = self.winfo_x() + (self.winfo_width()  - w) // 2
        y = self.winfo_y() + (self.winfo_height() - h) // 2
        win.geometry(f"{w}x{h}+{x}+{y}")

        # Icon row
        icon_row = tk.Frame(win, bg=T.BG_DEEP, pady=18)
        icon_row.pack(fill="x")
        tk.Label(icon_row, text="⬇", bg=T.BG_DEEP, fg=T.ACCENT,
                 font=("Segoe UI Emoji", 36)).pack()
        tk.Label(icon_row, text="FETCHPRO", bg=T.BG_DEEP, fg=T.TEXT_MAIN,
                 font=("Consolas", 20, "bold")).pack()
        tk.Label(icon_row, text="v5.3 — מנהל הורדות מקצועי", bg=T.BG_DEEP, fg=T.TEXT_DIM,
                 font=("Consolas", 9)).pack()

        tk.Frame(win, bg=T.BORDER, height=1).pack(fill="x", padx=20)

        # License block
        lic = tk.Frame(win, bg=T.BG_CARD, highlightthickness=1,
                       highlightbackground=T.BORDER, padx=24, pady=18)
        lic.pack(fill="x", padx=20, pady=16)

        tk.Label(lic, text="© כל הזכויות שמורות למשה פינחסי",
                 bg=T.BG_CARD, fg=T.ACCENT,
                 font=("Consolas", 11, "bold"), justify="center").pack()
        tk.Label(lic,
                 text="התוכנה מוצעת לשימוש חופשי לכלל הציבור.\n"
                      "All rights reserved to Moshe Pinhasi.\n"
                      "This software may be used freely.",
                 bg=T.BG_CARD, fg=T.TEXT_MAIN,
                 font=("Consolas", 9), justify="center").pack(pady=(8, 0))

        tk.Frame(win, bg=T.BG_DEEP).pack(expand=True)

        # Close button
        close_btn = tk.Label(win, text=_t("close"), cursor="hand2",
                             bg="#003d4d", fg=T.ACCENT,
                             font=("Consolas", 10, "bold"), padx=20, pady=6,
                             highlightthickness=1, highlightbackground=T.ACCENT)
        close_btn.bind("<Button-1>", lambda _: win.destroy())
        close_btn.pack(pady=(0, 16))

    # ------------------------------------------------------------------
    # Chrome Extension Bridge entry-point
    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Persistent Queue
    # ------------------------------------------------------------------
    def _restore_queue(self) -> None:
        """Restore pending downloads from previous session."""
        items = self._pq.load()
        if not items:
            return
        count = 0
        for item in items:
            if not item.url:
                continue
            self._enqueue_item(item)
            count += 1
        if count:
            self._flash(f"✓ שוחזרו {count} הורדות מהסשן הקודם")

    def _save_queue(self) -> None:
        """Save current queue to disk."""
        if self._settings.persistent_queue:
            with self._lock:
                items = list(self._items)
            self._pq.save(items)

    # ------------------------------------------------------------------
    # Shutdown after queue
    # ------------------------------------------------------------------
    def _shutdown_monitor(self) -> None:
        """Wait for all downloads to complete, then shut down."""
        # Give the app time to start
        time.sleep(10)
        while True:
            with self._lock:
                active = [i for i in self._items if i.status in _ACTIVE or
                          i.status in (DownloadStatus.PENDING, DownloadStatus.QUEUED)]
            if not active and self._items:
                # All done — trigger shutdown
                self.after(0, self._do_shutdown_action)
                return
            time.sleep(5)

    def _do_shutdown_action(self) -> None:
        action = self._settings.shutdown_action
        self._flash(f"✓ כל ההורדות הושלמו — {action} בעוד 30 שניות...")
        self.after(30_000, lambda: self._execute_shutdown(action))

    def _execute_shutdown(self, action: str) -> None:
        self._save_queue()
        if platform.system() == "Windows":
            cmds = {
                "shutdown":  ["shutdown", "/s", "/t", "0"],
                "hibernate": ["shutdown", "/h"],
                "sleep":     ["rundll32", "powrprof.dll,SetSuspendState", "0,1,0"],
            }
        else:
            cmds = {
                "shutdown":  ["sudo", "shutdown", "-h", "now"],
                "hibernate": ["systemctl", "hibernate"],
                "sleep":     ["systemctl", "suspend"],
            }
        cmd = cmds.get(action, cmds["sleep"])
        try:
            subprocess.Popen(cmd)
        except Exception as e:
            logger.error("Shutdown failed: %s", e)
            messagebox.showerror("שגיאת כיבוי", str(e), parent=self)

    # ------------------------------------------------------------------
    # REST API helpers
    # ------------------------------------------------------------------
    def _add_url_from_rest(self, url: str, fmt: str = "", audio: bool = False,
                           playlist: bool = False, tags: str = "") -> None:
        save_dir = Path(self._settings.save_dir)
        item = DownloadItem(url=url, save_dir=save_dir,
                            media_format=fmt, media_is_audio=audio,
                            media_playlist=playlist, tags=tags)
        self._enqueue_item(item)
        self.deiconify()
        self.lift()

    def _pause_all(self) -> None:
        with self._lock:
            targets = [i for i in self._items if i.status == DownloadStatus.DOWNLOADING]
        for item in targets:
            item.pause()
            self._refresh_card(item)

    def _resume_all(self) -> None:
        with self._lock:
            targets = [i for i in self._items if i.status == DownloadStatus.PAUSED]
        for item in targets:
            item.resume()
            self._refresh_card(item)

    # ------------------------------------------------------------------
    # Stats panel
    # ------------------------------------------------------------------
    def _open_language_dialog(self) -> None:
        """Quick language picker — applies immediately, no restart needed."""
        d = tk.Toplevel(self)
        d.title("🌍 Language / שפה")
        d.configure(bg=T.BG_DEEP)
        d.geometry("300x310")
        d.resizable(False, False)
        d.grab_set()

        tk.Label(d, text=_t("choose_lang"),
                 bg=T.BG_DEEP, fg=T.TEXT_MAIN,
                 font=("Consolas", 11, "bold"), pady=12).pack()
        tk.Frame(d, bg=T.BORDER, height=1).pack(fill="x")

        body = tk.Frame(d, bg=T.BG_DEEP, padx=24, pady=12)
        body.pack(fill="both", expand=True)

        selected = tk.StringVar(value=self._settings.language)

        lang_icons = {"he": "🇮🇱", "en": "🇺🇸", "ar": "🇸🇦",
                      "ru": "🇷🇺", "es": "🇪🇸", "fr": "🇫🇷"}

        for code, name in SUPPORTED_LANGS.items():
            icon = lang_icons.get(code, "🌐")
            row = tk.Frame(body, bg=T.BG_DEEP)
            row.pack(fill="x", pady=2)
            rb = tk.Radiobutton(
                row, text=f"  {icon}  {name}",
                variable=selected, value=code,
                bg=T.BG_DEEP, fg=T.TEXT_MAIN,
                selectcolor=T.BG_CARD, activebackground=T.BG_DEEP,
                font=("Consolas", 10), cursor="hand2",
            )
            rb.pack(anchor="w")

        tk.Frame(d, bg=T.BORDER, height=1).pack(fill="x")
        btn_row = tk.Frame(d, bg=T.BG_DEEP, pady=8)
        btn_row.pack(fill="x")

        def _apply() -> None:
            code = selected.get()
            global _LANG
            _LANG = code
            self._settings.language = code
            self._settings.save()
            # ── Update toolbar labels immediately ────────────────────
            self._refresh_toolbar_labels()
            d.destroy()

        save_btn = tk.Label(btn_row, text=_t("apply"), cursor="hand2",
                            bg="#003d4d", fg=T.ACCENT,
                            font=("Consolas", 11, "bold"),
                            padx=20, pady=5,
                            highlightthickness=1, highlightbackground=T.ACCENT)
        save_btn.bind("<Button-1>", lambda _: _apply())
        save_btn.pack(side="right", padx=12)

        cancel_btn = tk.Label(btn_row, text=_t("cancel_btn"), cursor="hand2",
                              bg=T.BG_DEEP, fg=T.TEXT_DIM,
                              font=("Consolas", 9), padx=12, pady=5,
                              highlightthickness=1, highlightbackground=T.BORDER)
        cancel_btn.bind("<Button-1>", lambda _: d.destroy())
        cancel_btn.pack(side="right")

    def _refresh_toolbar_labels(self) -> None:
        """Update all translatable toolbar labels to the current language."""
        updates = {
            "_tb_settings": _t("settings"),
            "_tb_media":    _t("media"),
            "_tb_retry":    _t("retry_all"),
            "_tb_theme":    _t("theme"),
            "_tb_lang":     "🌍 " + SUPPORTED_LANGS.get(_LANG, "שפה"),
            "_tb_more":     "⋯ " + _t("settings").split()[0][:2] + "..." if False else "⋯ עוד",
            "_tb_quit":     _t("quit"),
        }
        for attr, text in updates.items():
            widget = getattr(self, attr, None)
            if widget and widget.winfo_exists():
                widget.configure(text=text)
        # Update title
        self.title(f"{APP_NAME} v5.5 — {_t('app_title')}")

    def _open_stats(self) -> None:
        """Show usage statistics window."""
        d = tk.Toplevel(self)
        d.title("📊 סטטיסטיקות שימוש")
        d.configure(bg=T.BG_DEEP)
        d.geometry("480x340")
        d.grab_set()

        tk.Label(d, text=_t("stats_title_lbl"), bg=T.BG_DEEP, fg=T.TEXT_MAIN,
                 font=("Consolas", 13, "bold"), pady=14).pack()
        tk.Frame(d, bg=T.BORDER, height=1).pack(fill="x")

        stats = _STATS.data
        body  = tk.Frame(d, bg=T.BG_DEEP, padx=30, pady=16)
        body.pack(fill="both", expand=True)

        rows = [
            ("📦 סה\"כ הורדות",    f"{stats.get('total_files', 0):,} קבצים"),
            ("💾 סה\"כ נפח",       _fmt_bytes(stats.get("total_bytes", 0))),
            ("⚡ שיא מהירות",      _fmt_speed(stats.get("fastest_bps", 0))),
            ("🔄 סשנים",           f"{stats.get('total_sessions', 0):,}"),
            ("", ""),
            ("📥 הסשן הנוכחי",    f"{stats.get('session_files', 0):,} קבצים"),
            ("📊 נפח הסשן",        _fmt_bytes(stats.get("session_bytes", 0))),
        ]
        for label, val in rows:
            if not label:
                tk.Frame(body, bg=T.BORDER, height=1).pack(fill="x", pady=6)
                continue
            row = tk.Frame(body, bg=T.BG_DEEP)
            row.pack(fill="x", pady=3)
            tk.Label(row, text=label, bg=T.BG_DEEP, fg=T.TEXT_DIM,
                     font=("Consolas", 10), width=22, anchor="w").pack(side="left")
            tk.Label(row, text=val, bg=T.BG_DEEP, fg=T.ACCENT,
                     font=("Consolas", 10, "bold")).pack(side="left")

        # Active downloads right now
        with self._lock:
            active = sum(1 for i in self._items if i.status in _ACTIVE)
            total_speed = sum(i.speed_bps for i in self._items if i.status in _ACTIVE)
        row = tk.Frame(body, bg=T.BG_DEEP)
        row.pack(fill="x", pady=3)
        tk.Label(row, text=_t("active_now"), bg=T.BG_DEEP, fg=T.TEXT_DIM,
                 font=("Consolas", 10), width=22, anchor="w").pack(side="left")
        tk.Label(row, text=f"{active} הורדות  |  {_fmt_speed(total_speed)}",
                 bg=T.BG_DEEP, fg=T.GREEN, font=("Consolas", 10, "bold")).pack(side="left")

        tk.Button(d, text=_t("close"), command=d.destroy,
                  bg=T.BG_CARD, fg=T.TEXT_MAIN, font=("Consolas", 10),
                  relief="flat", padx=20, pady=6, cursor="hand2").pack(pady=12)

    # ------------------------------------------------------------------
    # YouTube / Media download dialog
    # ------------------------------------------------------------------
    def _open_media_dialog(self) -> None:
        """Open dedicated YouTube/media download dialog with format selection."""
        urls = self._read_urls()
        url  = urls[0] if urls else ""

        d = tk.Toplevel(self)
        d.title("🎬 הורדת מדיה — YouTube / מוזיקה")
        d.configure(bg=T.BG_DEEP)
        d.geometry("560x520")
        d.grab_set()

        # Header
        hdr = tk.Frame(d, bg="#0a1520", pady=14, padx=20)
        hdr.pack(fill="x")
        tk.Label(hdr, text=_t("media_dialog_hdr"),
                 bg="#0a1520", fg=T.TEXT_MAIN,
                 font=("Consolas", 12, "bold")).pack(anchor="w")
        tk.Label(hdr, text=_t("media_sites_list"),
                 bg="#0a1520", fg=T.TEXT_DIM, font=("Consolas", 8)).pack(anchor="w")
        tk.Frame(d, bg=T.BORDER, height=1).pack(fill="x")

        body = tk.Frame(d, bg=T.BG_DEEP, padx=24, pady=14)
        body.pack(fill="both", expand=True)

        def lbl(text):
            tk.Label(body, text=text, bg=T.BG_DEEP, fg=T.TEXT_DIM,
                     font=("Consolas", 9), anchor="w").pack(fill="x", pady=(10,2))

        # URL
        lbl(_t("media_url"))
        url_var = tk.StringVar(value=url)
        url_entry = tk.Entry(body, textvariable=url_var, bg=T.BG_CARD, fg=T.TEXT_MAIN,
                             insertbackground=T.ACCENT, font=("Consolas", 10),
                             relief="flat")
        url_entry.pack(fill="x")

        # Right-click context menu on URL entry
        url_ctx = tk.Menu(url_entry, tearoff=0,
                          bg=T.BG_CARD, fg=T.TEXT_MAIN,
                          activebackground=T.BG_HOVER, activeforeground=T.ACCENT,
                          font=("Consolas", 9), bd=0)

        def _paste_clipboard():
            try:
                clip = self.clipboard_get()
                if clip:
                    url_var.set(clip)
                    url_entry.icursor("end")
            except Exception:
                pass

        url_ctx.add_command(label=_t("paste"),          command=lambda: (url_entry.event_generate("<<Paste>>"), None))
        url_ctx.add_command(label=_t("cut"),             command=lambda: url_entry.event_generate("<<Cut>>"))
        url_ctx.add_command(label=_t("copy_action"),     command=lambda: url_entry.event_generate("<<Copy>>"))
        url_ctx.add_separator()
        url_ctx.add_command(label=_t("paste_from_clip"), command=_paste_clipboard)
        url_ctx.add_command(label=_t("select_all"),      command=lambda: url_entry.select_range(0, "end"))
        url_ctx.add_command(label=_t("clear"),           command=lambda: url_var.set(""))

        def _show_url_ctx(event):
            url_entry.focus_set()
            try:
                url_ctx.tk_popup(event.x_root, event.y_root)
            finally:
                url_ctx.grab_release()

        url_entry.bind("<Button-3>",  _show_url_ctx)
        url_entry.bind("<Button-2>",  _show_url_ctx)   # macOS
        url_entry.bind("<Control-a>", lambda _: (url_entry.select_range(0, "end"), "break"))
        url_entry.focus_set()

        # Mode: Video / Audio
        lbl(_t("media_type"))
        mode_var = tk.StringVar(value="video")
        mode_frame = tk.Frame(body, bg=T.BG_DEEP)
        mode_frame.pack(fill="x")
        for val, label in [("video", _t("media_video")), ("audio", _t("media_audio_only"))]:
            tk.Radiobutton(mode_frame, text=label, variable=mode_var, value=val,
                           bg=T.BG_DEEP, fg=T.TEXT_MAIN, selectcolor=T.BG_CARD,
                           activebackground=T.BG_DEEP, font=("Consolas", 10),
                           cursor="hand2").pack(side="left", padx=(0, 16))

        # Video quality
        lbl(_t("media_quality"))
        quality_var = tk.StringVar(value="best")
        quality_frame = tk.Frame(body, bg=T.BG_DEEP)
        quality_frame.pack(fill="x")
        for val, label in [("best", _t("best_quality")), ("4k","4K"), ("1080p","1080p"),
                            ("720p","720p"), ("480p","480p"), ("360p","360p")]:
            tk.Radiobutton(quality_frame, text=label, variable=quality_var, value=val,
                           bg=T.BG_DEEP, fg=T.TEXT_MAIN, selectcolor=T.BG_CARD,
                           activebackground=T.BG_DEEP, font=("Consolas", 9),
                           cursor="hand2").pack(side="left", padx=(0, 8))

        # Audio format
        lbl(_t("media_audio_format"))
        afmt_var = tk.StringVar(value="mp3")
        afmt_frame = tk.Frame(body, bg=T.BG_DEEP)
        afmt_frame.pack(fill="x")
        for val in ["mp3", "m4a", "opus", "flac", "wav"]:
            tk.Radiobutton(afmt_frame, text=val.upper(), variable=afmt_var, value=val,
                           bg=T.BG_DEEP, fg=T.TEXT_MAIN, selectcolor=T.BG_CARD,
                           activebackground=T.BG_DEEP, font=("Consolas", 9),
                           cursor="hand2").pack(side="left", padx=(0, 8))

        # Playlist
        lbl(_t("media_options"))
        opts_frame = tk.Frame(body, bg=T.BG_DEEP)
        opts_frame.pack(fill="x")
        playlist_var    = tk.BooleanVar(value=False)
        thumbnail_var   = tk.BooleanVar(value=True)
        metadata_var    = tk.BooleanVar(value=True)
        for var, text in [(playlist_var, _t("media_playlist")),
                          (thumbnail_var, _t("media_thumbnail")),
                          (metadata_var,  _t("media_metadata"))]:
            tk.Checkbutton(opts_frame, text=text, variable=var,
                           bg=T.BG_DEEP, fg=T.TEXT_MAIN, selectcolor=T.BG_CARD,
                           activebackground=T.BG_DEEP, font=("Consolas", 9),
                           cursor="hand2").pack(side="left", padx=(0, 14))

        # Tags
        lbl(_t("media_tags"))
        tags_var = tk.StringVar()
        tk.Entry(body, textvariable=tags_var, bg=T.BG_CARD, fg=T.TEXT_MAIN,
                 insertbackground=T.ACCENT, font=("Consolas", 9),
                 relief="flat").pack(fill="x")

        if not YTDLP_OK:
            tk.Label(body, text=_t("no_ytdlp"),
                     bg=T.BG_DEEP, fg=T.RED, font=("Consolas", 9)).pack(anchor="w", pady=4)

        tk.Frame(d, bg=T.BORDER, height=1).pack(fill="x")
        btn_row = tk.Frame(d, bg=T.BG_DEEP, pady=12)
        btn_row.pack(fill="x")

        def _submit() -> None:
            u = url_var.get().strip()
            if not u:
                messagebox.showwarning(_t("url_missing_title"), _t("no_url_warning"), parent=d)
                return
            is_audio = mode_var.get() == "audio"
            fmt = afmt_var.get() if is_audio else quality_var.get()
            # Patch settings for this download
            old_thumb = self._settings.ytdlp_embed_thumbnail
            old_meta  = self._settings.ytdlp_add_metadata
            self._settings.ytdlp_embed_thumbnail = thumbnail_var.get()
            self._settings.ytdlp_add_metadata    = metadata_var.get()
            item = DownloadItem(
                url           = u,
                save_dir      = Path(self._settings.save_dir),
                media_format  = fmt,
                media_is_audio= is_audio,
                media_playlist= playlist_var.get(),
                tags          = tags_var.get().strip(),
            )
            self._enqueue_item(item)
            # Restore
            self._settings.ytdlp_embed_thumbnail = old_thumb
            self._settings.ytdlp_add_metadata    = old_meta
            d.destroy()
            self._url_text.delete("1.0", "end")
            self._url_focus_out(None)
            self._flash(f"📥 נוסף: {u[:60]}")

        tk.Label(btn_row, text="", bg=T.BG_DEEP).pack(side="left", padx=10, expand=True)
        cancel_btn = tk.Label(btn_row, text=_t("cancel_btn"), cursor="hand2",
                              bg=T.BG_DEEP, fg=T.TEXT_DIM, font=("Consolas", 10),
                              padx=16, pady=6, highlightthickness=1,
                              highlightbackground=T.BORDER)
        cancel_btn.bind("<Button-1>", lambda _: d.destroy())
        cancel_btn.pack(side="right", padx=(0, 10))

        add_btn = tk.Label(btn_row, text=_t("add_to_download"), cursor="hand2",
                           bg="#003d4d", fg=T.ACCENT, font=("Consolas", 11, "bold"),
                           padx=18, pady=6, highlightthickness=1,
                           highlightbackground=T.ACCENT)
        add_btn.bind("<Button-1>", lambda _: _submit())
        add_btn.pack(side="right", padx=(0, 10))

    def _add_url_from_bridge(self, url: str) -> None:
        """Called from the bridge HTTP server (main-thread via after())."""
        self._url_text.delete("1.0", "end")
        self._url_text.configure(fg=T.TEXT_MAIN)
        self._url_text.insert("1.0", url)
        self._add_downloads()
        self.deiconify()
        self.lift()
        self._flash(f"📥 התקבל מהדפדפן: {url[:60]}…" if len(url) > 60 else f"📥 התקבל מהדפדפן: {url}")

    # ------------------------------------------------------------------
    # Shutdown
    # ------------------------------------------------------------------
    def _on_close(self) -> None:
        """X button — hide to tray if enabled, otherwise quit."""
        if self._settings.system_tray:
            # Always just hide — works even without pystray installed
            self.withdraw()
        else:
            self._quit_app()

    def _quit_app(self) -> None:
        """Full exit — called from toolbar Quit button or tray menu."""
        self.protocol("WM_DELETE_WINDOW", lambda: None)
        try:
            self._settings.window_geometry = self.geometry()
        except Exception:
            pass
        self._settings.save()
        self._save_queue()   # persist queue before exit
        with self._lock:
            targets = list(self._items)
        for item in targets:
            item.cancel()
        self._executor.shutdown(wait=False, cancel_futures=True)
        self._vfy_exec.shutdown(wait=False, cancel_futures=True)
        self._db.close()
        if self._tray:
            try:
                self._tray.stop()
            except Exception:
                pass
        self.destroy()


# ---------------------------------------------------------------------------
# Windows auto-start (Startup folder — no admin rights needed)
# ---------------------------------------------------------------------------

def _get_startup_folder() -> "Path | None":
    """Return the Windows per-user Startup folder path."""
    if platform.system() != "Windows":
        return None
    try:
        folder = Path(os.environ.get("APPDATA", "")) / \
                 "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
        if folder.exists():
            return folder
        # fallback: ask shell
        import subprocess
        result = subprocess.run(
            ["powershell", "-Command",
             "[Environment]::GetFolderPath('Startup')"],
            capture_output=True, text=True, timeout=5
        )
        p = Path(result.stdout.strip())
        return p if p.exists() else None
    except Exception:
        return None


def _get_startup_bat() -> "Path | None":
    folder = _get_startup_folder()
    return (folder / "FetchPro.bat") if folder else None


def _build_startup_cmd() -> str:
    """Build the startup command, works for both .py and .exe."""
    if getattr(sys, "frozen", False):
        # PyInstaller EXE
        return f'start "" "{EXE_PATH}" --minimized'
    else:
        py = Path(sys.executable).resolve()
        pythonw = py.parent / "pythonw.exe"
        interpreter = pythonw if pythonw.exists() else py
        script = Path(sys.argv[0]).resolve()
        return f'start "" "{interpreter}" "{script}" --minimized'


def _set_windows_startup(enable: bool) -> tuple[bool, str]:
    """
    Add or remove FetchPro from Windows Startup folder.
    Returns (success: bool, message: str).
    """
    if platform.system() != "Windows":
        return False, "לא Windows"

    bat = _get_startup_bat()
    if bat is None:
        return False, "לא נמצאה תיקיית Startup"

    try:
        if enable:
            cmd = _build_startup_cmd()
            bat.write_text(
                f"@echo off\r\n{cmd}\r\n",
                encoding="utf-8"
            )
            logger.info("Startup bat written: %s → %s", bat, cmd)
            return True, f"נוסף: {bat}"
        else:
            if bat.exists():
                bat.unlink()
                logger.info("Startup bat removed: %s", bat)
            return True, "הוסר"
    except Exception as exc:
        logger.warning("Startup folder error: %s", exc)
        return False, str(exc)


def _is_windows_startup_enabled() -> bool:
    """Check whether the FetchPro startup bat exists."""
    if platform.system() != "Windows":
        return False
    bat = _get_startup_bat()
    return bat is not None and bat.exists()


# ---------------------------------------------------------------------------
# Chrome Extension Bridge — local HTTP server on BRIDGE_PORT
# ---------------------------------------------------------------------------
class _BridgeHandler(http.server.BaseHTTPRequestHandler):
    """Minimal HTTP handler that lets the Chrome extension push URLs to FetchPro."""
    app: "FetchProApp | None" = None   # set after app is created

    # --- CORS pre-flight ---
    def do_OPTIONS(self) -> None:
        self.send_response(200)
        self._cors()
        self.end_headers()

    # --- status probe ---
    def do_GET(self) -> None:
        if self.path == "/status":
            self._reply(200, {"running": True, "app": APP_NAME, "version": "5.3"})
        else:
            self._reply(404, {"error": "not found"})

    # --- add URL ---
    def do_POST(self) -> None:
        if self.path != "/add":
            self._reply(404, {"error": "not found"})
            return
        try:
            length = int(self.headers.get("Content-Length", 0))
            body   = self.rfile.read(length)
            data   = json.loads(body)
            url    = (data.get("url") or "").strip()
            if not url:
                self._reply(400, {"error": "missing url"}); return
            if _BridgeHandler.app:
                _BridgeHandler.app.after(0,
                    lambda u=url: _BridgeHandler.app._add_url_from_bridge(u))
            self._reply(200, {"ok": True, "url": url})
        except Exception as exc:
            self._reply(400, {"error": str(exc)})

    def _reply(self, code: int, body: dict) -> None:
        payload = json.dumps(body).encode()
        self.send_response(code)
        self._cors()
        self.send_header("Content-Type",   "application/json")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def _cors(self) -> None:
        origin = self.headers.get("Origin", "")
        # Allow localhost and Chrome extension origins only
        if (origin.startswith("http://localhost") or
                origin.startswith("http://127.0.0.1") or
                origin.startswith("chrome-extension://")):
            allowed = origin
        else:
            allowed = "http://127.0.0.1"
        self.send_header("Access-Control-Allow-Origin",  allowed)
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def log_message(self, *_) -> None:   # silence console spam
        pass


def _start_bridge(app: "FetchProApp") -> None:
    """Start the bridge HTTP server in a daemon thread."""
    _BridgeHandler.app = app
    try:
        server = http.server.HTTPServer(("127.0.0.1", BRIDGE_PORT), _BridgeHandler)
        logger.info("Chrome bridge listening on port %d", BRIDGE_PORT)
        server.serve_forever()
    except OSError as exc:
        logger.warning("Bridge server could not start: %s", exc)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main() -> None:
    APP_DIR.mkdir(parents=True, exist_ok=True)
    app = FetchProApp()

    # If launched with --minimized (e.g. from Windows startup), hide to tray
    if "--minimized" in sys.argv:
        app.withdraw()

    app.mainloop()


if __name__ == "__main__":
    main()
