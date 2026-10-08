import json
import os
import re
import sys
import shutil
import hashlib
import subprocess
import unicodedata
from datetime import date, datetime, time, timedelta

from PyQt6.QtCore import QDate, Qt, QThread, QTimer, pyqtSignal
from PyQt6.QtGui import QKeySequence, QShortcut
from PyQt6.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QComboBox, QDateEdit, QFileDialog,
    QFormLayout, QGridLayout, QGroupBox, QHBoxLayout, QLabel, QLineEdit,
    QMainWindow, QMessageBox, QPlainTextEdit, QPushButton, QTreeWidget, QTreeWidgetItem,
    QVBoxLayout, QWidget,
)

DATA_FILE = './data.json'
IMG_PATH = './img/'
REPO_DIR = '.'

# Az oldal szekciói (a data.json kulcsa, felirat a kezelőben, fajta, ennyi látszik az oldalon)
LIST_SECTIONS = [
    ('kisokos',  'AI Kisokos (narancs sáv)',          'tile',    4),
    ('promptok', 'Top promptok (sárga sáv)',          'prompt',  4),
    ('kekSav',   'AI tudástár (kék sáv)',             'tile',    4),
    ('promo',    'Étlap + AI eszközök (2 nagy kép)',  'promo',   2),
    ('usecases', 'Sikertörténetek (L / M / S)',       'usecase', None),
]
EVENTS = 'esemenyek'     # a naptár — csak Excelből töltődik, kézzel nem szerkeszthető

# A régi képes útmutató-csempék. Az új oldal már nem mutatja őket (a sikertörténetek
# váltották ki), de az adatuk megmarad a data.json-ban. True-ra állítva újra szerkeszthetők.
SHOW_GUIDE_CARDS = False
SECTIONS = ['altalanos', 'copilot']
CARD_TYPES = ['largeCards', 'smallCards']
LEVELS = ['kezdo', 'halado']
SIZES = [('L', 'L — nagy projekt'), ('M', 'M — közepes projekt'), ('S', 'S — kicsi projekt')]

LABEL_OF = {key: label for key, label, _, _ in LIST_SECTIONS}
LABEL_OF.update({'altalanos': 'Általános útmutatók (régi)', 'copilot': 'Copilot útmutatók (régi)'})
KIND_OF = {key: kind for key, _, kind, _ in LIST_SECTIONS}
KIND_OF.update({'altalanos': 'card', 'copilot': 'card'})
LIMIT_OF = {key: limit for key, _, _, limit in LIST_SECTIONS}

# Melyik fajtánál melyik mező látszik az űrlapon
FIELDS = {
    'card':    {'type', 'level', 'title', 'img', 'link', 'date'},
    'tile':    {'title', 'sub', 'img', 'link', 'date'},
    'prompt':  {'title', 'text'},
    'promo':   {'title', 'img', 'link'},
    'usecase': {'title', 'sub', 'size', 'img', 'link'},
}
IMG_REQUIRED = {'card', 'promo'}
MANAGED_KEYS = {'title', 'sub', 'text', 'img', 'link', 'date', 'level', 'mark', 'size', 'isNew'}

SETTINGS_FILE = './manager_beallitasok.json'   # helyi beállítás (pl. a naptár Excel útvonala) — nem kerül GitHubra

ALLOWED_EXT = {'.png', '.jpg', '.jpeg', '.gif', '.webp', '.svg'}
MAX_IMG_BYTES = 5 * 1024 * 1024  # 5 MB — csak figyelmeztetés, nem tiltás
PUSH_DELAY_MS = 4000  # ennyi nyugalom után indul az automatikus feltöltés
NEW_DAYS = 30  # ennyi napig 'ÚJ' egy kártya — az index.html-ben is ugyanennyi legyen
NO_DATE = QDate(2000, 1, 1)  # a dátummező 'nincs dátum' állása

# Mi történjen, ha a data.json a gépen ÉS a GitHubon is változott?
#   'newer'  -> amelyik frissebb (helyi fájl módosítási ideje vs. GitHub commit ideje)
#   'local'  -> mindig a gépen lévő marad
#   'remote' -> mindig a GitHubos marad
CONFLICT_POLICY = 'remote'
# Ütközésnél a VESZTES oldal mindig ide mentődik, így semmi nem vész el.
# (Nem kerül fel GitHubra, mert a szinkron csak a data.json-t és az img-t tölti fel.)
BACKUP_DIR = './backup/'

# Világos és sötét témán is olvasható színek. A None a téma alap szövegszínét jelenti.
STATUS_NEUTRAL = None
STATUS_PENDING = '#c9911a'
STATUS_OK = '#3aa14b'
STATUS_ERROR = '#e5534b'

ROLE = Qt.ItemDataRole.UserRole


# --- Segédfüggvények -------------------------------------------------------

def sanitize_filename(name):
    """'Képernyőkép 2026-08-30.PNG' -> 'kepernyokep-2026-08-30.png'"""
    base, ext = os.path.splitext(name)

    ext = ext.lower()
    if ext == '.jpeg':
        ext = '.jpg'

    # ékezetek leszedése: NFKD szétbontja az ő-t o + jelre, az ascii ignore eldobja a jelet
    base = unicodedata.normalize('NFKD', base)
    base = base.encode('ascii', 'ignore').decode('ascii')
    base = base.lower()
    base = re.sub(r'\s+', '-', base)
    base = re.sub(r'[^a-z0-9._-]', '', base)
    base = re.sub(r'-{2,}', '-', base).strip('-_.')

    if not base:
        base = 'kep'

    return base + ext


def parse_date(text):
    try:
        return datetime.strptime(text, '%Y-%m-%d').date()
    except (TypeError, ValueError):
        return None


def days_left(card):
    """Hány napig számít még újnak a kártya. None, ha már nem új (vagy nincs dátuma)."""
    d = parse_date(card.get('date'))
    if d is None:
        return None
    left = NEW_DAYS - (date.today() - d).days
    return left if left >= 0 else None


def migrate_cards(data):
    """A régi isNew mezőt dátumra cseréli. True, ha bármi változott.

    isNew: true  -> date = ma (innentől NEW_DAYS napig új)
    isNew: false -> a mező egyszerűen eltűnik
    """
    today = date.today().isoformat()
    changed = False

    lists = []
    for section in data.values():
        if isinstance(section, dict):
            lists.extend(v for v in section.values() if isinstance(v, list))
        elif isinstance(section, list):
            lists.append(section)  # pl. a 'kisokos' tömb

    for cards in lists:
        for card in cards:
            if not isinstance(card, dict) or 'isNew' not in card:
                continue
            if card.pop('isNew') and not card.get('date'):
                card['date'] = today
            changed = True

    return changed


def file_hash(path):
    h = hashlib.md5()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(65536), b''):
            h.update(chunk)
    return h.hexdigest()


def same_content(path_a, path_b):
    try:
        if os.path.getsize(path_a) != os.path.getsize(path_b):
            return False
        return file_hash(path_a) == file_hash(path_b)
    except OSError:
        return False


def run_git(args):
    """Git parancs futtatása. Sosem kérdez interaktívan, sosem nyit konzolablakot."""
    env = dict(os.environ)
    env['GIT_TERMINAL_PROMPT'] = '0'

    kwargs = {}
    if sys.platform == 'win32':
        kwargs['creationflags'] = subprocess.CREATE_NO_WINDOW

    return subprocess.run(
        ['git', '-C', REPO_DIR] + args,
        capture_output=True,
        text=True,
        encoding='utf-8',
        errors='replace',
        env=env,
        **kwargs
    )


# --- Naptár: Excel beolvasása ---------------------------------------------

def fold(text):
    """'Kezdés időpontja' -> 'kezdes idopontja' (ékezet nélkül, kisbetűvel)"""
    text = unicodedata.normalize('NFKD', str(text or ''))
    text = text.encode('ascii', 'ignore').decode('ascii').lower()
    return re.sub(r'[^a-z0-9]+', ' ', text).strip()


# Az oszlopok fejléce ezek bármelyikével kezdődhet (ékezet és kis/nagybetű nem számít)
EVENT_COLUMNS = {
    'date':     ['datum', 'date', 'nap'],
    'start':    ['kezdes', 'kezdete', 'start', 'tol', 'idopont', 'ido'],
    'end':      ['vege', 'befejezes', 'end', 'ig'],
    'title':    ['esemeny', 'cim', 'megnevezes', 'title', 'targy', 'tema'],
    'location': ['helyszin', 'hely', 'location', 'terem'],
    'link':     ['link', 'url', 'teams'],
}


def excel_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and 20000 < value < 80000:      # Excel sorszám
        return (datetime(1899, 12, 30) + timedelta(days=float(value))).date()
    text = str(value or '').strip()
    m = re.search(r'(\d{4})\D+(\d{1,2})\D+(\d{1,2})', text)            # 2026.10.15. / 2026-10-15
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    m = re.fullmatch(r'(\d{1,2})\s*[./-]\s*(\d{1,2})\.?', text)         # 10.15 -> idei év
    if m:
        try:
            return date(date.today().year, int(m.group(1)), int(m.group(2)))
        except ValueError:
            return None
    return None


def excel_time(value):
    if value is None or value == '':
        return ''
    if isinstance(value, datetime):
        value = value.time()
    if isinstance(value, time):
        return f'{value.hour:02d}:{value.minute:02d}'
    if isinstance(value, float) and 0 <= value < 1:                        # a nap tört része
        minutes = round(value * 24 * 60)
        return f'{minutes // 60:02d}:{minutes % 60:02d}'
    if isinstance(value, int) and 0 <= value <= 23:
        return f'{value:02d}:00'
    m = re.match(r'\s*(\d{1,2})(?:[:.](\d{2}))?', str(value))
    if m and int(m.group(1)) <= 23:
        return f'{int(m.group(1)):02d}:{int(m.group(2) or 0):02d}'
    return ''


def read_events_xlsx(path):
    """(események, kihagyott sorok listája). Az első munkalapot olvassa."""
    import openpyxl  # csak itt kell, hogy nélküle is elinduljon a kezelő

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        rows = list(wb.worksheets[0].iter_rows(values_only=True))
    finally:
        wb.close()

    # fejléc keresése az első 15 sorban
    header_row, cols = None, {}
    for r, row in enumerate(rows[:15]):
        found = {}
        for c, cell in enumerate(row or ()):
            name = fold(cell)
            for key, aliases in EVENT_COLUMNS.items():
                if key not in found and any(name == a or name.startswith(a + ' ') for a in aliases):
                    found[key] = c
                    break
        if 'date' in found and 'title' in found:
            header_row, cols = r, found
            break
    if header_row is None:
        raise ValueError('Nem találom a fejlécet. Kell legalább egy "Dátum" és egy "Esemény" (vagy "Cím") oszlop.')

    def cell(row, key):
        c = cols.get(key)
        return row[c] if c is not None and c < len(row) else None

    events, skipped = [], []
    for r, row in enumerate(rows[header_row + 1:], start=header_row + 2):
        if not row or all(v in (None, '') for v in row):
            continue
        d = excel_date(cell(row, 'date'))
        title = str(cell(row, 'title') or '').strip()
        if not d or not title:
            skipped.append(r)
            continue
        start, end = excel_time(cell(row, 'start')), excel_time(cell(row, 'end'))
        raw_start = str(cell(row, 'start') or '')
        if start and not end:                                                 # "14:00-15:00" egy cellában
            m = re.search(r'[-–]\s*(\d{1,2}[:.]\d{2})', raw_start)
            if m:
                end = excel_time(m.group(1))
        event = {'date': d.isoformat(), 'start': start, 'end': end, 'title': title}
        location = str(cell(row, 'location') or '').strip()
        link = str(cell(row, 'link') or '').strip()
        if location:
            event['location'] = location
        if link:
            event['link'] = link
        events.append(event)

    events.sort(key=lambda e: (e['date'], e['start']))
    return events, skipped


def load_settings():
    try:
        with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return {}


def save_settings(settings):
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f, indent=2, ensure_ascii=False)
    except OSError:
        pass


# --- Fa nézet, csoporton belüli húzással -----------------------------------

class CardTree(QTreeWidget):
    """Kétszintű fa: szekció/típus csoportok, alattuk a kártyák.

    A húzás csak azonos csoporton belül engedélyezett, mert a JSON-ben
    minden csoport külön lista.
    """

    orderChanged = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setColumnCount(5)
        self.setHeaderLabels(['Cím (Keresőhöz)', 'Szint / Alcím', 'Dátum', 'Kép', 'Link'])
        self.setAlternatingRowColors(True)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setUniformRowHeights(True)

        self.setColumnWidth(0, 300)
        self.setColumnWidth(1, 140)
        self.setColumnWidth(2, 170)
        self.setColumnWidth(3, 220)

    def dropEvent(self, event):
        dragged = self.currentItem()
        target = self.itemAt(event.position().toPoint())

        # Csak kártyát lehet húzni, csoportot nem
        if dragged is None or dragged.parent() is None:
            event.ignore()
            return

        if target is None:
            event.ignore()
            return

        target_group = target.parent() if target.parent() is not None else target
        if target_group is not dragged.parent():
            event.ignore()
            return

        # Kártyára ejtés helyett mindig sorok közé kerüljön
        position = self.dropIndicatorPosition()
        if (position == QAbstractItemView.DropIndicatorPosition.OnItem
                and target.parent() is not None):
            event.ignore()
            return

        super().dropEvent(event)
        self.orderChanged.emit()


# --- GitHub szinkron háttérszálon ------------------------------------------

def read_bytes(path):
    try:
        with open(path, 'rb') as f:
            return f.read()
    except OSError:
        return None


class GitSyncWorker(QThread):
    """Előbb behúzza a GitHub állapotát, eldönti ki nyer a data.json-nál,
    aztán feltölti, ami helyben új."""

    done = pyqtSignal(str, str, str, bool)  # státusz, szín, hibaüzenet, változott-e a data.json

    def run(self):
        self.start_bytes = read_bytes(DATA_FILE)
        try:
            self.sync()
        except FileNotFoundError:
            self.finish('Nincs git.', STATUS_ERROR,
                        'A git parancs nem található. Telepítsd a Git for Windows csomagot.')
        except Exception as error:  # ne haljon el csendben a szál
            self.finish('Váratlan hiba.', STATUS_ERROR, repr(error))

    def finish(self, status, color, detail=''):
        changed = read_bytes(DATA_FILE) != self.start_bytes
        self.done.emit(status, color or '', detail, changed)

    def git(self, args):
        r = run_git(args)
        return r.returncode, (r.stdout or '').strip(), (r.stderr or r.stdout or '').strip()

    def sync(self):
        rc, _, _ = self.git(['rev-parse', '--is-inside-work-tree'])
        if rc != 0:
            return self.finish('Nem git repo.', STATUS_ERROR,
                               f'A(z) {os.path.abspath(REPO_DIR)} mappa nem git repository.')

        rc, upstream, _ = self.git(['rev-parse', '--abbrev-ref', '--symbolic-full-name', '@{u}'])
        if rc != 0:
            return self.finish('Nincs követett ág.', STATUS_ERROR,
                               'Az aktuális ágnak nincs beállítva távoli párja.\n\n'
                               "Futtasd egyszer kézzel a mappában:  git push -u origin main")

        # 1. Távoli állapot letöltése (még nem nyúl a helyi fájlokhoz)
        rc, _, err = self.git(['fetch', '--quiet'])
        if rc != 0:
            return self.finish('Letöltés sikertelen.', STATUS_ERROR,
                               'Nem sikerült elérni a GitHubot:\n\n' + err)

        # A helyi data.json mentése, mielőtt bármi hozzányúlna
        local_bytes = read_bytes(DATA_FILE)
        local_mtime = os.path.getmtime(DATA_FILE) if local_bytes is not None else 0

        # 2. Helyi változások commitolása
        paths = [p for p in ('data.json', 'img') if os.path.exists(p)]
        if paths:
            rc, _, err = self.git(['add', '--'] + paths)
            if rc != 0:
                return self.finish('Add sikertelen.', STATUS_ERROR, err)

        stamp = datetime.now().strftime('%Y-%m-%d %H:%M')
        rc, _, _ = self.git(['diff', '--cached', '--quiet'])
        if rc == 1:
            rc, _, err = self.git(['commit', '-m', f'Kártyák frissítése — {stamp}'])
            if rc != 0:
                return self.finish('Commit sikertelen.', STATUS_ERROR, err)

        # 3. Ki változtatta a data.json-t a közös pont óta?
        rc, base, err = self.git(['merge-base', 'HEAD', upstream])
        if rc != 0:
            return self.finish('Nincs közös előzmény.', STATUS_ERROR,
                               'A helyi és a GitHubos repo előzménye nem kapcsolódik:\n\n' + err)

        local_changed = self.git(['diff', '--quiet', base, 'HEAD', '--', 'data.json'])[0] == 1
        remote_changed = self.git(['diff', '--quiet', base, upstream, '--', 'data.json'])[0] == 1

        conflict = local_changed and remote_changed
        local_wins = True
        if conflict:
            if CONFLICT_POLICY == 'local':
                local_wins = True
            elif CONFLICT_POLICY == 'remote':
                local_wins = False
            else:
                _, ct, _ = self.git(['log', '-1', '--format=%ct', upstream, '--', 'data.json'])
                remote_time = int(ct) if ct.isdigit() else 0
                local_wins = local_mtime >= remote_time

        # 4. Rebase a GitHubos állapotra. Rebase közben az "ours" a távoli,
        #    a "theirs" a helyi commit — ezért fordítva kell megadni.
        behind = self.git(['rev-list', '--count', f'HEAD..{upstream}'])[1]
        if behind not in ('', '0'):
            strategy = 'theirs' if local_wins else 'ours'
            # --autostash: a nem általunk kezelt, módosított fájlokat (pl. index.html)
            # félreteszi a rebase idejére, utána visszarakja
            rc, _, err = self.git(['rebase', '--autostash', '-X', strategy, upstream])
            if rc != 0:
                self.git(['rebase', '--abort'])
                return self.finish('Összefésülés sikertelen.', STATUS_ERROR,
                                   'Nem sikerült összefésülni a GitHubos változásokkal:\n\n' + err)

        # 5. Ütközésnél a nyertes data.json-t egészben érvényesítjük,
        #    hogy ne legyen belőle félig ez, félig az.
        note = ''
        if conflict:
            backup_stamp = datetime.now().strftime('%Y%m%d-%H%M%S')
            os.makedirs(BACKUP_DIR, exist_ok=True)
            if local_wins:
                rc_show, remote_text, _ = self.git(['show', f'{upstream}:data.json'])
                if rc_show == 0:
                    with open(os.path.join(BACKUP_DIR, f'data-github-{backup_stamp}.json'),
                              'w', encoding='utf-8') as f:
                        f.write(remote_text + '\n')
            elif local_bytes is not None:
                with open(os.path.join(BACKUP_DIR, f'data-helyi-{backup_stamp}.json'), 'wb') as f:
                    f.write(local_bytes)

            if local_wins and local_bytes is not None:
                with open(DATA_FILE, 'wb') as f:
                    f.write(local_bytes)
                note = ' (ütközés: a gépen lévő maradt, a GitHubos a backup mappában)'
            else:
                self.git(['checkout', upstream, '--', 'data.json'])
                note = ' (ütközés: a GitHubos maradt, a helyi a backup mappában)'

            self.git(['add', '--', 'data.json'])
            if self.git(['diff', '--cached', '--quiet'])[0] == 1:
                rc, _, err = self.git(['commit', '-m', f'data.json ütközés feloldva — {stamp}'])
                if rc != 0:
                    return self.finish('Commit sikertelen.', STATUS_ERROR, err)

        # 6. Push, ha van mit
        ahead = self.git(['rev-list', '--count', f'{upstream}..HEAD'])[1]
        if ahead in ('', '0'):
            return self.finish('Naprakész' + note + '.', STATUS_OK)

        rc, _, err = self.git(['push'])
        if rc != 0:
            return self.finish('Push sikertelen.', STATUS_ERROR,
                               'Nem sikerült feltölteni:\n\n' + err
                               + "\n\nHa hitelesítési hibát látsz, futtass egy 'git push' parancsot "
                                 'kézzel a mappában, és jelentkezz be egyszer.')

        self.finish(f'Feltöltve — {stamp}{note}', STATUS_OK)


# --- Főablak ---------------------------------------------------------------

class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('SharePoint AI oldal — kezelő')
        self.resize(1180, 900)

        self.data = self.load_data()
        self.disk_snapshot = read_bytes(DATA_FILE)  # amit utoljára láttunk a lemezen
        self.settings = load_settings()
        self.original_image_path = None
        self.push_worker = None

        # Több gyors változtatást (pl. húzogatás) egy commitba fogunk össze
        self.push_timer = QTimer(self)
        self.push_timer.setSingleShot(True)
        self.push_timer.timeout.connect(self.start_push)

        self.build_ui()
        self.refresh_tree()

        # Indításkor először behúzzuk a GitHub állapotát
        QTimer.singleShot(0, self.start_push)

    # --- Adatkezelés -------------------------------------------------------

    @staticmethod
    def ensure_sections(data):
        for key, *_ in LIST_SECTIONS:
            if not isinstance(data.get(key), list):
                data[key] = []
        if not isinstance(data.get(EVENTS), list):
            data[EVENTS] = []
        return data

    def load_data(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            except (OSError, json.JSONDecodeError):
                data = None

            if data is not None:
                # isNew -> date átállás CSAK memóriában. A fájlhoz betöltéskor nem
                # nyúlunk, mert az friss módosítási időt adna neki, és szinkronnál
                # tévesen "újabbnak" látszana a GitHubos változatnál.
                # A következő mentéskor a már átállított adat kerül ki.
                migrate_cards(data)
                return self.ensure_sections(data)
        return self.ensure_sections({
            'altalanos': {'largeCards': [], 'smallCards': []},
            'copilot': {'largeCards': [], 'smallCards': []},
        })

    def cards_of(self, section, ctype):
        """Egy csoport listája. Az új szekciók sima listák, a régi útmutatók szekció/típus szerint."""
        if ctype is None:
            if not isinstance(self.data.get(section), list):
                self.data[section] = []
            return self.data[section]
        return self.data.setdefault(section, {}).setdefault(ctype, [])

    def set_cards(self, section, ctype, cards):
        if ctype is None:
            self.data[section] = cards
        else:
            self.data.setdefault(section, {})[ctype] = cards

    def save_data(self):
        """Mentés. Ha a data.json-t közben kívülről (kézzel, másik programmal)
        átírták, nem írjuk felül szó nélkül."""
        on_disk = read_bytes(DATA_FILE)
        if on_disk is not None and on_disk != self.disk_snapshot:
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle('A data.json közben megváltozott')
            box.setText('A data.json-t a kezelőn kívül módosították, mióta betöltötted.\n\n'
                        'Betöltés: a lemezen lévő változat jön be, ez az utolsó módosításod elvész.\n'
                        'Felülírás: a kezelőben lévő adat ment ki, a lemezen lévő a backup mappába kerül.')
            reload_btn = box.addButton('Betöltés', QMessageBox.ButtonRole.AcceptRole)
            box.addButton('Felülírás', QMessageBox.ButtonRole.DestructiveRole)
            box.exec()

            if box.clickedButton() is reload_btn:
                self.data = self.load_data()
                self.disk_snapshot = read_bytes(DATA_FILE)
                self.refresh_tree()
                self.clear_form()
                return False

            os.makedirs(BACKUP_DIR, exist_ok=True)
            stamp = datetime.now().strftime('%Y%m%d-%H%M%S')
            with open(os.path.join(BACKUP_DIR, f'data-lemezen-{stamp}.json'), 'wb') as f:
                f.write(on_disk)

        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=4, ensure_ascii=False)
        self.disk_snapshot = read_bytes(DATA_FILE)

        self.schedule_push()
        return True

    # --- Felület -----------------------------------------------------------

    def build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # Fa
        tree_box = QGroupBox('Az oldal tartalma — kattints a szerkesztéshez, húzd a sorokat a sorrendezéshez (vagy Alt + ↑/↓)')
        tree_layout = QVBoxLayout(tree_box)

        self.tree = CardTree()
        self.tree.setHeaderLabels(['Cím', 'Részlet', 'Dátum', 'Kép', 'Link'])
        self.tree.setColumnWidth(1, 220)
        self.tree.itemSelectionChanged.connect(self.on_select)
        self.tree.orderChanged.connect(self.on_order_changed)
        tree_layout.addWidget(self.tree)
        layout.addWidget(tree_box, stretch=1)
        self.tree_box = tree_box

        # Űrlap
        form_box = QGroupBox('Elem szerkesztése / hozzáadása')
        grid = QGridLayout(form_box)
        self.labels = {}

        def label(key, text):
            lbl = QLabel(text)
            self.labels[key] = lbl
            return lbl

        self.section_combo = QComboBox()
        for key, text, _, _ in LIST_SECTIONS:
            self.section_combo.addItem(text, key)
        if SHOW_GUIDE_CARDS:
            for key in SECTIONS:
                self.section_combo.addItem(LABEL_OF[key], key)
        self.section_combo.currentIndexChanged.connect(self.update_form_mode)

        self.type_combo = QComboBox()
        self.type_combo.addItems(CARD_TYPES)
        self.type_combo.setCurrentText('smallCards')
        grid.addWidget(QLabel('Szekció:'), 0, 0)
        grid.addWidget(self.section_combo, 0, 1)
        grid.addWidget(label('type', 'Kártya típusa:'), 0, 2)
        grid.addWidget(self.type_combo, 0, 3)

        self.level_combo = QComboBox()
        self.level_combo.addItems(LEVELS)
        self.size_combo = QComboBox()
        for code, text in SIZES:
            self.size_combo.addItem(text, code)
        self.size_combo.setToolTip('Az oldal négyes blokkokat rak ki: 1 nagy, 2 kicsi, 1 hosszú (közepes).\n'
                                   'A méret dönti el, melyik helyre kerül a projekt.')
        grid.addWidget(label('level', 'Szint:'), 1, 0)
        grid.addWidget(self.level_combo, 1, 1)
        grid.addWidget(label('size', 'Méret:'), 1, 2)
        grid.addWidget(self.size_combo, 1, 3)

        self.title_edit = QLineEdit()
        grid.addWidget(label('title', 'Cím:'), 2, 0)
        grid.addWidget(self.title_edit, 2, 1, 1, 3)

        self.sub_edit = QLineEdit()
        self.sub_edit.setPlaceholderText('egy rövid sor a cím alatt')
        grid.addWidget(label('sub', 'Alcím:'), 3, 0)
        grid.addWidget(self.sub_edit, 3, 1, 1, 3)

        self.text_edit = QPlainTextEdit()
        self.text_edit.setPlaceholderText('A teljes prompt. A [szögletes zárójeles] részeket az oldal kiemeli — '
                                          'ezeket kell a felhasználónak kitöltenie.')
        self.text_edit.setFixedHeight(96)
        grid.addWidget(label('text', 'Prompt szövege:'), 4, 0, Qt.AlignmentFlag.AlignTop)
        grid.addWidget(self.text_edit, 4, 1, 1, 3)

        self.img_edit = QLineEdit()
        self.img_edit.setReadOnly(True)
        self.img_edit.setPlaceholderText('Nincs kép kiválasztva')
        self.browse_btn = QPushButton('Tallózás')
        self.browse_btn.clicked.connect(self.browse_image)
        self.img_clear_btn = QPushButton('Kép törlése')
        self.img_clear_btn.clicked.connect(self.clear_image)
        img_row = QHBoxLayout()
        img_row.addWidget(self.img_edit, stretch=1)
        img_row.addWidget(self.browse_btn)
        img_row.addWidget(self.img_clear_btn)
        grid.addWidget(label('img', 'Kép:'), 5, 0)
        grid.addLayout(img_row, 5, 1, 1, 3)

        self.link_edit = QLineEdit()
        grid.addWidget(label('link', 'Link (cél URL):'), 6, 0)
        grid.addWidget(self.link_edit, 6, 1, 1, 3)

        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat('yyyy.MM.dd')
        self.date_edit.setMinimumDate(NO_DATE)
        self.date_edit.setSpecialValueText('nincs dátum')  # a minimum dátum helyett ez látszik
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.dateChanged.connect(self.update_date_hint)
        self.today_btn = QPushButton('Ma')
        self.today_btn.setToolTip('Mai dátum — az elem újra ' + str(NEW_DAYS) + ' napig ÚJ lesz')
        self.today_btn.clicked.connect(lambda: self.date_edit.setDate(QDate.currentDate()))
        self.date_hint = QLabel()
        self.mark_cb = QCheckBox('Piros felkiáltójel')
        date_row = QHBoxLayout()
        date_row.addWidget(self.date_edit)
        date_row.addWidget(self.today_btn)
        date_row.addWidget(self.date_hint)
        date_row.addStretch()
        date_row.addWidget(self.mark_cb)
        grid.addWidget(label('date', 'Dátum:'), 7, 0)
        grid.addLayout(date_row, 7, 1, 1, 3)

        # mezők -> widgetek (a láthatóság szekciófajtánként változik)
        self.field_widgets = {
            'type': [self.type_combo], 'level': [self.level_combo], 'size': [self.size_combo],
            'title': [self.title_edit], 'sub': [self.sub_edit], 'text': [self.text_edit],
            'img': [self.img_edit, self.browse_btn, self.img_clear_btn], 'link': [self.link_edit],
            'date': [self.date_edit, self.today_btn, self.date_hint],
        }

        # Gombsor
        buttons = QHBoxLayout()
        add_btn = QPushButton('Új hozzáadása')
        add_btn.clicked.connect(self.add_card)
        buttons.addWidget(add_btn)

        self.update_btn = QPushButton('Kiválasztott módosítása')
        self.update_btn.clicked.connect(self.update_card)
        self.update_btn.setEnabled(False)
        buttons.addWidget(self.update_btn)

        self.delete_btn = QPushButton('Kiválasztott törlése')
        self.delete_btn.clicked.connect(self.delete_card)
        self.delete_btn.setEnabled(False)
        buttons.addWidget(self.delete_btn)

        clear_btn = QPushButton('Kijelölés törlése')
        clear_btn.clicked.connect(self.clear_form)
        buttons.addWidget(clear_btn)

        buttons.addSpacing(20)

        self.up_btn = QPushButton('↑ Fel')
        self.up_btn.clicked.connect(lambda: self.nudge_selected(-1))
        self.up_btn.setEnabled(False)
        buttons.addWidget(self.up_btn)

        self.down_btn = QPushButton('↓ Le')
        self.down_btn.clicked.connect(lambda: self.nudge_selected(1))
        self.down_btn.setEnabled(False)
        buttons.addWidget(self.down_btn)

        buttons.addStretch()
        grid.addLayout(buttons, 8, 0, 1, 4)
        layout.addWidget(form_box)
        self.form_box = form_box

        self.update_date_hint()
        self.update_form_mode()

        # Naptár (Excel)
        cal_box = QGroupBox('Naptár — az eseményeket Excelből tölti be')
        cal_layout = QHBoxLayout(cal_box)
        pick_btn = QPushButton('Excel kiválasztása…')
        pick_btn.clicked.connect(lambda: self.import_events(None))
        self.cal_refresh_btn = QPushButton('Frissítés az Excelből')
        self.cal_refresh_btn.clicked.connect(lambda: self.import_events(self.settings.get('naptar_excel')))
        self.cal_status = QLabel()
        cal_layout.addWidget(pick_btn)
        cal_layout.addWidget(self.cal_refresh_btn)
        cal_layout.addWidget(self.cal_status, stretch=1)
        layout.addWidget(cal_box)
        self.cal_box = cal_box
        self.update_cal_status()

        # GitHub
        git_box = QGroupBox('GitHub')
        git_layout = QHBoxLayout(git_box)

        self.push_btn = QPushButton('Feltöltés most')
        self.push_btn.setToolTip('Minden mentés automatikusan feltöltődik — '
                                 'ez a gomb csak nem vár a késleltetésre.')
        self.push_btn.clicked.connect(self.push_now)
        git_layout.addWidget(self.push_btn)

        self.git_status = QLabel('Automatikus feltöltés bekapcsolva.')
        git_layout.addWidget(self.git_status)
        git_layout.addStretch()
        layout.addWidget(git_box)

        # Billentyűparancsok
        QShortcut(QKeySequence('Alt+Up'), self, lambda: self.nudge_selected(-1))
        QShortcut(QKeySequence('Alt+Down'), self, lambda: self.nudge_selected(1))

    # --- Fa feltöltése -----------------------------------------------------

    def groups(self):
        """A fa csoportjai sorrendben: (section, ctype). ctype None = sima lista."""
        result = [(key, None) for key, *_ in LIST_SECTIONS]
        if SHOW_GUIDE_CARDS:
            result += [(sec, ct) for sec in SECTIONS for ct in CARD_TYPES]
        return result

    @staticmethod
    def detail(kind, card):
        if kind == 'card':
            return card.get('level', '')
        if kind == 'tile':
            return ('❗ ' if card.get('mark') else '') + card.get('sub', '')
        if kind == 'prompt':
            text = ' '.join(card.get('text', '').split())
            return text[:70] + ('…' if len(text) > 70 else '')
        if kind == 'usecase':
            return f"{card.get('size', 'S')}  ·  {card.get('sub', '')}"
        return ''

    def refresh_tree(self, select=None):
        """select = (section, ctype, index), ha újra ki akarunk jelölni egy elemet."""
        self.tree.blockSignals(True)
        self.tree.clear()
        bold = None

        for section, ctype in self.groups():
            cards = self.cards_of(section, ctype)
            kind = KIND_OF[section]
            limit = LIMIT_OF.get(section)
            if ctype is None:
                label = f'{LABEL_OF[section]}  ·  {len(cards)} db'
                if limit:
                    label += f'  (az oldalon az első {limit} látszik)'
            else:
                label = f'{LABEL_OF[section]}  ›  {ctype}'
            group = QTreeWidgetItem(self.tree, [label])
            group.setData(0, ROLE, (section, ctype))
            group.setFirstColumnSpanned(True)
            # Nem állítunk fix színt, hogy sötét témában is olvasható maradjon
            bold = group.font(0)
            bold.setBold(True)
            group.setFont(0, bold)
            # A csoport nem húzható, de rá lehet ejteni
            group.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsDropEnabled)

            for i, card in enumerate(cards):
                item = QTreeWidgetItem(group, [
                    card.get('title', ''),
                    self.detail(kind, card),
                    self.date_label(card),
                    card.get('img', ''),
                    card.get('link', ''),
                ])
                item.setData(0, ROLE, card)
                if limit and i >= limit:
                    item.setToolTip(0, f'Ez nem látszik az oldalon — csak az első {limit}. Húzd feljebb, ha kell.')
                    for col in range(5):
                        item.setForeground(col, self.palette().placeholderText())
                # Az elem húzható, de nem ejthető rá másik elem
                item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsDragEnabled)

        # Naptár: csak megjelenítés, az Excel a forrás
        events = self.data.get(EVENTS, [])
        group = QTreeWidgetItem(self.tree, [f'Naptár (Excelből)  ·  {len(events)} esemény — itt nem szerkeszthető'])
        group.setData(0, ROLE, (EVENTS, None))
        group.setFirstColumnSpanned(True)
        if bold is not None:
            group.setFont(0, bold)
        group.setFlags(Qt.ItemFlag.ItemIsEnabled)
        for e in events:
            when = (e.get('start') or 'egész nap') + (f"–{e['end']}" if e.get('end') else '')
            place = f"  ·  {e['location']}" if e.get('location') else ''
            d = parse_date(e.get('date'))
            item = QTreeWidgetItem(group, [e.get('title', ''), when + place,
                                           d.strftime('%Y.%m.%d') if d else e.get('date', ''), '', e.get('link', '')])
            item.setFlags(Qt.ItemFlag.ItemIsEnabled)
        group.setExpanded(len(events) <= 12)

        for i in range(self.tree.topLevelItemCount() - 1):
            self.tree.topLevelItem(i).setExpanded(True)
        self.tree.blockSignals(False)

        if select:
            self.select_card(*select)

    def select_card(self, section, ctype, index):
        for i in range(self.tree.topLevelItemCount()):
            group = self.tree.topLevelItem(i)
            if tuple(group.data(0, ROLE)) != (section, ctype):
                continue
            if 0 <= index < group.childCount():
                item = group.child(index)
                self.tree.setCurrentItem(item)
                self.tree.scrollToItem(item)
            return

    def selected_path(self):
        """(section, ctype, index) vagy None."""
        item = self.tree.currentItem()
        if item is None or item.parent() is None:
            return None
        group = item.parent()
        section, ctype = group.data(0, ROLE)
        if section == EVENTS:
            return None
        return section, ctype, group.indexOfChild(item)

    # --- Sorrendezés -------------------------------------------------------

    def on_order_changed(self):
        """Húzás után a fa a mérvadó: abból építjük újra a listákat."""
        for i in range(self.tree.topLevelItemCount()):
            group = self.tree.topLevelItem(i)
            section, ctype = group.data(0, ROLE)
            if section == EVENTS:
                continue
            self.set_cards(section, ctype, [group.child(j).data(0, ROLE) for j in range(group.childCount())])

        self.save_data()
        QTimer.singleShot(0, self.refresh_tree)   # a húzás befejezése után frissítünk (szürkítés a limit felett)

    def nudge_selected(self, delta):
        path = self.selected_path()
        if not path:
            return

        section, ctype, index = path
        cards = self.cards_of(section, ctype)
        new_index = index + delta

        if not 0 <= new_index < len(cards):
            return

        cards[index], cards[new_index] = cards[new_index], cards[index]
        self.save_data()
        self.refresh_tree(select=(section, ctype, new_index))

    # --- Űrlap <-> adat ----------------------------------------------------

    def current_section(self):
        return self.section_combo.currentData()

    def kind(self):
        return KIND_OF[self.current_section()]

    def set_section(self, key):
        idx = self.section_combo.findData(key)
        if idx >= 0:
            self.section_combo.setCurrentIndex(idx)

    def target_group(self):
        """Az űrlap szerinti cél csoport: (section, ctype)."""
        section = self.current_section()
        return (section, self.type_combo.currentText()) if self.kind() == 'card' else (section, None)

    def update_form_mode(self, *_):
        """Csak azok a mezők látszanak, amik az adott szekcióhoz tartoznak."""
        kind = self.kind()
        fields = FIELDS[kind]
        for key, widgets in self.field_widgets.items():
            shown = key in fields
            for w in widgets:
                w.setVisible(shown)
            if key in self.labels:
                self.labels[key].setVisible(shown)
        self.mark_cb.setVisible(self.current_section() == 'kisokos')
        self.img_clear_btn.setVisible('img' in fields and kind not in IMG_REQUIRED)
        self.labels['img'].setText({'usecase': 'Ikon (nem kötelező):', 'tile': 'Kép (nem kötelező):'}.get(kind, 'Kép:'))
        self.labels['title'].setText('Cím (a keresőhöz is!):' if kind != 'prompt' else 'A prompt neve:')

    def on_select(self):
        path = self.selected_path()
        if not path:
            for btn in (self.update_btn, self.delete_btn, self.up_btn, self.down_btn):
                btn.setEnabled(False)
            return

        section, ctype, index = path
        card = self.cards_of(section, ctype)[index]
        self.original_image_path = None

        self.set_section(section)
        if ctype:
            self.type_combo.setCurrentText(ctype)
        self.level_combo.setCurrentText(card.get('level', 'kezdo'))
        size_idx = self.size_combo.findData(card.get('size', 'S'))
        self.size_combo.setCurrentIndex(size_idx if size_idx >= 0 else 2)
        self.title_edit.setText(card.get('title', ''))
        self.sub_edit.setText(card.get('sub', ''))
        self.text_edit.setPlainText(card.get('text', ''))
        self.mark_cb.setChecked(bool(card.get('mark')))
        self.img_edit.setText(os.path.basename(card.get('img', '')))
        self.link_edit.setText(card.get('link', '#'))
        d = parse_date(card.get('date'))
        self.date_edit.setDate(QDate(d.year, d.month, d.day) if d else NO_DATE)

        for btn in (self.update_btn, self.delete_btn, self.up_btn, self.down_btn):
            btn.setEnabled(True)

    def clear_form(self):
        self.tree.clearSelection()
        self.tree.setCurrentItem(None)
        self.original_image_path = None
        for w in (self.title_edit, self.img_edit, self.link_edit, self.sub_edit):
            w.clear()
        self.text_edit.clear()
        self.mark_cb.setChecked(False)
        self.date_edit.setDate(QDate.currentDate())
        for btn in (self.update_btn, self.delete_btn, self.up_btn, self.down_btn):
            btn.setEnabled(False)

    # --- Kép ---------------------------------------------------------------

    def browse_image(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            'Válassz képet',
            '',
            'Képfájlok (*.png *.jpg *.jpeg *.gif *.webp *.svg);;Minden fájl (*)'
        )
        if file_path:
            self.original_image_path = os.path.abspath(file_path)
            self.img_edit.setText(os.path.basename(file_path))

    def clear_image(self):
        self.original_image_path = None
        self.img_edit.clear()

    def resolve_destination(self, filename, source):
        """(végleges fájlnév, cél útvonal, másolás kihagyható-e).

        Ha a név foglalt, de a tartalom ugyanaz, újrahasználjuk a meglévő fájlt.
        Ha foglalt és más a tartalom, -1, -2 ... utótagot kap.
        """
        base, ext = os.path.splitext(filename)
        candidate = filename
        counter = 0

        while True:
            dest = os.path.join(IMG_PATH, candidate)
            if not os.path.exists(dest):
                return candidate, dest, False
            if same_content(source, dest):
                return candidate, dest, True
            counter += 1
            candidate = f'{base}-{counter}{ext}'

    def resolve_image(self, required):
        """Új kép esetén bemásolja az img/ mappába. Visszaadja a végleges fájlnevet,
        '' ha nincs kép (és nem is kötelező), None hiba esetén."""
        source = self.original_image_path

        # Nincs új kép kiválasztva -> az elem meglévő képe marad
        if not source:
            existing = self.img_edit.text().strip()
            if not existing:
                if required:
                    QMessageBox.critical(self, 'Hiba', 'Kép megadása kötelező!')
                    return None
                return ''
            return os.path.basename(existing)

        ext = os.path.splitext(source)[1].lower()
        if ext not in ALLOWED_EXT:
            QMessageBox.critical(
                self, 'Hiba',
                f'Nem támogatott képformátum: {ext or "(nincs kiterjesztés)"}'
            )
            return None

        try:
            size = os.path.getsize(source)
        except OSError as error:
            QMessageBox.critical(self, 'Hiba', str(error))
            return None

        if size > MAX_IMG_BYTES:
            mb = size / (1024 * 1024)
            answer = QMessageBox.question(
                self, 'Nagy fájl',
                f'A kép {mb:.1f} MB. Ez lassíthatja az oldal betöltését. '
                f'Biztosan ezt akarod használni?'
            )
            if answer != QMessageBox.StandardButton.Yes:
                return None

        filename = sanitize_filename(os.path.basename(source))

        try:
            os.makedirs(IMG_PATH, exist_ok=True)
            filename, destination, already_there = self.resolve_destination(filename, source)
            if not already_there:
                shutil.copy2(source, destination)
        except OSError as error:
            QMessageBox.critical(self, 'Hiba', str(error))
            return None

        return filename

    def get_form_data(self):
        kind = self.kind()
        fields = FIELDS[kind]
        title = self.title_edit.text().strip()

        if not title:
            QMessageBox.critical(self, 'Hiba', 'A cím megadása kötelező — erre keres a kereső is!')
            return None

        text = self.text_edit.toPlainText().strip()
        if kind == 'prompt' and not text:
            QMessageBox.critical(self, 'Hiba', 'A prompt szövege üres.')
            return None

        img_name = ''
        if 'img' in fields:
            img_name = self.resolve_image(required=kind in IMG_REQUIRED)
            if img_name is None:
                return None

        link = self.link_edit.text().strip()
        if not link:
            link = '#'
        elif link != '#' and not link.startswith('http') and not link.startswith('/'):
            link = f'https://{link}'

        card = {'title': title}
        if 'sub' in fields:
            card['sub'] = self.sub_edit.text().strip()
        if 'text' in fields:
            card['text'] = text
        if 'size' in fields:
            card['size'] = self.size_combo.currentData()
        if 'level' in fields:
            card['level'] = self.level_combo.currentText()
        if img_name:
            card['img'] = f'img/{img_name}'
        if 'link' in fields:
            card['link'] = link
        if self.current_section() == 'kisokos':
            card['mark'] = self.mark_cb.isChecked()
        if 'date' in fields and self.date_edit.date() != NO_DATE:
            card['date'] = self.date_edit.date().toString('yyyy-MM-dd')
        return card

    # --- Dátum megjelenítés ------------------------------------------------

    @staticmethod
    def date_label(card):
        d = parse_date(card.get('date'))
        if d is None:
            return '—'
        left = days_left(card)
        text = d.strftime('%Y.%m.%d')
        return f'{text}  · ÚJ még {left} napig' if left is not None else text

    def update_date_hint(self):
        qd = self.date_edit.date()
        if qd == NO_DATE:
            self.date_hint.setText('nem kap ÚJ jelölést')
            return
        left = days_left({'date': qd.toString('yyyy-MM-dd')})
        if left is None:
            self.date_hint.setText('már nem ÚJ')
        else:
            self.date_hint.setText(f'ÚJ jelölés még {left} napig')

    # --- CRUD --------------------------------------------------------------

    def add_card(self):
        new_card = self.get_form_data()
        if not new_card:
            return

        section, ctype = self.target_group()
        cards = self.cards_of(section, ctype)
        cards.append(new_card)

        if not self.save_data():
            return
        self.refresh_tree()
        self.clear_form()

        limit = LIMIT_OF.get(section)
        if limit and len(cards) > limit:
            QMessageBox.information(
                self, 'Siker',
                f'Hozzáadva — de ebben a sávban csak az első {limit} látszik. '
                f'Húzd feljebb, ha ennek is meg kell jelennie.')
        else:
            QMessageBox.information(self, 'Siker', 'Hozzáadva!')

    def update_card(self):
        path = self.selected_path()
        if not path:
            return

        updated_card = self.get_form_data()
        if not updated_card:
            return

        old_section, old_ctype, index = path
        new_section, new_ctype = self.target_group()
        old_card = self.cards_of(old_section, old_ctype)[index]

        if (old_section, old_ctype) == (new_section, new_ctype):
            # A kezelő által nem ismert, kézzel beírt mezők (pl. promo "button") megmaradnak
            extras = {k: v for k, v in old_card.items() if k not in MANAGED_KEYS}
            self.cards_of(old_section, old_ctype)[index] = {**updated_card, **extras}
            target = (old_section, old_ctype, index)
        else:
            self.cards_of(old_section, old_ctype).pop(index)
            dest = self.cards_of(new_section, new_ctype)
            dest.append(updated_card)
            target = (new_section, new_ctype, len(dest) - 1)

        if not self.save_data():
            return
        self.refresh_tree(select=target)
        QMessageBox.information(self, 'Siker', 'Módosítva!')

    def delete_card(self):
        path = self.selected_path()
        if not path:
            return

        answer = QMessageBox.question(self, 'Megerősítés', 'Biztosan törölni szeretnéd ezt az elemet?')
        if answer != QMessageBox.StandardButton.Yes:
            return

        section, ctype, index = path
        self.cards_of(section, ctype).pop(index)
        if not self.save_data():
            return
        self.refresh_tree()
        self.clear_form()

    # --- Naptár (Excel) ----------------------------------------------------

    def update_cal_status(self):
        path = self.settings.get('naptar_excel')
        self.cal_refresh_btn.setEnabled(bool(path))
        if not path:
            self.cal_status.setText('Még nincs Excel kiválasztva.')
        elif not os.path.exists(path):
            self.cal_status.setText(f'A megjegyzett Excel nem található: {path}')
        else:
            last = self.settings.get('naptar_utolso_betoltes', '')
            self.cal_status.setText(f'{os.path.basename(path)}' + (f'  ·  utoljára betöltve: {last}' if last else ''))
        self.cal_status.setToolTip(path or '')

    def import_events(self, path):
        if not path:
            path, _ = QFileDialog.getOpenFileName(self, 'Naptár Excel kiválasztása', '',
                                                  'Excel (*.xlsx *.xlsm);;Minden fájl (*)')
            if not path:
                return
        try:
            events, skipped = read_events_xlsx(path)
        except ImportError:
            QMessageBox.critical(self, 'Hiányzó csomag',
                                 'Az Excel olvasásához kell az openpyxl csomag. Telepítsd egyszer:\n\n'
                                 'pip install openpyxl')
            return
        except PermissionError:
            QMessageBox.critical(self, 'Hiba', 'Nem tudom megnyitni az Excelt. Ha nyitva van, zárd be, és próbáld újra.')
            return
        except Exception as error:  # rossz formátum, sérült fájl stb.
            QMessageBox.critical(self, 'Hiba', f'Nem sikerült beolvasni az Excelt:\n\n{error}')
            return

        self.data[EVENTS] = events
        if not self.save_data():
            return
        self.settings['naptar_excel'] = os.path.abspath(path)
        self.settings['naptar_utolso_betoltes'] = datetime.now().strftime('%Y.%m.%d %H:%M')
        save_settings(self.settings)
        self.update_cal_status()
        self.refresh_tree()

        msg = f'{len(events)} esemény betöltve.'
        if skipped:
            shown = ', '.join(str(r) for r in skipped[:12]) + (' …' if len(skipped) > 12 else '')
            msg += f'\n\nKihagyott sorok (nincs dátum vagy cím): {shown}'
        QMessageBox.information(self, 'Naptár frissítve', msg)

    # --- GitHub ------------------------------------------------------------

    def set_git_status(self, text, color=STATUS_NEUTRAL):
        """color=None esetén a téma alap szövegszíne marad, így sötét témán is látszik."""
        self.git_status.setText(text)
        self.git_status.setStyleSheet(f'color: {color};' if color else '')

    def schedule_push(self):
        """Minden mentés után hívódik. Nem indít azonnal feltöltést: vár egy kicsit,
        hogy a gyors egymás utáni változtatások egy commitba kerüljenek."""
        self.set_git_status('Változás mentve — feltöltés hamarosan...', STATUS_PENDING)
        self.push_timer.start(PUSH_DELAY_MS)

    def push_now(self):
        self.push_timer.stop()
        self.start_push()

    def start_push(self):
        # Ha még fut az előző feltöltés, kicsit később próbáljuk újra
        if self.push_worker is not None and self.push_worker.isRunning():
            self.push_timer.start(PUSH_DELAY_MS)
            return

        self.push_btn.setEnabled(False)
        # Szinkron alatt nem lehet szerkeszteni, különben a git felülírhatná a friss mentést
        self.tree_box.setEnabled(False)
        self.form_box.setEnabled(False)
        self.cal_box.setEnabled(False)
        self.set_git_status('Szinkronizálás a GitHubbal...')

        self.push_worker = GitSyncWorker()
        self.push_worker.done.connect(self.on_push_done)
        self.push_worker.start()

    def on_push_done(self, status, color, error_detail, data_changed):
        self.set_git_status(status, color or STATUS_NEUTRAL)
        self.push_btn.setEnabled(True)
        self.tree_box.setEnabled(True)
        self.form_box.setEnabled(True)
        self.cal_box.setEnabled(True)

        # Ha a GitHubról jött új data.json, töltsük újra a felületet
        if data_changed:
            self.data = self.load_data()
            self.refresh_tree()
            self.clear_form()
        self.disk_snapshot = read_bytes(DATA_FILE)

        if error_detail:
            QMessageBox.critical(self, 'GitHub hiba', error_detail)

    def closeEvent(self, event):
        """Kilépés előtt még feltöltjük, ami a késleltetés miatt bent ragadt."""
        if self.push_timer.isActive():
            self.push_timer.stop()
            if self.push_worker is None or not self.push_worker.isRunning():
                self.push_worker = GitSyncWorker()
                self.push_worker.start()

        if self.push_worker is not None and self.push_worker.isRunning():
            self.set_git_status('Feltöltés befejezése...')
            self.push_worker.wait(30000)

        event.accept()


if __name__ == '__main__':
    # Mindig a script mappájából dolgozzon, akárhonnan indítják
    os.chdir(os.path.dirname(os.path.abspath(sys.argv[0])))
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec())