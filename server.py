# -*- coding: utf-8 -*-
"""
Lokal backend: layiheler.xlsx faylını Apps Script əvəzinə istifadə edir.

İşə salmaq:   python server.py
Açmaq:        http://localhost:8000

- index.html-i verir və POST /api sorğularına Apps Script ilə eyni formatda
  ({ok, data} / {ok:false, error}) cavab verir.
- İstifadəçilər, Qeydlər, Log vərəqləri layiheler.xlsx-dən oxunur və ora yazılır.
- Layihələr, Ödənişlər, İştirakçı tarixçəsi, Parametrlər vərəqlərində proqramın bütün
  məlumatı saxlanılır (getData / syncData). Fayl Google Sheets-ə yüklənsə, eyni vərəqlərlə
  apps-script/Code.gs işləyir.
- "admin" / "admin" istifadəçisi yoxdursa, avtomatik əlavə olunur.
- Diqqət: Excel faylı Excel-də açıq olarsa, yazmaq alınmayacaq — faylı bağlayın.
"""
import hashlib
import json
import os
import secrets
import sys
import threading
from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import openpyxl

BASE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.environ.get('LAYIHE_XLSX') or os.path.join(BASE, 'layiheler.xlsx')
PORT = int(os.environ.get('PORT', '8000'))

SH_USERS, SH_NOTES, SH_LOG = 'İstifadəçilər', 'Qeydlər', 'Log'
OPEN, DONE, DELETED = 'Açıq', 'Həll edildi', 'Silindi'

# Qeydlər vərəqinin sütunları (1-dən başlayır)
C_ID, C_PARENT, C_DATE, C_LOGIN, C_NAME, C_TEXT, C_SCREEN, C_KEY, C_SEL, C_X, C_Y, C_STATUS, C_EDITED = range(1, 14)

LOCK = threading.Lock()
SESSIONS = {}  # token -> user dict


# ---------- köməkçilər ----------
def iso(v):
    if isinstance(v, datetime):
        return v.isoformat(timespec='milliseconds')
    return v or ''


HASH_ROUNDS = 100  # Apps Script-dəki hashPassword_ ilə eyni


def hash_pw(pw, salt=None):
    """Apps Script hashPassword_ ilə eyni: h = salt + parol; 100 dəfə h = sha256(h + salt)."""
    salt = salt or secrets.token_hex(8)
    h = salt + str(pw)
    for _ in range(HASH_ROUNDS):
        h = hashlib.sha256((h + salt).encode('utf-8')).hexdigest()
    return 'sha256$%s$%s' % (salt, h)


def check_pw(pw, stored):
    stored = str(stored or '')
    if not stored.startswith('sha256$'):
        return stored != '' and stored == pw  # açıq mətnlə yazılmış parol
    try:
        _, salt, h = stored.split('$', 2)
    except ValueError:
        return False
    if hash_pw(pw, salt) == stored:
        return True
    # Köhnə lokal format (server.py-ın əvvəlki versiyası, tək sha256) — geriyə uyğunluq üçün
    return hashlib.sha256((salt + pw).encode('utf-8')).hexdigest() == h


def load():
    return openpyxl.load_workbook(XLSX)


def save(wb):
    try:
        wb.save(XLSX)
    except PermissionError:
        raise RuntimeError('layiheler.xlsx yazıla bilmədi — fayl Excel-də açıqdırsa, bağlayın.')


def last_row(ws):
    r = ws.max_row
    while r > 1 and all(ws.cell(r, c).value in (None, '') for c in range(1, ws.max_column + 1)):
        r -= 1
    return r


def append(ws, values):
    r = last_row(ws) + 1
    for i, v in enumerate(values, 1):
        ws.cell(r, i, v)


def log(wb, user, op, note_id, old=None, new=None):
    append(wb[SH_LOG], [datetime.now(), user['login'], user['name'], op, note_id, old, new])


def user_rows(ws):
    for r in range(2, last_row(ws) + 1):
        login = ws.cell(r, 1).value
        if login:
            yield r, str(login).strip()


def ensure_admin():
    wb = load()
    ws = wb[SH_USERS]
    if any(l.lower() == 'admin' for _, l in user_rows(ws)):
        return
    append(ws, ['admin', hash_pw('admin'), 'Admin', 'Bəli', 'admin'])
    save(wb)
    print('admin / admin istifadəçisi əlavə olundu.')


def auth(token):
    u = SESSIONS.get(token)
    if not u:
        raise RuntimeError('SESSION: sessiya etibarsızdır')
    return u


def note_from_row(ws, r):
    g = lambda c: ws.cell(r, c).value
    return {
        'id': g(C_ID), 'parentId': g(C_PARENT) or '', 'date': iso(g(C_DATE)),
        'login': g(C_LOGIN) or '', 'name': g(C_NAME) or '', 'text': g(C_TEXT) or '',
        'screen': g(C_SCREEN) or '', 'screenKey': g(C_KEY) or '', 'selector': g(C_SEL) or '',
        'x': g(C_X), 'y': g(C_Y), 'status': g(C_STATUS) or OPEN, 'edited': iso(g(C_EDITED)),
    }


def find_note(ws, note_id):
    for r in range(2, last_row(ws) + 1):
        if ws.cell(r, C_ID).value == note_id and ws.cell(r, C_STATUS).value != DELETED:
            return r
    raise RuntimeError('Qeyd tapılmadı')


def own_or_admin(ws, r, user):
    if user.get('role') == 'admin':
        return
    if str(ws.cell(r, C_LOGIN).value or '').lower() != user['login'].lower():
        raise RuntimeError('Yalnız öz qeydinizi dəyişə bilərsiniz')


def pos_text(screen, sel, x, y):
    return '%s | %s (%s%%, %s%%)' % (screen or '', sel or '', x, y)


# ---------- API əməliyyatları ----------
def a_login(login, pw, remember=False):
    login = str(login or '').strip()
    wb = load()
    ws = wb[SH_USERS]
    for r, l in user_rows(ws):
        if l.lower() != login.lower():
            continue
        if str(ws.cell(r, 4).value or 'Bəli').strip().lower() not in ('bəli', 'beli', 'yes', 'true', '1'):
            raise RuntimeError('Bu istifadəçi aktiv deyil')
        if not check_pw(pw, ws.cell(r, 2).value):
            break
        user = {'login': l, 'name': ws.cell(r, 3).value or l, 'role': ws.cell(r, 5).value or ''}
        token = secrets.token_urlsafe(24)
        SESSIONS[token] = user
        return {'token': token, 'user': user, 'mustChange': False}
    raise RuntimeError('Login və ya parol yanlışdır')


def a_logout(token):
    SESSIONS.pop(token, None)
    return True


def a_changePassword(token, old, new):
    user = auth(token)
    if len(new or '') < 6:
        raise RuntimeError('Yeni parol ən azı 6 simvol olmalıdır')
    wb = load()
    ws = wb[SH_USERS]
    for r, l in user_rows(ws):
        if l.lower() == user['login'].lower():
            if not check_pw(old, ws.cell(r, 2).value):
                raise RuntimeError('Köhnə parol yanlışdır')
            ws.cell(r, 2, hash_pw(new))
            save(wb)
            return True
    raise RuntimeError('İstifadəçi tapılmadı')


def a_getNotes(token):
    auth(token)
    ws = load()[SH_NOTES]
    return [note_from_row(ws, r) for r in range(2, last_row(ws) + 1)
            if ws.cell(r, C_ID).value and ws.cell(r, C_STATUS).value != DELETED]


def a_addNote(token, data):
    user = auth(token)
    text = str(data.get('text') or '').strip()
    if not text:
        raise RuntimeError('Mətn boşdur')
    wb = load()
    ws = wb[SH_NOTES]
    nid = 'Q' + secrets.token_hex(6)
    parent = data.get('parentId') or None
    append(ws, [nid, parent, datetime.now(), user['login'], user['name'], text,
                data.get('screen'), data.get('screenKey'), data.get('selector'),
                data.get('x'), data.get('y'), OPEN, None])
    log(wb, user, 'Cavab yazıldı' if parent else 'Qeyd yaradıldı', nid, None, text)
    save(wb)
    return note_from_row(ws, last_row(ws))


def a_editNote(token, nid, text):
    user = auth(token)
    wb = load()
    ws = wb[SH_NOTES]
    r = find_note(ws, nid)
    own_or_admin(ws, r, user)
    old = ws.cell(r, C_TEXT).value
    ws.cell(r, C_TEXT, text)
    ws.cell(r, C_EDITED, datetime.now())
    log(wb, user, 'Redaktə edildi', nid, old, text)
    save(wb)
    return note_from_row(ws, r)


def a_deleteNote(token, nid):
    user = auth(token)
    wb = load()
    ws = wb[SH_NOTES]
    r = find_note(ws, nid)
    own_or_admin(ws, r, user)
    for rr in range(2, last_row(ws) + 1):
        if ws.cell(rr, C_ID).value == nid or ws.cell(rr, C_PARENT).value == nid:
            ws.cell(rr, C_STATUS, DELETED)
    log(wb, user, 'Silindi', nid, ws.cell(r, C_TEXT).value, None)
    save(wb)
    return True


def a_moveNote(token, nid, pos):
    user = auth(token)
    wb = load()
    ws = wb[SH_NOTES]
    r = find_note(ws, nid)
    g = lambda c: ws.cell(r, c).value
    old = pos_text(g(C_SCREEN), g(C_SEL), g(C_X), g(C_Y))
    for c, k in ((C_SCREEN, 'screen'), (C_KEY, 'screenKey'), (C_SEL, 'selector'), (C_X, 'x'), (C_Y, 'y')):
        ws.cell(r, c, pos.get(k))
    log(wb, user, 'Yeri dəyişdi', nid, old, pos_text(pos.get('screen'), pos.get('selector'), pos.get('x'), pos.get('y')))
    save(wb)
    return True


def a_setNoteStatus(token, nid, status):
    user = auth(token)
    if status not in (OPEN, DONE):
        raise RuntimeError('Yanlış status')
    wb = load()
    ws = wb[SH_NOTES]
    r = find_note(ws, nid)
    old = ws.cell(r, C_STATUS).value
    ws.cell(r, C_STATUS, status)
    log(wb, user, 'Status dəyişdi', nid, old, status)
    save(wb)
    return True


# ---------- Layihə məlumatları (layihələr, ödənişlər, iştirakçı tarixçəsi, parametrlər) ----------
# Hər qeyd bir sətirdir: 1-ci sütun açar, sonra oxunaqlı sütunlar, sonda tam obyekt JSON kimi.
# JSON uzun olarsa "JSON", "JSON 2", "JSON 3"... sütunlarına bölünür (Excel xanası ≤ 32767 simvol).
# Vərəq və sütun adları Google Sheet (Apps Script, apps-script/Code.gs) ilə eynidir.
JSON_CHUNK = 30000


def _d(v):
    v = str(v or '')
    return '.'.join(reversed(v.split('-'))) if len(v) == 10 and v[4] == '-' else v


DATA_SHEETS = {
    'layihe': ('Layihələr', ['ID', 'Tip', 'Əsas layihə ID', 'Ad', 'Status', 'Başlama', 'Bitmə'],
               lambda o: [o.get('tip', ''), o.get('parentId', ''), o.get('ad', ''), o.get('status', ''),
                          _d(o.get('baslama') or o.get('yaranma')), _d(o.get('bitme'))]),
    'odenis': ('Ödənişlər', ['ID', 'Layihə ID', 'Ödənişin təyinatı', 'Ödəmə tarixi', 'Ümumi (AZN)', 'Status'],
               lambda o: [o.get('layiheId', ''), o.get('teyinat', ''), _d(o.get('tarix')), o.get('umumi', ''), o.get('status', '')]),
    'istirakTarixce': ('İştirakçı tarixçəsi', ['FİN', 'Qeyd sayı', 'Son əməliyyat'],
                       lambda o: [len(o), (o[-1].get('emel', '') + ' — ' + o[-1].get('layihe', '')) if o else '']),
    'parametr': ('Parametrlər', ['Açar', 'Təsvir'], lambda o: ['Formdan əlavə edilən ölkə və şəhərlər']),
}
META = ['Yenilənib', 'Yeniləyən']


def data_sheet(wb, kind):
    name, cols, _ = DATA_SHEETS[kind]
    if name in wb.sheetnames:
        return wb[name]
    ws = wb.create_sheet(name)
    for i, h in enumerate(cols + META + ['JSON'], 1):
        ws.cell(1, i, h)
    return ws


def json_col(kind):
    return len(DATA_SHEETS[kind][1]) + len(META) + 1


def read_json(ws, r, kind):
    c0 = json_col(kind)
    parts = []
    for c in range(c0, ws.max_column + 1):
        v = ws.cell(r, c).value
        if v in (None, ''):
            break
        parts.append(str(v))
    return json.loads(''.join(parts)) if parts else None


def a_getData(token):
    auth(token)
    wb = load()
    out = {'layihe': [], 'odenis': [], 'istirakTarixce': {}, 'parametr': {}}
    for kind, (name, _, _) in DATA_SHEETS.items():
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        for r in range(2, last_row(ws) + 1):
            key = ws.cell(r, 1).value
            if key in (None, ''):
                continue
            obj = read_json(ws, r, kind)
            if obj is None:
                continue
            if kind in ('layihe', 'odenis'):
                out[kind].append(obj)
            else:
                out[kind][str(key)] = obj
    out['layihe'].sort(key=lambda o: o.get('id', 0))
    out['odenis'].sort(key=lambda o: o.get('id', 0))
    return out


def a_syncData(token, ops):
    user = auth(token)
    wb = load()
    for op in ops or []:
        kind, key = op.get('kind'), str(op.get('id'))
        if kind not in DATA_SHEETS:
            raise RuntimeError('Naməlum məlumat növü: %s' % kind)
        ws = data_sheet(wb, kind)
        row = next((r for r in range(2, last_row(ws) + 1) if str(ws.cell(r, 1).value) == key), None)
        if op.get('op') == 'remove':
            if row:
                ws.delete_rows(row)
            continue
        obj = op.get('obj')
        if row is None:
            row = last_row(ws) + 1
        _, cols, readable = DATA_SHEETS[kind]
        text = json.dumps(obj, ensure_ascii=False)
        chunks = [text[i:i + JSON_CHUNK] for i in range(0, len(text), JSON_CHUNK)] or ['']
        values = [int(key) if key.isdigit() else key] + readable(obj) + [datetime.now(), user['name']] + chunks
        c0 = json_col(kind)
        for c in range(c0 + len(chunks), ws.max_column + 1):  # köhnə artıq JSON hissələrini təmizlə
            ws.cell(row, c).value = None
        for i, v in enumerate(values, 1):
            ws.cell(row, i, v)
        for n in range(2, len(chunks) + 1):
            if ws.cell(1, c0 + n - 1).value in (None, ''):
                ws.cell(1, c0 + n - 1, 'JSON %d' % n)
    save(wb)
    return True


ACTIONS = {k[2:]: v for k, v in globals().items() if k.startswith('a_')}


# ---------- HTTP ----------
class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=BASE, **kw)

    def do_GET(self):
        if self.path in ('/', ''):
            self.path = '/index.html'
        if self.path.split('?')[0].endswith(('.py', '.xlsx')):
            return self.send_error(404)
        return super().do_GET()

    def do_POST(self):
        if self.path.split('?')[0] != '/api':
            return self.send_error(404)
        try:
            body = json.loads(self.rfile.read(int(self.headers.get('Content-Length') or 0)).decode('utf-8'))
            fn = ACTIONS.get(body.get('action'))
            if not fn:
                raise RuntimeError('Naməlum əməliyyat: %s' % body.get('action'))
            with LOCK:
                res = {'ok': True, 'data': fn(*(body.get('args') or []))}
        except Exception as e:  # noqa: BLE001
            res = {'ok': False, 'error': str(e)}
        out = json.dumps(res, ensure_ascii=False, default=str).encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(out)))
        self.end_headers()
        self.wfile.write(out)


if __name__ == '__main__':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:  # noqa: BLE001
        pass
    ensure_admin()
    print('Server: http://localhost:%d  (dayandırmaq üçün Ctrl+C)' % PORT)
    HTTPServer(('127.0.0.1', PORT), Handler).serve_forever()
