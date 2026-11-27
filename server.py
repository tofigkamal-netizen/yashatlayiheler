# -*- coding: utf-8 -*-
"""
Lokal backend: layiheler.xlsx faylını Apps Script əvəzinə istifadə edir.

İşə salmaq:   python server.py
Açmaq:        http://localhost:8000

- index.html-i verir və POST /api sorğularına Apps Script ilə eyni formatda
  ({ok, data} / {ok:false, error}) cavab verir.
- İstifadəçilər, Qeydlər, Log vərəqləri layiheler.xlsx-dən oxunur və ora yazılır.
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
XLSX = os.path.join(BASE, 'layiheler.xlsx')
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


def hash_pw(pw, salt=None):
    salt = salt or secrets.token_hex(8)
    return 'sha256$%s$%s' % (salt, hashlib.sha256((salt + pw).encode('utf-8')).hexdigest())


def check_pw(pw, stored):
    stored = str(stored or '')
    if not stored.startswith('sha256$'):
        return stored != '' and stored == pw  # açıq mətnlə yazılmış parol
    try:
        _, salt, h = stored.split('$', 2)
    except ValueError:
        return False
    cands = (salt + pw, pw + salt)
    return any(hashlib.sha256(c.encode('utf-8')).hexdigest() == h for c in cands)


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
