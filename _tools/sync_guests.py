#!/usr/bin/env python3
"""
Sync the website's GUEST_LIST from the Google Sheet.

Usage:
    python3 _tools/sync_guests.py <sheet_export.xlsx> <index.html>

Source of truth: the 'Wedding Guest List' tab of "CJ & Sheilla-Wedding Prep"
  - column C (ATTENDEES)     -> guest name   (rows 6 down)
  - column G (RESERVED SEAT) -> seats        (blank -> 1)
A suffix like " + baby" is dropped from the name the site shows/submits
(the sheet's ATTENDING? formula ignores that suffix too, so RSVPs still match).

Nickname keys already on the site for a guest are kept; new guests get
their full name plus their first name (if no one else shares it).

Prints a summary and exits 0. First line is "CHANGED" or "NO CHANGES".
"""
import re, sys
import openpyxl

TAB, NAME_COL, SEAT_COL, FIRST_ROW = 'Wedding Guest List', 3, 7, 6
START = 'var GUEST_LIST = ['
COMMENT = ("// GUEST LIST - auto-synced from the 'Wedding Guest List' tab of the "
           "Google Sheet. Edit the sheet, not this list.")


def norm(s):
    return re.sub(r'\s+', ' ', s.strip()).lower()


def display_name(raw):
    return re.sub(r'\s+', ' ', raw.split('+')[0]).strip()


def js_str(s):
    return "'" + s.replace('\\', '\\\\').replace("'", "\\'") + "'"


def read_sheet(path):
    ws = openpyxl.load_workbook(path, data_only=True)[TAB]
    guests, seen = [], set()
    for r in range(FIRST_ROW, ws.max_row + 1):
        raw = ws.cell(r, NAME_COL).value
        if raw is None or not str(raw).strip():
            continue
        name = display_name(str(raw))
        if not name or norm(name) in seen:
            continue
        seat = ws.cell(r, SEAT_COL).value
        try:
            seats = int(float(seat)) if seat not in (None, '') else 1
        except (TypeError, ValueError):
            seats = 1
        guests.append((name, max(seats, 1)))
        seen.add(norm(name))
    return guests


def read_site(html):
    i = html.index(START)
    j = html.index('];', i)
    old = {}
    for m in re.finditer(r"\{name:'((?:[^'\\]|\\.)*)',seats:(\d+),keys:\[([^\]]*)\]\}", html[i:j]):
        keys = re.findall(r"'((?:[^'\\]|\\.)*)'", m.group(3))
        old[norm(m.group(1))] = (m.group(1), int(m.group(2)), keys)
    return i, j, old


def build_keys(name, existing, first_counts):
    keys = list(existing) if existing else []
    full = norm(name)
    if full not in keys:
        keys.insert(0, full)
    words = full.split(' ')
    if words[0].endswith('.') and len(words) > 1:          # "Ma. Josephine Carpio"
        alt = ' '.join(words[1:])
        if alt not in keys:
            keys.append(alt)
        words = words[1:]
    first = words[0]
    if len(words) > 1 and first_counts.get(first, 0) == 1 and first not in keys and not existing:
        keys.append(first)
    return keys


def main(xlsx, html_path):
    guests = read_sheet(xlsx)
    html = open(html_path, encoding='utf-8').read()
    i, j, old = read_site(html)

    first_counts = {}
    for n, _ in guests:
        w = norm(n).split(' ')
        f = w[1] if w[0].endswith('.') and len(w) > 1 else w[0]
        first_counts[f] = first_counts.get(f, 0) + 1

    lines, added, seat_changes = [], [], []
    for name, seats in guests:
        prev = old.get(norm(name))
        keys = build_keys(name, prev[2] if prev else None, first_counts)
        if not prev:
            added.append(name)
        elif prev[1] != seats:
            seat_changes.append(f'{name}: {prev[1]} -> {seats}')
        lines.append('  {name:%s,seats:%d,keys:[%s]}' % (js_str(name), seats, ','.join(js_str(k) for k in keys)))
    removed = [v[0] for k, v in old.items() if k not in {norm(n) for n, _ in guests}]

    # Replace the comment line directly above the list (if it is a comment) and the list body
    line_start = html.rfind('\n', 0, i) + 1
    prev_line_start = html.rfind('\n', 0, line_start - 1) + 1
    prev_line = html[prev_line_start:line_start - 1]
    head_start = prev_line_start if prev_line.strip().startswith('//') else line_start
    new_block = COMMENT + '\n' + START + '\n' + ',\n'.join(lines) + '\n'
    new_html = html[:head_start] + new_block + html[j:]

    changed = new_html != html
    print('CHANGED' if changed else 'NO CHANGES')
    print(f'Guests on site: {len(guests)}  (total seats {sum(s for _, s in guests)})')
    for label, items in (('Added', added), ('Removed', removed), ('Seat changes', seat_changes)):
        if items:
            print(f'{label}: ' + '; '.join(items))
    if changed:
        open(html_path, 'w', encoding='utf-8').write(new_html)


if __name__ == '__main__':
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2])
