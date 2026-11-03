"""Operacje na zestawieniach zwierząt, niezależne od interfejsu (łatwe do testowania)."""
import re
from datetime import date

import pandas as pd

COLUMNS = ['Nr', 'ID', 'Imię', 'Wiek', 'Przyjęcie', 'Foto', 'Ogłoszenie']

MONTHS = {'stycznia': 1, 'lutego': 2, 'marca': 3, 'kwietnia': 4, 'maja': 5, 'czerwca': 6, 'lipca': 7,
          'sierpnia': 8, 'września': 9, 'października': 10, 'listopada': 11, 'grudnia': 12}

# miejsca rozpoznawane w ostatniej linii kolumny "Przyjęcie"
MIEJSCA = ['Azyl', 'Ochota', 'właściciela']

SUMMARY_COLUMNS = ['Nr', 'ID', 'Imię', 'Płeć', 'Wiek', 'Przyjęto', 'Dni w schronisku', 'Pilne', 'Miejsce', 'Status']


# --- numery zwierząt ---

def animal_id(number, year):
    """Pełny numer zwierzęcia, np. '0958/26' (numery powtarzają się w kolejnych latach)."""
    return '%s/%s' % (number, year) if year else str(number)


def id_number(value):
    m = re.match(r'\s*(\d+)', str(value))
    return int(m.group(1)) if m else 0


def id_year(value):
    m = re.search(r'/(\d+)', str(value))
    return int(m.group(1)) if m else -1


def has_full_ids(frame):
    return len(frame) > 0 and all('/' in str(v) for v in frame['ID'])


def comparedf(one, two):
    """Zwierzęta z zestawienia `one`, których nie ma w `two`.

    Nowe zestawienia mają pełne numery ('0958/26'). Starsze zapisy mają tylko liczbę (958),
    więc gdy któreś zestawienie jest w starym formacie, porównujemy same liczby.
    """
    full = has_full_ids(one) and has_full_ids(two)
    key = (lambda f: f['ID'].map(str)) if full else (lambda f: f['ID'].map(id_number))
    return one[~key(one).isin(set(key(two)))]


def sort_frame(frame):
    """Sortuje zestawienie po roku i numerze zwierzęcia."""
    order = sorted(range(len(frame)), key=lambda i: (id_year(frame['ID'].iloc[i]), id_number(frame['ID'].iloc[i])))
    return frame.iloc[order]


# --- daty i wiek ---

def parse_date(text):
    """Data z tekstu, np. '29 października 2026', '29.10.2026' albo '2026-10-29'; None, gdy się nie da."""
    text = str(text).strip().lower()
    m = re.search(r'(\d{1,2})\s+([a-ząćęłńóśźż]+)\s+(\d{4})', text)
    try:
        if m and m.group(2) in MONTHS:
            return date(int(m.group(3)), MONTHS[m.group(2)], int(m.group(1)))
        m = re.search(r'(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})', text)
        if m:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        m = re.search(r'(\d{4})-(\d{1,2})-(\d{1,2})', text)
        if m:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        pass
    return None


def age_months(wiek):
    """Wiek w miesiącach z tekstu '3 miesiące', '1 rok', '5 lat'; None, gdy brak liczby."""
    m = re.match(r'\s*(\d+)', str(wiek))
    if not m:
        return None
    n = int(m.group(1))
    return n * 12 if ('rok' in wiek or 'lat' in wiek) else n


def admission_date(przyjecie):
    """Data przyjęcia z kolumny 'Przyjęcie' (pierwsza linia)."""
    return parse_date(str(przyjecie).split('\n')[0])


def miejsce(przyjecie):
    """Miejsce pobytu z ostatniej linii kolumny 'Przyjęcie' albo 'brak informacji'."""
    last = str(przyjecie).split('\n')[-1]
    if 'Ochota' in last:
        return 'Ochota na Kota'
    return last if any(m in last for m in MIEJSCA) else 'brak informacji'


def status(przyjecie):
    """Odznaki stanu zwierzęcia i informacja o znalezieniu z kolumny 'Przyjęcie'."""
    parts = []
    for line in str(przyjecie).split('\n'):
        if line.startswith('Status: '):
            parts.append(line[len('Status: '):])
        elif line.startswith('Ostatnio znaleziony'):
            parts.append(line)
    return '; '.join(parts)


# --- budowanie zestawień ---

def build_row(nr, animal, details, image):
    """Wiersz zestawienia z danych karty z listy i strony zwierzęcia."""
    details = details or {}
    fields = details.get('fields', {})
    badges = details.get('badges', [])
    location = details.get('location', '')
    if location == 'Brak danych':
        location = ''
    ochota = (animal.get('ochota') or 'Ochota' in location
              or any('ochota na kota' in b.lower() for b in badges))
    urgent = animal['urgent'] or any('pilnej' in b.lower() for b in badges)
    przyjecie = details.get('date', '')
    przyjecie += '\n\nZdjęć:' + str(details.get('photos', ''))
    przyjecie += '\n\nRasa: ' + (animal['breed'] or fields.get('Typ rasy', ''))
    status_badges = [b for b in badges if 'ochota' not in b.lower()]
    if status_badges:
        przyjecie += '\n\nStatus: ' + ', '.join(status_badges)
    if animal.get('found'):
        przyjecie += '\n\nOstatnio znaleziony: ' + ', '.join(x for x in (animal['found'], animal.get('found_where')) if x)
    if urgent:
        przyjecie += '\n\nPILNE'
    # ostatnia linia jest używana w statystyce miejsc
    if ochota:
        przyjecie += '\n\nOchota na Kota' + (' – ' + location if location and 'Ochota' not in location else '')
    elif location:
        przyjecie += '\n\n' + location
    return {'Nr': nr, 'ID': animal_id(animal['number'], animal['year']),
            'Imię': animal['name'] + '\n' + details.get('sex', animal['sex']),
            'Wiek': animal['age'] or fields.get('Wiek', ''), 'Przyjęcie': przyjecie, 'Foto': image,
            'Ogłoszenie': details.get('desc', '')}


def summary(frame, today=None):
    """Tabela do przeglądania: imię, płeć, wiek, data przyjęcia, liczba dni w schronisku, miejsce."""
    today = today or date.today()
    rows = []
    for nr, (_, r) in enumerate(frame.iterrows(), start=1):
        name = str(r['Imię']).split('\n')
        przyjeto = admission_date(r['Przyjęcie'])
        rows.append({
            'Nr': nr,
            'ID': str(r['ID']),
            'Imię': name[0],
            'Płeć': name[1] if len(name) > 1 else '',
            'Wiek': str(r['Wiek']),
            'Przyjęto': przyjeto.strftime('%d.%m.%Y') if przyjeto else '',
            'Dni w schronisku': (today - przyjeto).days if przyjeto else None,
            'Pilne': 'tak' if 'PILNE' in str(r['Przyjęcie']) else '',
            'Miejsce': miejsce(r['Przyjęcie']),
            'Status': status(r['Przyjęcie']),
            '_wiek': age_months(r['Wiek']),
            '_foto': r['Foto'] if hasattr(r['Foto'], 'size') else None,  # miniatura do podglądu
        })
    return pd.DataFrame(rows, columns=SUMMARY_COLUMNS + ['_wiek', '_foto'])


def trend(days):
    """Liczba zwierząt oraz nowe i zniknięte w kolejnych dniach.

    `days` to lista (dzień jako date, zestawienie) posortowana od najstarszego.
    """
    rows = []
    previous = None
    for day, frame in days:
        new = len(comparedf(frame, previous)) if previous is not None else None
        gone = len(comparedf(previous, frame)) if previous is not None else None
        rows.append({'Dzień': day, 'Liczba': len(frame), 'Nowe': new, 'Zniknięte': gone})
        previous = frame
    return pd.DataFrame(rows, columns=['Dzień', 'Liczba', 'Nowe', 'Zniknięte'])
