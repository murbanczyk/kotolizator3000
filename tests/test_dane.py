from datetime import date

import pandas as pd

from analizatorkotow import dane


def frame(ids, **extra):
    rows = []
    for i, x in enumerate(ids):
        rows.append({'Nr': i + 1, 'ID': x, 'Imię': 'Kot%d\nsamica' % i, 'Wiek': '2 lata',
                     'Przyjęcie': extra.get('przyjecie', '1 października 2026\n\nZdjęć:1'),
                     'Foto': None, 'Ogłoszenie': ''})
    return pd.DataFrame(rows, columns=dane.COLUMNS)


def test_same_number_different_year_is_different_animal():
    old = frame(['0958/25', '0100/26'])
    new = frame(['0958/26', '0100/26'])
    assert list(dane.comparedf(new, old)['ID']) == ['0958/26']
    assert list(dane.comparedf(old, new)['ID']) == ['0958/25']


def test_compare_with_old_format_uses_numbers():
    old = frame([958, 100])  # zestawienia zapisane starszą wersją programu
    new = frame(['0958/26', '0101/26'])
    assert list(dane.comparedf(new, old)['ID']) == ['0101/26']
    assert list(dane.comparedf(old, new)['ID']) == [100]


def test_sort_by_year_then_number():
    f = frame(['0010/26', '0958/25', '0002/26'])
    assert list(dane.sort_frame(f)['ID']) == ['0958/25', '0002/26', '0010/26']


def test_parse_date():
    assert dane.parse_date('29 października 2026') == date(2026, 10, 29)
    assert dane.parse_date('1 Stycznia 2025') == date(2025, 1, 1)
    assert dane.parse_date('03.11.2026') == date(2026, 11, 3)
    assert dane.parse_date('2026-11-03') == date(2026, 11, 3)
    assert dane.parse_date('') is None
    assert dane.parse_date('31 lutego 2026') is None


def test_age_months():
    assert dane.age_months('2 miesiące') == 2
    assert dane.age_months('1 rok') == 12
    assert dane.age_months('5 lat') == 60
    assert dane.age_months('brak') is None


def test_miejsce():
    assert dane.miejsce('data\n\nPILNE\n\nOchota na Kota') == 'Ochota na Kota'
    assert dane.miejsce('data\n\nRasa: EUROPEJSKA') == 'brak informacji'


def test_build_row():
    animal = {'number': '0958', 'year': '26', 'name': 'Loki', 'sex': '', 'breed': 'EUROPEJSKA',
              'age': '2 miesiące', 'urgent': True}
    details = {'sex': 'samiec', 'date': '29 października 2026', 'location': 'Ochota', 'photos': 3, 'desc': 'Opis'}
    row = dane.build_row(1, animal, details, None)
    assert row['ID'] == '0958/26'
    assert row['Imię'] == 'Loki\nsamiec'
    assert 'PILNE' in row['Przyjęcie']
    assert row['Przyjęcie'].endswith('Ochota (Ochota na Kota)')
    assert dane.miejsce(row['Przyjęcie']).startswith('Ochota')
    # bez strony zwierzęcia (np. strona zablokowała pobieranie)
    row = dane.build_row(2, animal, None, None)
    assert row['Imię'] == 'Loki\n' and row['Ogłoszenie'] == ''


def test_summary_days_in_shelter():
    f = frame(['0958/26'], przyjecie='29 października 2026\n\nZdjęć:3\n\nPILNE')
    s = dane.summary(f, today=date(2026, 11, 2))
    r = s.iloc[0]
    assert r['Imię'] == 'Kot0' and r['Płeć'] == 'samica'
    assert r['Przyjęto'] == '29.10.2026'
    assert r['Dni w schronisku'] == 4
    assert r['Pilne'] == 'tak'
    assert r['_wiek'] == 24


def test_summary_without_date():
    s = dane.summary(frame([958], przyjecie='\n\nZdjęć:'))
    assert pd.isna(s.iloc[0]['Dni w schronisku'])


def test_trend():
    days = [(date(2026, 10, 1), frame(['1/26', '2/26'])),
            (date(2026, 10, 8), frame(['2/26', '3/26', '4/26']))]
    t = dane.trend(days)
    assert list(t['Liczba']) == [2, 3]
    assert t.iloc[1]['Nowe'] == 2 and t.iloc[1]['Zniknięte'] == 1
    assert pd.isna(t.iloc[0]['Nowe'])
