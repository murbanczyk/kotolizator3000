"""Testy parsera stron schroniska na zapisanych fragmentach prawdziwych stron (tests/fixtures).

Gdy strona schroniska znowu zmieni wygląd, wystarczy zapisać nowe strony do tests/fixtures
i uruchomić `python -m pytest`, żeby zobaczyć, co przestało działać.
"""
import os

import pytest
from bs4 import BeautifulSoup

from analizatorkotow import scraper as S

FIXTURES = os.path.join(os.path.dirname(__file__), 'fixtures')


def soup(name):
    with open(os.path.join(FIXTURES, name), encoding='utf-8') as f:
        return BeautifulSoup(f.read(), 'lxml')


def test_parse_cards():
    cards = S.PaluchScraper.parse_cards(soup('lista_koty.html'))
    assert [c['number'] + '/' + c['year'] for c in cards] == ['0958/26', '0902/26', '0387/26', '0176/25']
    loki = cards[0]
    assert loki['name'] == 'Loki'
    assert loki['link'] == 'https://napaluchu.waw.pl/animal/0958-26k/'
    assert loki['breed'] == 'EUROPEJSKA'
    assert loki['age'] == '2 miesiące'
    assert loki['image'].startswith('https://images.napaluchu.waw.pl/')
    assert not loki['urgent']


def test_card_without_name():
    cards = S.PaluchScraper.parse_cards(soup('lista_koty.html'))
    assert cards[1]['name'] == ''
    assert cards[1]['number'] == '0902'


def test_urgent_badge():
    cards = S.PaluchScraper.parse_cards(soup('lista_koty.html'))
    assert cards[2]['urgent']


def test_parse_details():
    d = S.PaluchScraper.parse_details(soup('zwierze.html'))
    assert d['sex'] == 'samiec'
    assert d['date'] == '29 października 2026'
    assert d['location'] == 'Brak danych'
    assert d['photos'] == 3
    assert d['desc'].startswith('Loki to malutki kotek')
    assert 'Zapraszamy do' not in d['desc']
    assert d['fields']['Identyfikator'] == '0958/26'


class FakeResponse:
    def __init__(self, text='', content=b'', status=200):
        self.text, self.content, self.status_code = text, content, status

    def raise_for_status(self):
        if self.status_code >= 400:
            raise Exception('HTTP %d' % self.status_code)


def make_scraper(tmp_path, pages):
    calls = []
    sc = S.PaluchScraper(str(tmp_path / 'cache.kotolizer'), delay=(0, 0))

    def get(url, headers=None, timeout=None):
        calls.append(url)
        for key, response in pages.items():
            if key in url:
                return response
        return FakeResponse('<html></html>')
    sc.session.get = get
    return sc, calls


def test_get_animals_and_cache(tmp_path):
    with open(os.path.join(FIXTURES, 'lista_koty.html'), encoding='utf-8') as f:
        lista = f.read()
    with open(os.path.join(FIXTURES, 'zwierze.html'), encoding='utf-8') as f:
        zwierze = f.read()
    pages = {'koty-do-adopcji': FakeResponse(lista), '/animal/': FakeResponse(zwierze)}
    sc, calls = make_scraper(tmp_path, pages)
    animals = sc.get_animals('koty')
    assert len(animals) == 4
    for a in animals:
        details, new = sc.get_details(a['link'])
        assert new and details['sex'] == 'samiec'
    sc.save_cache()
    assert len(calls) == 1 + 4

    # drugie uruchomienie: szczegóły z pamięci podręcznej, tylko lista jest pobierana
    sc2, calls2 = make_scraper(tmp_path, pages)
    sc2.get_animals('koty')
    for a in animals:
        details, new = sc2.get_details(a['link'])
        assert not new and details['date'] == '29 października 2026'
    assert len(calls2) == 1


def test_empty_list_is_an_error(tmp_path):
    sc, _ = make_scraper(tmp_path, {})
    with pytest.raises(ValueError):
        sc.get_animals('koty')


def test_blocked_stops_downloading(tmp_path):
    sc, calls = make_scraper(tmp_path, {'/animal/': FakeResponse(status=403)})
    details, new = sc.get_details('https://napaluchu.waw.pl/animal/0001-26k/')
    assert details is None and not new and sc.blocked
    sc.get_details('https://napaluchu.waw.pl/animal/0002-26k/')
    assert len(calls) == 1
