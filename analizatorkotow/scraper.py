import bz2
import pickle
import random
import re
import threading
import time
from concurrent.futures import Future, ThreadPoolExecutor
from io import BytesIO
from urllib.parse import urljoin

import requests
from PIL import Image
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = 'https://napaluchu.waw.pl/'

# listy zwierzat do adopcji (cala lista miesci sie na jednej stronie)
LIST_URLS = {
    'koty': BASE_URL + 'zwierzeta/koty-do-adopcji/',
    'psy': BASE_URL + 'zwierzeta/psy-do-adopcji/',
}

# ostatnio znalezione zwierzeta danego gatunku (moga jeszcze trafic do wlascicieli)
FOUND_URLS = {
    'koty': BASE_URL + 'zwierzeta/ostatnio-znalezione/gatunek-kot/',
    'psy': BASE_URL + 'zwierzeta/ostatnio-znalezione/gatunek-pies/',
}

# strona punktu adopcyjnego Ochota na Kota z karta kazdego kota, ktory tam przebywa
OCHOTA_URL = BASE_URL + 'o-schronisku/ochota-na-kota/'

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/124.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
    'Accept-Language': 'pl-PL,pl;q=0.9,en-US;q=0.6,en;q=0.5',
    'Connection': 'keep-alive',
}


class BlockedError(Exception):
    """Strona odmawia dostepu (403/429) mimo ponawiania."""


class PaluchScraper:
    """Pobiera dane ze strony schroniska oszczedzajac polaczenia.

    Imie, numer, wiek, zdjecie i "pilne" sa na stronie z lista zwierzat (1 zapytanie).
    Plec, data przyjecia, lokalizacja, opis i liczba zdjec sa tylko na stronie
    zwierzecia - ta jest pobierana raz i trzymana w pamieci podrecznej przez
    details_days dni. Miniatury zdjec sa pobierane tylko raz.

    - jedna sesja HTTP (keep-alive) do stron i osobna do zdjec, naglowki przegladarki i Referer,
    - odstepy z losowym jitterem i ponawianie z backoffem
      (z respektowaniem Retry-After) przy 429/5xx,
    - zdjecia sa na osobnym serwerze (images.napaluchu.waw.pl), wiec pobiera je
      jeden watek w tle z krotszymi odstepami, rownolegle ze stronami zwierzat;
      na kazdy serwer idzie naraz najwyzej jedno zapytanie,
    - po 403/429 nic wiecej nie jest pobierane z danego serwera; uzywane sa dane
      z pamieci podrecznej (nawet starsze), a brakujace zdjecia zastepuje pusty obrazek.
    """

    DETAIL_LABELS = ('Typ rasy', 'Płeć', 'Data przyjęcia', 'Wiek', 'Identyfikator', 'Grupa Zwierząt',
                     'Waga', 'Lokalizacja (boks)', 'Kastracja')
    DESC_END = ('Zapraszamy do', 'Planujesz adopcję', 'Poradnik:', 'Zapraszamy do zapoznania się z ankietą')

    def __init__(self, cache_path, details_days=7, delay=(1.0, 2.5), image_delay=(0.1, 0.4)):
        self.cache_path = str(cache_path)
        self.details_max_age = details_days * 24 * 3600
        self.delay = delay
        self.image_delay = image_delay
        self._last_request = 0.0
        self._last_image_request = 0.0
        self.blocked = False         # strona schroniska blokuje zapytania
        self.images_blocked = False  # serwer zdjec blokuje zapytania
        self.session = self._new_session()
        self.image_session = self._new_session()
        self._lock = threading.Lock()
        self._image_worker = None
        cache = self._load_cache()
        self.images = cache.get('images', {})    # url zdjecia -> miniatura w PNG
        self.details = cache.get('details', {})  # link do zwierzecia -> dane ze strony zwierzecia

    @staticmethod
    def _new_session():
        session = requests.Session()
        session.headers.update(HEADERS)
        retry = Retry(total=4, backoff_factor=2, status_forcelist=(429, 500, 502, 503, 504),
                      allowed_methods=('GET',), respect_retry_after_header=True,
                      raise_on_status=False)
        adapter = HTTPAdapter(max_retries=retry, pool_connections=1, pool_maxsize=1)
        session.mount('https://', adapter)
        session.mount('http://', adapter)
        return session

    # --- pamiec podreczna ---

    def _load_cache(self):
        try:
            with bz2.open(self.cache_path, 'rb') as f:
                return pickle.load(f)
        except Exception:
            return {}

    def save_cache(self):
        with self._lock:  # zdjecia moga byc wlasnie dopisywane przez watek w tle
            data = {'images': dict(self.images), 'details': dict(self.details)}
        try:
            with bz2.open(self.cache_path, 'wb') as f:
                pickle.dump(data, f)
        except Exception as e:
            print('Nie udalo sie zapisac cache: ' + str(e))

    # --- HTTP ---

    def _get(self, url, referer=None):
        """Strona schroniska (z dluzszymi odstepami)."""
        wait = random.uniform(*self.delay) - (time.time() - self._last_request)
        if wait > 0:
            time.sleep(wait)
        headers = {'Referer': referer} if referer else None
        try:
            response = self.session.get(url, headers=headers, timeout=30)
        finally:
            self._last_request = time.time()
        return self._check(response, url)

    def _get_image(self, url, referer=None):
        """Zdjecie z serwera zdjec (krotsze odstepy, osobna sesja)."""
        wait = random.uniform(*self.image_delay) - (time.time() - self._last_image_request)
        if wait > 0:
            time.sleep(wait)
        headers = {'Referer': referer} if referer else None
        try:
            response = self.image_session.get(url, headers=headers, timeout=30)
        finally:
            self._last_image_request = time.time()
        return self._check(response, url)

    @staticmethod
    def _check(response, url):
        if response.status_code in (403, 429):
            raise BlockedError('%d dla %s' % (response.status_code, url))
        response.raise_for_status()
        return response

    def getsoup(self, url, referer=None):
        return BeautifulSoup(self._get(url, referer).text, 'lxml')

    # --- listy zwierzat ---

    @staticmethod
    def parse_cards(soup):
        """Zwraca liste slownikow z danymi zwierzat z kart na stronie listy."""
        animals = []
        for card in soup.select('li.post-card'):
            title = card.find('h3')
            link = title.find('a', href=True) if title else None
            if link is None or '/animal/' not in link['href']:
                continue
            text = link.get_text(' ', strip=True)
            m = re.match(r'(?:(.*?)\s*[\u2013\u2014-]\s*)?(\d+)/(\d+)\s*$', text)
            if m:
                name, number, year = m.group(1) or '', m.group(2), m.group(3)
            else:
                name = text
                m = re.search(r'/animal/(\d+)-(\d+)', link['href'])
                number, year = (m.group(1), m.group(2)) if m else ('0', '')
            # pola karty: napis zakończony dwukropkiem to etykieta, następny napis to wartość
            # (np. 'Rasa:' 'EUROPEJSKA', 'Wiek:' '2 miesiące', 'Znaleziono:' '2 listopada 2026', 'Gdzie:' ...)
            fields = {}
            strings = list(card.stripped_strings)
            for i, t in enumerate(strings[:-1]):
                if t.endswith(':') and t[:-1].strip() not in fields:
                    fields[t[:-1].strip()] = strings[i + 1]
            img = card.find('img')
            imgurl = (img.get('data-src') or img.get('src')) if img is not None else None
            if imgurl and imgurl.startswith('data:'):
                imgurl = img.get('data-orig-src') or img.get('data-lazy-src')
            animals.append({
                'link': link['href'],
                'name': name,
                'number': number,
                'year': year,
                'breed': fields.get('Rasa', ''),
                'age': fields.get('Wiek', ''),
                'image': imgurl,
                'urgent': 'pilnej adopcji' in card.get_text(' ', strip=True).lower(),
                'sex': '',
                'found': fields.get('Znaleziono', ''),
                'found_where': fields.get('Gdzie', ''),
                'ochota': False,
            })
        return animals

    def get_list(self, url, referer=None, max_pages=20):
        """Pobiera liste zwierzat; idzie za linkiem do nastepnej strony, jesli taki jest."""
        animals = []
        for _ in range(max_pages):
            soup = self.getsoup(url, referer=referer)
            animals += self.parse_cards(soup)
            nxt = soup.select_one('.pagination a.pagination-next[href], .pagination a[rel=next][href]')
            if nxt is None:
                break
            referer, url = url, urljoin(url, nxt['href'])
        return animals

    def get_animals(self, species):
        """Zwraca liste zwierzat danego gatunku ('koty' lub 'psy').

        Lista do adopcji jest wymagana. Do niej dochodza ostatnio znalezione zwierzeta
        (pole 'found'), a u kotow oznaczenie, ktore sa w punkcie Ochota na Kota (pole 'ochota').
        Te dwie strony sa dodatkowe - gdy sie nie wczytaja, lista i tak jest zwracana.
        """
        base = LIST_URLS[species]
        animals = self.get_list(base, referer=BASE_URL)
        if not animals:
            raise ValueError('Na stronie ' + base + ' nie znaleziono żadnych zwierząt. '
                             'Strona mogła zablokować pobieranie albo zmienić wygląd.')
        bylink = dict((a['link'], a) for a in animals)
        for a in self._optional_list(FOUND_URLS[species], base, 'ostatnio znalezionych'):
            bylink.setdefault(a['link'], a)
        if species == 'koty':
            for a in self._optional_list(OCHOTA_URL, base, 'Ochoty na Kota'):
                if a['link'] in bylink:
                    bylink[a['link']]['ochota'] = True
        return list(bylink.values())

    def _optional_list(self, url, referer, what):
        if self.blocked:
            return []
        try:
            return self.get_list(url, referer=referer)
        except BlockedError as e:
            print('Strona blokuje zapytania (' + str(e) + '), pomijam liste ' + what)
            self.blocked = True
        except Exception as e:
            print('Nie udało się pobrać listy ' + what + ': ' + str(e))
        return []

    # --- strona zwierzecia ---

    @classmethod
    def parse_details(cls, soup):
        texts = [d.get_text(' ', strip=True) for d in soup.select('div.fusion-text')]
        fields = {}
        for i, t in enumerate(texts[:-1]):
            if t in cls.DETAIL_LABELS and t not in fields:
                fields[t] = texts[i + 1]
        desc = ''
        for h in soup.find_all(['h2', 'h3', 'h4']):
            if h.get_text(strip=True) == 'Informacje':
                body = h.find_next('div', class_='fusion-text')
                if body is not None:
                    desc = body.get_text('\n', strip=True)
                break
        for marker in cls.DESC_END:
            desc = desc.split(marker)[0]
        # odznaki stanu zwierzecia to przyciski prowadzace do jego wlasnej strony,
        # np. 'Do adopcji', 'Nie do adopcji', 'Ochota na kota', 'Do pilnej adopcji'
        badges = []
        for span in soup.select('a[href*="/animal/"] span.fusion-button-text'):
            text = span.get_text(' ', strip=True)
            if text and text not in badges:
                badges.append(text)
        return {
            'sex': fields.get('Płeć', '').lower(),
            'date': fields.get('Data przyjęcia', ''),
            'location': fields.get('Lokalizacja (boks)', ''),
            'desc': desc.strip(),
            'photos': len(soup.select('a.fusion-lightbox[data-rel^="iLightbox[gallery"]')),
            'fields': fields,
            'badges': badges,
        }

    def get_details(self, link, referer=None):
        """Dane ze strony zwierzecia lub None. Zwraca (dane, czy_pobrano)."""
        entry = self.details.get(link)
        fresh = entry is not None and time.time() - entry['time'] < self.details_max_age
        if fresh or self.blocked:
            return entry, False
        try:
            data = self.parse_details(self.getsoup(link, referer=referer))
        except BlockedError as e:
            print('Strona blokuje zapytania (' + str(e) + '), dalej używam tylko pamięci podręcznej')
            self.blocked = True
            return entry, False
        except Exception as e:
            print('Problem z ' + link + ': ' + str(e))
            return entry, False
        data['time'] = time.time()
        self.details[link] = data
        return data, True

    # --- zdjecia ---

    def get_thumbnail(self, url, referer=None, fixed_height=150):
        """Miniatura zdjecia; pobierana tylko raz, potem brana z pamieci podrecznej."""
        with self._lock:
            data = self.images.get(url) if url else None
        if data is None and url and not self.images_blocked:
            try:
                image = Image.open(BytesIO(self._get_image(url, referer=referer).content))
                height_percent = (fixed_height / float(image.size[0]))
                width_size = int((float(image.size[1]) * float(height_percent)))
                image = image.convert('RGB').resize((fixed_height, width_size), Image.NEAREST)
                buf = BytesIO()
                image.save(buf, format='PNG')
                data = buf.getvalue()
                with self._lock:
                    self.images[url] = data
            except BlockedError as e:
                print('Serwer zdjęć blokuje zapytania (' + str(e) + '), pozostałe zdjęcia pomijam')
                self.images_blocked = True
            except Exception as e:
                print('Problem ze zdjęciem ' + str(url) + ': ' + str(e))
        if data is None:
            return Image.new('RGB', (fixed_height, fixed_height * 3 // 4), 'white')
        image = Image.open(BytesIO(data))
        image.load()
        return image

    def thumbnail_later(self, url, referer=None):
        """Zleca pobranie miniatury w tle i od razu zwraca Future z obrazkiem.

        Zdjecie z pamieci podrecznej (albo brak zdjecia) jest gotowe od razu, bez watku.
        """
        with self._lock:
            cached = not url or url in self.images
        if cached or self.images_blocked:
            done = Future()
            done.set_result(self.get_thumbnail(url, referer))
            return done
        if self._image_worker is None:
            self._image_worker = ThreadPoolExecutor(max_workers=1, thread_name_prefix='zdjecia')
        return self._image_worker.submit(self.get_thumbnail, url, referer)

    def close(self):
        """Czeka na zlecone zdjecia i zamyka watek w tle."""
        if self._image_worker is not None:
            self._image_worker.shutdown(wait=True)
            self._image_worker = None
